"""End-to-end pipeline: requirement -> sampled tests -> validation -> execution -> UQ -> triage -> ground truth."""
from __future__ import annotations

import json
import os
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np

from . import sim
from .evaluate import ground_truth
from .generate import build_prompt, sample_one
from .requirements import REQUIREMENTS
from .scenario import validate
from .sut import VERSIONS
from .uq import quantify, triage

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNS = os.path.join(ROOT, "results", "runs")


def _jsonable(o):
    if isinstance(o, dict):
        return {k: _jsonable(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_jsonable(v) for v in o]
    if isinstance(o, (np.bool_,)):
        return bool(o)
    if isinstance(o, (np.floating, float)):
        return None if not np.isfinite(o) else float(o)
    if isinstance(o, np.integer):
        return int(o)
    return o


def run_pipeline(rid: str, model: str, n: int, workers: int = 5, emit=None, run_id: str | None = None,
                 ground_truth_mode: bool = True) -> dict:
    emit = emit or (lambda ev, data: None)
    run_id = run_id or f"{rid}_{model}_{time.strftime('%Y%m%d-%H%M%S')}_{uuid.uuid4().hex[:4]}"
    req = REQUIREMENTS[rid]
    t_start = time.perf_counter()
    emit("stage", {"stage": "generate", "status": "running", "prompt": build_prompt(rid, req["text"])})

    # 1) generate N samples (independent LLM sessions)
    samples = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(sample_one, rid, req["text"], model, i) for i in range(n)]
        for f in as_completed(futs):
            s = f.result()
            s["validation_errors"] = validate(s["scenario"]) if s["scenario"] is not None else ["unparseable reply"]
            samples.append(s)
            emit("sample", _jsonable({k: s[k] for k in ("idx", "latency_s", "tokens", "scenario", "validation_errors", "llm_error")}))
    samples.sort(key=lambda s: s["idx"])
    emit("stage", {"stage": "generate", "status": "done"})

    # 2) rule-based validation gate
    valid = [s for s in samples if not s["validation_errors"]]
    emit("stage", {"stage": "validate", "status": "done", "valid": len(valid), "total": len(samples)})

    # 3) execute on the SUT under test (the reference controller)
    emit("stage", {"stage": "execute", "status": "running"})
    traces_sut = {}
    for s in valid:
        traces_sut[s["idx"]] = sim.run(s["scenario"], "correct")
        emit("executed", {"idx": s["idx"], "sim_seconds": traces_sut[s["idx"]]["sim_seconds"]})
    emit("stage", {"stage": "execute", "status": "done"})

    # 4) uncertainty quantification (no ground truth used)
    uq = quantify(valid, traces_sut) if valid else {"per_sample": {}, "batch": {}}
    for s in samples:
        u = uq["per_sample"].get(s["idx"])
        s["uq"] = u
        s["triage"] = triage(u["U_comb"] if u else None, not s["validation_errors"])
    emit("uq", _jsonable({"batch": uq["batch"], "per_sample": {i: {k: v for k, v in u.items() if k != "features"} | {"triage": next(x["triage"] for x in samples if x["idx"] == i)}
                                                                for i, u in uq["per_sample"].items()}}))
    emit("stage", {"stage": "uq", "status": "done"})

    # 5) ground truth via fault injection (research mode)
    if ground_truth_mode:
        emit("stage", {"stage": "ground_truth", "status": "running"})
        for s in valid:
            traces = {"correct": traces_sut[s["idx"]], **{v: sim.run(s["scenario"], v) for v in VERSIONS if v != "correct"}}
            s["gt"] = ground_truth(rid, s["scenario"], traces)
            emit("gt", _jsonable({"idx": s["idx"], **{k: s["gt"][k] for k in ("label", "correct", "ms_gen", "ms_ref", "oracle_agreement", "killed_gen")}}))
        emit("stage", {"stage": "ground_truth", "status": "done"})

    result = _jsonable({"run_id": run_id, "rid": rid, "requirement": req, "model": model, "n": n,
                        "wall_seconds": time.perf_counter() - t_start, "batch_uq": uq["batch"],
                        "samples": [{k: v for k, v in s.items() if k != "raw"} | {"raw": s.get("raw", "")[:4000]} for s in samples]})
    os.makedirs(RUNS, exist_ok=True)
    with open(os.path.join(RUNS, f"{run_id}.json"), "w", encoding="utf-8") as fh:
        json.dump(result, fh, indent=1)
    emit("done", {"run_id": run_id})
    return result
