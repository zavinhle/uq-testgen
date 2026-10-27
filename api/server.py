"""FastAPI backend: runs the pipeline with live progress (Server-Sent Events) and serves results + UI."""
from __future__ import annotations

import asyncio
import glob
import json
import os
import queue
import sys
import threading

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from core import sim  # noqa: E402
from core.llm import MODELS  # noqa: E402
from core.oracle import TEMPLATES  # noqa: E402
from core.pipeline import RUNS, run_pipeline  # noqa: E402
from core.requirements import REQUIREMENTS  # noqa: E402
from core.sut import MUTANTS  # noqa: E402
from core.uq import annotate_run  # noqa: E402

app = FastAPI(title="UQ-TestGen")
GIFS = os.path.join(ROOT, "results", "gifs")
os.makedirs(GIFS, exist_ok=True)
_streams: dict[str, queue.Queue] = {}


class RunReq(BaseModel):
    rid: str
    model: str = MODELS[0]
    n: int = 8
    ground_truth: bool = True


@app.get("/api/meta")
def meta():
    return {"requirements": REQUIREMENTS, "models": MODELS, "mutants": list(MUTANTS), "templates": TEMPLATES}


@app.post("/api/runs")
def start_run(req: RunReq):
    if req.rid not in REQUIREMENTS:
        raise HTTPException(400, "unknown requirement")
    import time
    import uuid
    run_id = f"{req.rid}_{req.model}_{time.strftime('%Y%m%d-%H%M%S')}_{uuid.uuid4().hex[:4]}"
    q: queue.Queue = queue.Queue()
    _streams[run_id] = q

    def work():
        try:
            run_pipeline(req.rid, req.model, max(2, min(req.n, 20)), workers=5, run_id=run_id,
                         ground_truth_mode=req.ground_truth, emit=lambda ev, data: q.put((ev, data)))
        except Exception as ex:  # noqa: BLE001
            q.put(("error", {"message": repr(ex)}))
        finally:
            q.put(None)

    threading.Thread(target=work, daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/runs/{run_id}/events")
async def events(run_id: str):
    q = _streams.get(run_id)
    if q is None:
        raise HTTPException(404, "no live run")

    async def gen():
        while True:
            try:
                item = q.get_nowait()
            except queue.Empty:
                await asyncio.sleep(0.15)
                yield ": keep-alive\n\n"
                continue
            if item is None:
                _streams.pop(run_id, None)
                return
            ev, data = item
            yield f"event: {ev}\ndata: {json.dumps(data)}\n\n"

    return StreamingResponse(gen(), media_type="text/event-stream")


@app.get("/api/runs")
def list_runs():
    out = []
    for f in sorted(glob.glob(os.path.join(RUNS, "*.json")), key=os.path.getmtime, reverse=True):
        r = annotate_run(json.load(open(f, encoding="utf-8")))
        ss = r["samples"]
        out.append({"run_id": r["run_id"], "rid": r["rid"], "model": r["model"], "n": r["n"],
                    "correct": sum(1 for s in ss if (s.get("gt") or {}).get("correct")),
                    "accept": sum(s["triage"] == "accept" for s in ss), "review": sum(s["triage"] == "review" for s in ss),
                    "reject": sum(s["triage"] == "reject" for s in ss), "batch_uq": r.get("batch_uq"),
                    "n_groups": len((r.get("batch_uq") or {}).get("groups") or []),
                    "wall_seconds": r.get("wall_seconds")})
    return out


@app.get("/api/runs/{run_id}")
def get_run(run_id: str):
    p = os.path.join(RUNS, f"{run_id}.json")
    if not os.path.exists(p):
        raise HTTPException(404)
    return JSONResponse(annotate_run(json.load(open(p, encoding="utf-8"))))


@app.get("/api/runs/{run_id}/gif/{idx}")
def gif(run_id: str, idx: int, version: str = "correct"):
    if version != "correct" and version not in MUTANTS:
        raise HTTPException(400)
    out = os.path.join(GIFS, f"{run_id}_{idx}_{version}.gif")
    if not os.path.exists(out):
        r = json.load(open(os.path.join(RUNS, f"{run_id}.json"), encoding="utf-8"))
        s = next(s for s in r["samples"] if s["idx"] == idx)
        if s["validation_errors"]:
            raise HTTPException(400, "invalid scenario")
        sim.render_gif(s["scenario"], version, out, title=f"{r['rid']} test #{idx} on {version}")
    return FileResponse(out, media_type="image/gif")


@app.get("/api/summary")
def summary():
    p = os.path.join(ROOT, "results", "summary.json")
    if not os.path.exists(p):
        raise HTTPException(404, "run analysis/analyze.py first")
    return JSONResponse(json.load(open(p, encoding="utf-8")))


app.mount("/figures", StaticFiles(directory=os.path.join(ROOT, "results", "figures"), check_dir=False), name="figures")
app.mount("/", StaticFiles(directory=os.path.join(ROOT, "ui"), html=True), name="ui")
