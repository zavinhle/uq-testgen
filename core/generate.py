"""Requirement -> N independently sampled test artefacts (scenario + oracle + self-reported confidence)."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor

from .llm import codex_generate, extract_json
from .oracle import TEMPLATES
from .scenario import validate

PROMPT = """You are a test engineer for an Adaptive Cruise Control (ACC) with emergency braking on a highway.
Generate ONE simulation-based system test for the requirement below. Do not run any commands; answer only.

REQUIREMENT {rid}: {text}

SIMULATION ENVIRONMENT
- Straight 3-lane highway, lanes 0,1,2 (adjacent lanes differ by 1), lane width 4 m, vehicles 5 m long.
- Ego vehicle (with ACC) starts at x=0 in `ego.lane` with `speed` [m/s] and driver `set_speed` [m/s].
- Up to 4 scripted vehicles ("npcs"). `rel_x` = initial longitudinal position of the npc's REAR relative to the
  ego's FRONT (+ = ahead) [m]. Scripted events (times in s from start):
    {{"t": 5, "type": "set_speed", "target_speed": 0, "rate": 4}}   change speed at `rate` m/s^2
    {{"t": 3, "type": "lane_change", "target_lane": 1}}              move to an adjacent lane (~2 s)
- duration: 5..60 s. Speeds 0..40 m/s.

ORACLE: list of 1..3 clauses that must ALL hold; each clause is one template:
{templates}
Signals: lead_gap = bumper-to-bumper gap to the vehicle ahead in the ego lane; headway = lead_gap / ego_speed.

Respond with ONLY a JSON object:
{{"rationale": "<one sentence>", "duration": <s>, "ego": {{"lane": <int>, "speed": <m/s>, "set_speed": <m/s>}},
 "npcs": [{{"id": "<name>", "lane": <int>, "rel_x": <m>, "speed": <m/s>, "events": [...]}}],
 "oracle": [{{"template": "<name>", "params": {{...}}}}],
 "confidence": <0..1, your probability that this test correctly and effectively checks the requirement>}}
"""


def build_prompt(rid: str, text: str) -> str:
    tpl = "\n".join(f"  - {k}: {v}" for k, v in TEMPLATES.items())
    return PROMPT.format(rid=rid, text=text, templates=tpl)


def sample_one(rid: str, text: str, model: str, idx: int) -> dict:
    r = codex_generate(build_prompt(rid, text), model=model)
    rec = {"rid": rid, "model": model, "idx": idx, "latency_s": r.latency_s, "tokens": r.tokens,
           "raw": r.text, "llm_error": r.error, "scenario": None, "parse_error": None}
    if r.error:
        return rec
    try:
        rec["scenario"] = extract_json(r.text)
    except Exception as ex:  # noqa: BLE001
        rec["parse_error"] = repr(ex)[:200]
    return rec


def sample_many(rid: str, text: str, model: str, n: int, workers: int = 5, on_sample=None) -> list[dict]:
    out: list[dict] = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(sample_one, rid, text, model, i) for i in range(n)]
        for f in futs:
            rec = f.result()
            rec["validation_errors"] = validate(rec["scenario"]) if rec["scenario"] is not None else ["unparseable"]
            out.append(rec)
            if on_sample:
                on_sample(rec)
    return out
