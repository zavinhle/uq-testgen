"""Rule-based validation of LLM-generated test artefacts (schema + physical feasibility).

This is the 'v1' trust mechanism: it catches malformed and physically impossible outputs,
but cannot tell whether a well-formed test is relevant, has a correct oracle, or is effective.
"""
from __future__ import annotations

from .oracle import TEMPLATES

PARAMS = {"no_collision": [], "min_gap": ["d"], "time_headway": ["h", "t_from"], "speed_limit": ["vmax"], "stop_gap": ["d"], "max_decel": ["a"]}


def validate(sc) -> list[str]:
    errs: list[str] = []
    if not isinstance(sc, dict):
        return ["not a JSON object"]
    for key in ("duration", "ego", "npcs", "oracle"):
        if key not in sc:
            errs.append(f"missing '{key}'")
    if errs:
        return errs
    try:
        dur = float(sc["duration"])
        if not 5 <= dur <= 60:
            errs.append("duration outside 5..60 s")
        e = sc["ego"]
        if int(e["lane"]) not in (0, 1, 2):
            errs.append("ego lane not in 0..2")
        if not 0 <= float(e["speed"]) <= 40 or not 5 <= float(e["set_speed"]) <= 40:
            errs.append("ego speed/set_speed out of range")
        if not isinstance(sc["npcs"], list) or len(sc["npcs"]) > 4:
            errs.append("npcs must be a list of at most 4")
        occupied = [(int(e["lane"]), 0.0)]
        for n in sc["npcs"]:
            if int(n["lane"]) not in (0, 1, 2):
                errs.append("npc lane not in 0..2")
            if not 0 <= float(n["speed"]) <= 40:
                errs.append("npc speed out of range")
            x = float(n["rel_x"])
            for lane, ox in occupied:
                if lane == int(n["lane"]) and abs(x - ox) < 8:
                    errs.append("vehicles overlap at start")
            occupied.append((int(n["lane"]), x))
            for ev in n.get("events", []):
                if not 0 <= float(ev["t"]) <= dur:
                    errs.append("event time outside scenario")
                if ev["type"] == "set_speed":
                    if not 0 <= float(ev["target_speed"]) <= 40 or not 0.3 <= float(ev["rate"]) <= 10:
                        errs.append("set_speed event infeasible")
                elif ev["type"] == "lane_change":
                    if int(ev["target_lane"]) not in (0, 1, 2) or abs(int(ev["target_lane"]) - int(n["lane"])) != 1:
                        errs.append("lane_change must target an adjacent lane")
                else:
                    errs.append(f"unknown event type {ev['type']}")
        orc = sc["oracle"]
        if not isinstance(orc, list) or not 1 <= len(orc) <= 3:
            errs.append("oracle must be a list of 1..3 clauses")
        else:
            for c in orc:
                if c.get("template") not in TEMPLATES:
                    errs.append(f"unknown oracle template {c.get('template')}")
                    continue
                for p in PARAMS[c["template"]]:
                    float((c.get("params") or {})[p])
    except (KeyError, TypeError, ValueError) as ex:
        errs.append(f"malformed field: {ex!r}")
    return sorted(set(errs))
