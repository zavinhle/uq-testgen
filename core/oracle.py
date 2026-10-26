"""STL-style oracle templates with quantitative robustness (>= 0 means satisfied)."""
from __future__ import annotations

import numpy as np

BIG = 1e3

TEMPLATES = {
    "no_collision": "G (not crashed)",
    "min_gap": "G (lead_gap >= d)                         params: d [m]",
    "time_headway": "G_[t_from, end] (headway >= h)           params: h [s], t_from [s]",
    "speed_limit": "G (ego_speed <= vmax)                     params: vmax [m/s]",
    "stop_gap": "F_[end] (ego_speed <= 0.5 and lead_gap >= d)  params: d [m]",
    "max_decel": "G (ego_deceleration <= a)                 params: a [m/s^2]",
}


def _min_finite(x, default=BIG):
    x = x[np.isfinite(x)]
    return float(np.min(x)) if x.size else default


def robustness(clause: dict, tr: dict) -> float:
    k, p = clause["template"], clause.get("params", {}) or {}
    if k == "no_collision":
        return -1.0 if tr["crashed"].any() else min(1.0, _min_finite(tr["min_any_gap"], 1.0))
    if k == "min_gap":
        if tr["crashed"].any():
            return -float(p["d"]) - 1.0
        return _min_finite(tr["lead_gap"] - float(p["d"]))
    if k == "time_headway":
        m = tr["t"] >= float(p.get("t_from", 0.0))
        if tr["crashed"].any():
            return -float(p["h"])
        return _min_finite(tr["headway"][m] - float(p["h"]))
    if k == "speed_limit":
        return float(np.min(float(p["vmax"]) - tr["ego_v"]))
    if k == "stop_gap":
        if tr["crashed"].any():
            return -float(p["d"])
        g = tr["lead_gap"][-1]
        g = BIG if not np.isfinite(g) else g
        return float(min(g - float(p["d"]), 0.5 - tr["ego_v"][-1]))
    if k == "max_decel":
        return float(float(p["a"]) - np.max(-tr["ego_a"][1:])) if len(tr["ego_a"]) > 1 else BIG
    raise ValueError(f"unknown template {k}")


def evaluate(oracle: list[dict], tr: dict) -> tuple[bool, float]:
    """Conjunction of clauses; returns (pass, robustness)."""
    r = min(robustness(c, tr) for c in oracle)
    return r >= 0.0, r
