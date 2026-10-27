"""Natural-language ACC requirements, their reference formalisation (ground truth, hidden from the LLM)
and the activation condition a scenario must meet to actually exercise the requirement."""
from __future__ import annotations

import numpy as np

REQUIREMENTS = {
    "R1": {
        "title": "Steady following time gap",
        "text": "When the ego vehicle follows a slower lead vehicle in its lane, after an initial settling "
                "period of 10 s the time headway to the lead vehicle shall be at least 1.2 s.",
    },
    "R2": {
        "title": "Lead vehicle emergency braking",
        "text": "If the lead vehicle in the ego lane brakes with a deceleration of up to 6 m/s^2, the ego "
                "vehicle shall not collide and shall keep a bumper-to-bumper gap of at least 2 m.",
    },
    "R3": {
        "title": "Cut-in response",
        "text": "If a vehicle from an adjacent lane cuts into the ego lane at least 15 m ahead of the ego "
                "vehicle, the ego vehicle shall not collide, and from 6 s after the cut-in starts the time "
                "headway shall be at least 1.0 s.",
    },
    "R4": {
        "title": "Set-speed limit",
        "text": "On a free road (no vehicle ahead in the ego lane), the ego speed shall never exceed the "
                "driver set speed by more than 1 m/s.",
    },
    "R5": {
        "title": "Stop-and-go",
        "text": "If the lead vehicle in the ego lane comes to a full stop, the ego vehicle shall come to a "
                "stop without collision, keeping a gap of at least 3 m to the lead vehicle.",
    },
    "R6": {
        "title": "Cut-out reveals stopped vehicle",
        "text": "If the lead vehicle changes out of the ego lane and thereby reveals a stopped vehicle ahead in "
                "the ego lane, the ego vehicle shall not collide and shall stop with a gap of at least 2 m.",
    },
    "R7": {
        "title": "Comfortable approach",
        "text": "When the ego vehicle approaches a slower vehicle in its lane starting from a gap of more than "
                "100 m with a speed difference of at most 10 m/s, it shall adapt its speed comfortably: the ego "
                "deceleration shall never exceed 3 m/s^2.",
    },
}

# Plain-language versions for people (dashboard, slides). The LLM always receives "text" above,
# exactly as in the experiment.
PLAIN = {
    "R1": ("Following a slower car",
           "When our car drives behind a slower car, then after the first 10 seconds it must keep a safe "
           "distance: at least 33 metres at 100 km/h (more when faster, less when slower)."),
    "R2": ("The car ahead brakes hard",
           "If the car ahead brakes hard, our car must not hit it and must stay at least 2 metres behind it."),
    "R3": ("A car cuts in",
           "If a car from the next lane cuts in at least 15 metres in front of us, our car must not hit it, "
           "and 6 seconds later it must again keep a safe distance: at least 28 metres at 100 km/h (more when "
           "faster, less when slower)."),
    "R4": ("Keep to the chosen speed",
           "When the road ahead is empty, our car must never drive more than 1 m/s (3.6 km/h) faster than "
           "the speed the driver chose."),
    "R5": ("Stop-and-go traffic",
           "If the car ahead stops completely, our car must also stop without hitting it, leaving at least "
           "3 metres."),
    "R6": ("A hidden stopped car",
           "If the car ahead suddenly changes lane and there is a stopped car behind it, our car must not hit "
           "the stopped car and must stop at least 2 metres before it."),
    "R7": ("Slowing down smoothly",
           "When our car comes up on a slower car from far away (more than 100 metres), it must slow down "
           "smoothly and never brake harder than a comfortable 3 m/s²."),
}
for _rid, (_title, _plain) in PLAIN.items():
    REQUIREMENTS[_rid]["plain_title"], REQUIREMENTS[_rid]["plain"] = _title, _plain


def _cut_in_time(sc):
    ego_lane = int(sc["ego"]["lane"])
    for n in sc.get("npcs", []):
        if int(n["lane"]) != ego_lane:
            for e in n.get("events", []):
                if e["type"] == "lane_change" and int(e["target_lane"]) == ego_lane:
                    return float(e["t"])
    return None


def reference_oracle(rid: str, sc: dict) -> list[dict]:
    if rid == "R1":
        return [{"template": "time_headway", "params": {"h": 1.2, "t_from": 10.0}}]
    if rid == "R2":
        return [{"template": "no_collision"}, {"template": "min_gap", "params": {"d": 2.0}}]
    if rid == "R3":
        tc = _cut_in_time(sc)
        return [{"template": "no_collision"},
                {"template": "time_headway", "params": {"h": 1.0, "t_from": (tc if tc is not None else 0.0) + 6.0}}]
    if rid == "R4":
        return [{"template": "speed_limit", "params": {"vmax": float(sc["ego"]["set_speed"]) + 1.0}}]
    if rid == "R5":
        return [{"template": "no_collision"}, {"template": "stop_gap", "params": {"d": 3.0}}]
    if rid == "R6":
        return [{"template": "no_collision"}, {"template": "stop_gap", "params": {"d": 2.0}}]
    if rid == "R7":
        return [{"template": "max_decel", "params": {"a": 3.0}}]
    raise KeyError(rid)


def activated(rid: str, sc: dict, tr: dict) -> tuple[bool, str]:
    """Does the scenario (as executed on the reference SUT) exercise the requirement's precondition?"""
    t, lv, gap, v = tr["t"], tr["lead_v"], tr["lead_gap"], tr["ego_v"]
    has_lead = np.isfinite(gap) & (gap < 150)
    dur = float(sc["duration"])
    if rid == "R1":
        m = t >= 10.0
        follow = has_lead[m] & (lv[m] < float(sc["ego"]["set_speed"]) - 0.5)
        if dur < 15:
            return False, "too short for a 10 s settling period"
        return (follow.mean() > 0.6, f"following a slower lead {follow.mean():.0%} of evaluation window")
    if rid == "R2":
        lead_acc = np.gradient(np.where(np.isfinite(lv), lv, np.nan), t)
        lead_acc = np.where(has_lead, lead_acc, np.nan)
        dmax = -np.nanmin(lead_acc) if np.isfinite(lead_acc).any() else 0.0
        ok = 2.5 <= dmax <= 6.5
        return ok, f"in-lane lead max deceleration {dmax:.1f} m/s^2 (needs 2.5..6.5)"
    if rid == "R3":
        tc = _cut_in_time(sc)
        if tc is None:
            return False, "no lane change into the ego lane"
        k = min(int(tc * 15), len(t) - 1)
        ego_lane = int(sc["ego"]["lane"])
        cutter = [n for n in sc["npcs"] if any(e["type"] == "lane_change" and int(e["target_lane"]) == ego_lane for e in n.get("events", []))][0]
        # gap at cut-in start, from the executed trace when possible, otherwise from kinematics
        rel = float(cutter["rel_x"]) + (float(cutter["speed"]) - float(sc["ego"]["speed"])) * tc
        ok = rel >= 15.0 and dur >= tc + 8.0
        return ok, f"cut-in at t={tc:.1f}s with approx. gap {rel:.0f} m (needs >=15 m, >=8 s after)"
    if rid == "R4":
        free = (~has_lead).mean()
        return free > 0.8, f"free road {free:.0%} of the time"
    if rid == "R5":
        stopped = has_lead & (np.nan_to_num(lv, nan=99) < 0.3)
        tail = stopped[t >= dur - 3.0]
        return bool(tail.size and tail.all()), "lead stopped during final 3 s" if (tail.size and tail.all()) else "lead never fully stops before end"
    if rid == "R6":
        ego_lane = int(sc["ego"]["lane"])
        movers = [n for n in sc["npcs"] if int(n["lane"]) == ego_lane and float(n["speed"]) > 1
                  and any(e["type"] == "lane_change" for e in n.get("events", []))]
        stopped = [n for n in sc["npcs"] if int(n["lane"]) == ego_lane and float(n["speed"]) < 0.5 and not n.get("events")]
        if not movers or not stopped:
            return False, "needs a moving lead that leaves the lane and a stationary vehicle in the ego lane"
        ok = float(min(s_["rel_x"] for s_ in stopped)) > float(min(m["rel_x"] for m in movers))
        end_stop = v[-1] < 0.5
        return ok and end_stop, ("stopped vehicle hidden behind the lead; ego reaches standstill" if ok and end_stop
                                 else "stopped vehicle not behind the lead or ego never stops")
    if rid == "R7":
        g0 = gap[0] if np.isfinite(gap[0]) else np.inf
        dv0 = float(sc["ego"]["speed"]) - (lv[0] if np.isfinite(lv[0]) else 0.0)
        approach = has_lead.mean() > 0.5
        ok = 100.0 < g0 < 150.0 and 0.5 < dv0 <= 10.0 and approach
        return ok, f"initial gap {g0:.0f} m (needs 100..150 within radar), speed difference {dv0:.1f} m/s (needs <=10)"
    raise KeyError(rid)
