"""Executes a scenario on one SUT version in highway-env and records a trace."""
from __future__ import annotations

import copy
import time

import numpy as np
from highway_env.road.road import Road, RoadNetwork

from .sut import VERSIONS, EgoAcc, ScriptedNpc

LANES = 3
LANE_WIDTH = 4.0
SIM_HZ = 15
EGO_X0 = 50.0


def build_road(scenario: dict, version: str):
    net = RoadNetwork.straight_road_network(lanes=LANES, length=20000, speed_limit=40)
    road = Road(network=net, np_random=np.random.RandomState(0), record_history=False)
    ego_s = scenario["ego"]
    lane = net.get_lane(("0", "1", int(ego_s["lane"])))
    ego = EgoAcc(road, lane.position(EGO_X0, 0), heading=lane.heading_at(EGO_X0),
                 speed=float(ego_s["speed"]), set_speed=float(ego_s["set_speed"]), params=VERSIONS[version])
    road.vehicles.append(ego)
    npcs = []
    for i, n in enumerate(scenario.get("npcs", [])):
        nl = net.get_lane(("0", "1", int(n["lane"])))
        x = EGO_X0 + float(n["rel_x"]) + ego.LENGTH  # rel_x: npc rear vs ego front
        npc = ScriptedNpc(road, nl.position(x, 0), heading=nl.heading_at(x), speed=float(n["speed"]),
                          events=copy.deepcopy(n.get("events", [])), npc_id=n.get("id", f"npc{i}"))
        road.vehicles.append(npc)
        npcs.append(npc)
    return road, ego, npcs


def _gap_to(ego, other):
    return other.position[0] - ego.position[0] - ego.LENGTH


def run(scenario: dict, version: str = "correct", frame_cb=None) -> dict:
    """Simulate; returns a trace of per-step signals used by the oracles and UQ features."""
    t0 = time.perf_counter()
    road, ego, npcs = build_road(scenario, version)
    dt = 1.0 / SIM_HZ
    steps = int(float(scenario["duration"]) * SIM_HZ)
    tr = {k: [] for k in ("t", "ego_v", "ego_a", "lead_gap", "lead_v", "headway", "ttc", "min_any_gap", "aeb", "crashed")}
    prev_v = ego.speed
    for k in range(steps):
        road.act()
        road.step(dt)
        lead = _true_lead(ego, npcs)
        gap = _gap_to(ego, lead) if lead is not None else np.inf
        lv = lead.speed if lead is not None else np.nan
        hw = gap / ego.speed if (lead is not None and ego.speed > 1.0) else np.inf
        closing = (ego.speed - lv) if lead is not None else 0.0
        ttc = gap / closing if (lead is not None and closing > 0.1) else np.inf
        near = [abs(_gap_to(ego, n)) if n.position[0] >= ego.position[0] else abs(ego.position[0] - n.position[0] - n.LENGTH)
                for n in npcs if abs(n.position[1] - ego.position[1]) < 2.0]
        tr["t"].append((k + 1) * dt)
        tr["ego_v"].append(ego.speed)
        tr["ego_a"].append((ego.speed - prev_v) / dt)
        tr["lead_gap"].append(gap)
        tr["lead_v"].append(lv)
        tr["headway"].append(hw)
        tr["ttc"].append(ttc)
        tr["min_any_gap"].append(min(near) if near else np.inf)
        tr["aeb"].append(bool(ego.aeb_active))
        tr["crashed"].append(bool(ego.crashed))
        prev_v = ego.speed
        if frame_cb is not None:
            frame_cb(road, ego, k)
    out = {k: np.asarray(v, dtype=float) for k, v in tr.items()}
    out["sim_seconds"] = time.perf_counter() - t0
    out["lead_events"] = _lead_events(scenario)
    return out


def _true_lead(ego, npcs):
    """Ground-truth in-lane lead (independent of the SUT's possibly faulty perception)."""
    best, bdx = None, np.inf
    for n in npcs:
        dx = n.position[0] - ego.position[0]
        if 0 < dx < bdx and abs(n.position[1] - ego.position[1]) < 2.0:
            best, bdx = n, dx
    return best


def _lead_events(scenario):
    ev = []
    for n in scenario.get("npcs", []):
        for e in n.get("events", []):
            ev.append({**e, "npc_lane": n["lane"], "npc_rel_x": n["rel_x"], "npc_speed": n["speed"]})
    return ev


def render_gif(scenario: dict, version: str, path: str, every: int = 3, width: int = 960, height: int = 200,
               title: str = ""):
    """Renders a scenario with highway-env's viewer to an animated GIF with a telemetry overlay."""
    import gymnasium as gym
    import highway_env  # noqa: F401  (registers envs)
    import imageio
    from PIL import Image, ImageDraw

    env = gym.make("highway-v0", render_mode="rgb_array",
                   config={"lanes_count": LANES, "vehicles_count": 0, "screen_width": width, "screen_height": height,
                           "scaling": 7.0, "centering_position": [0.3, 0.5], "simulation_frequency": SIM_HZ})
    env.reset(seed=0)
    base = env.unwrapped
    road, ego, npcs = build_road(scenario, version)
    ego.color = (50, 200, 0)
    for n in npcs:
        n.color = (100, 200, 255)
    base.road, base.vehicle, base.controlled_vehicles = road, ego, [ego]
    frames = []
    dt = 1.0 / SIM_HZ
    for k in range(int(float(scenario["duration"]) * SIM_HZ)):
        road.act()
        road.step(dt)
        if k % every:
            continue
        img = Image.fromarray(env.render())
        lead = _true_lead(ego, npcs)
        gap = f"{_gap_to(ego, lead):5.1f} m" if lead is not None else "  --  "
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, width, 22], fill=(20, 24, 33))
        txt = f"{title}  t={(k + 1) * dt:5.1f}s   ego v={ego.speed:4.1f} m/s   gap={gap}"
        if ego.aeb_active:
            txt += "   AEB!"
        if ego.crashed:
            txt += "   COLLISION"
        d.text((8, 5), txt, fill=(255, 255, 255))
        frames.append(img)
    imageio.mimsave(path, frames, duration=every / SIM_HZ * 1000, loop=0)
    env.close()
    return path
