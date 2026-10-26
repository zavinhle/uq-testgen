"""System under test: an ACC + AEB + lane-keeping controller, plus seeded faulty variants (mutants).

The ego vehicle is a highway-env ControlledVehicle whose longitudinal command comes from
`AccController`. Mutants change one parameter or one piece of logic each, mimicking
realistic faults (wrong calibration, sensor range, actuator saturation, logic slips).
"""
from __future__ import annotations

from dataclasses import dataclass, replace

import numpy as np
from highway_env.vehicle.controller import ControlledVehicle


@dataclass(frozen=True)
class AccParams:
    time_gap: float = 1.6          # desired time headway [s]
    min_gap: float = 4.0           # standstill gap [m]
    k_gap: float = 0.35            # gap error gain
    k_rel: float = 0.9             # relative speed gain
    k_cruise: float = 0.6          # cruise speed gain
    aeb_ttc: float = 2.2           # AEB trigger time-to-collision [s]
    max_brake: float = -9.0        # AEB deceleration [m/s^2]
    comfort_brake: float = -5.0    # ACC deceleration limit [m/s^2]
    max_accel: float = 2.0
    sensor_range: float = 150.0    # forward radar range [m]
    det_half_width: float = 2.6    # lateral half-width of the in-path detection corridor [m]
    set_speed_scale: float = 1.0   # bug hook: set speed misinterpretation
    use_rel_speed: bool = True     # bug hook: relative-speed term dropped


MUTANTS: dict[str, AccParams] = {
    "M1_short_time_gap": replace(AccParams(), time_gap=0.7),
    "M2_late_aeb": replace(AccParams(), aeb_ttc=0.6),
    "M3_narrow_corridor": replace(AccParams(), det_half_width=0.8),
    "M4_short_sensor": replace(AccParams(), sensor_range=25.0),
    "M5_set_speed_overshoot": replace(AccParams(), set_speed_scale=1.12),
    "M6_weak_actuator": replace(AccParams(), max_brake=-3.5, comfort_brake=-3.5),
    "M7_no_rel_speed": replace(AccParams(), use_rel_speed=False),
    "M8_zero_standstill_gap": replace(AccParams(), min_gap=0.5, aeb_ttc=0.9),
}
VERSIONS: dict[str, AccParams] = {"correct": AccParams(), **MUTANTS}


class EgoAcc(ControlledVehicle):
    """Ego vehicle driven by the ACC/AEB controller; keeps its lane."""

    def __init__(self, road, position, heading=0.0, speed=0.0, set_speed=25.0, params: AccParams = AccParams()):
        super().__init__(road, position, heading=heading, speed=speed)
        self.set_speed = set_speed
        self.params = params
        self.aeb_active = False
        self.lead = None

    def find_lead(self):
        p = self.params
        best, best_dx = None, np.inf
        for v in self.road.vehicles:
            if v is self:
                continue
            dx = v.position[0] - self.position[0]
            dy = abs(v.position[1] - self.position[1])
            if 0 < dx < best_dx and dx - self.LENGTH <= p.sensor_range and dy < p.det_half_width:
                best, best_dx = v, dx
        return best

    def acc_command(self) -> float:
        p = self.params
        v = self.speed
        a = p.k_cruise * (self.set_speed * p.set_speed_scale - v)
        self.lead = lead = self.find_lead()
        self.aeb_active = False
        if lead is not None:
            gap = lead.position[0] - self.position[0] - self.LENGTH
            dv = lead.speed - v
            desired = p.min_gap + p.time_gap * v
            a_follow = p.k_gap * (gap - desired) + (p.k_rel * dv if p.use_rel_speed else 0.0)
            a = min(a, a_follow)
            a = max(a, p.comfort_brake)
            closing = -dv
            if closing > 0.1 and gap / closing < p.aeb_ttc:
                a, self.aeb_active = p.max_brake, True
            if gap < 0.5 * p.min_gap and v > 0.1:
                a, self.aeb_active = p.max_brake, True
        a = min(a, p.max_accel)
        if v <= 0.0 and a < 0:
            a = -v  # no reversing
        return float(a)

    def act(self, action=None):
        if self.crashed:
            super(ControlledVehicle, self).act({"steering": 0.0, "acceleration": -self.speed})
            return
        self.follow_road()
        steering = self.steering_control(self.target_lane_index)
        super(ControlledVehicle, self).act({"steering": steering, "acceleration": self.acc_command()})


class ScriptedNpc(ControlledVehicle):
    """Traffic participant that follows a timed list of scenario events."""

    def __init__(self, road, position, heading=0.0, speed=0.0, events=None, npc_id="npc"):
        super().__init__(road, position, heading=heading, speed=speed)
        self.npc_id = npc_id
        self.events = sorted(events or [], key=lambda e: e["t"])
        self.cmd_speed = speed
        self.rate = 2.0
        self.t = 0.0

    def act(self, action=None):
        if self.crashed:
            super(ControlledVehicle, self).act({"steering": 0.0, "acceleration": -self.speed})
            return
        for e in self.events:
            if not e.get("_done") and self.t >= e["t"]:
                e["_done"] = True
                if e["type"] == "set_speed":
                    self.cmd_speed, self.rate = float(e["target_speed"]), float(e["rate"])
                elif e["type"] == "lane_change":
                    self.target_lane_index = (self.lane_index[0], self.lane_index[1], int(e["target_lane"]))
        self.follow_road()
        err = self.cmd_speed - self.speed
        acc = float(np.clip(3.0 * err, -self.rate, self.rate))
        if self.speed <= 0.0 and acc < 0:
            acc = -self.speed
        super(ControlledVehicle, self).act({"steering": self.steering_control(self.target_lane_index), "acceleration": acc})

    def step(self, dt):
        super().step(dt)
        self.t += dt
