"""Hand-written reference scenarios, one per requirement (sanity check of SUT, oracles and mutants)."""
HAND = {
 "R1": {"duration": 30, "ego": {"lane": 1, "speed": 28, "set_speed": 30},
        "npcs": [{"lane": 1, "rel_x": 60, "speed": 22, "events": []}], "oracle": []},
 "R2": {"duration": 20, "ego": {"lane": 1, "speed": 25, "set_speed": 25},
        "npcs": [{"lane": 1, "rel_x": 45, "speed": 25, "events": [{"t": 8, "type": "set_speed", "target_speed": 0, "rate": 6}]}], "oracle": []},
 "R3": {"duration": 20, "ego": {"lane": 1, "speed": 27, "set_speed": 27},
        "npcs": [{"lane": 0, "rel_x": 25, "speed": 22, "events": [{"t": 2, "type": "lane_change", "target_lane": 1}]}], "oracle": []},
 "R4": {"duration": 25, "ego": {"lane": 1, "speed": 15, "set_speed": 28}, "npcs": [], "oracle": []},
 "R5": {"duration": 30, "ego": {"lane": 1, "speed": 15, "set_speed": 20},
        "npcs": [{"lane": 1, "rel_x": 40, "speed": 14, "events": [{"t": 6, "type": "set_speed", "target_speed": 0, "rate": 2.5}]}], "oracle": []},
}
HAND["R6"] = {"duration": 25, "ego": {"lane": 1, "speed": 22, "set_speed": 22},
              "npcs": [{"lane": 1, "rel_x": 35, "speed": 22, "events": [{"t": 3, "type": "lane_change", "target_lane": 2}]},
                       {"lane": 1, "rel_x": 140, "speed": 0, "events": []}], "oracle": []}
HAND["R7"] = {"duration": 30, "ego": {"lane": 1, "speed": 30, "set_speed": 30},
              "npcs": [{"lane": 1, "rel_x": 120, "speed": 22, "events": []}], "oracle": []}
