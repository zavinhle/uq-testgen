"""Uncertainty quantification for a batch of N sampled test artefacts for one requirement.

All metrics are computable WITHOUT ground truth: they use the samples themselves and their
execution on the system under test (SUT) only.

  U_verbal   1 - self-reported confidence                                (LLM verbalised baseline)
  U_struct   surprisal of the sample's structural cluster                (semantic-entropy style, Kuhn et al.)
  U_behav    surprisal of the sample's behavioural cluster, where samples are clustered by what they
             DO when executed on the SUT (novel, CPS-specific)
  U_oracle   fraction of peer oracles that disagree with the sample's own verdict on its own trace
  U_comb     mean of U_struct and U_behav (U_oracle was uninformative on the development round)
"""
from __future__ import annotations

import math
from collections import Counter

import numpy as np

from . import oracle


def _surprisal(keys: list) -> tuple[list[float], float]:
    n = len(keys)
    if n <= 1:
        return [0.0] * n, 0.0
    c = Counter(keys)
    per = [abs(-math.log(c[k] / n) / math.log(n)) for k in keys]
    ent = abs(-sum((v / n) * math.log(v / n) for v in c.values()) / math.log(n))
    return per, ent


def struct_key(sc: dict) -> tuple:
    ego = int(sc["ego"]["lane"])
    evs = []
    for n in sc["npcs"]:
        rel = "in" if int(n["lane"]) == ego else "adj"
        for e in n.get("events", []):
            if e["type"] == "set_speed":
                kind = "stop" if float(e["target_speed"]) < 0.5 else ("slow" if float(e["target_speed"]) < float(n["speed"]) else "speedup")
            else:
                kind = "cut_in" if int(e["target_lane"]) == ego else "cut_out"
            evs.append(f"{rel}:{kind}")
    tpl = tuple(sorted(c["template"] for c in sc["oracle"]))
    return tpl, tuple(sorted(set(evs))), sum(int(n["lane"]) == ego for n in sc["npcs"]) > 0


def _bucket(x, edges):
    return int(np.searchsorted(edges, x))


def behav_features(sc: dict, tr: dict) -> dict:
    ok, rob = oracle.evaluate(sc["oracle"], tr)
    g = tr["lead_gap"][np.isfinite(tr["lead_gap"])]
    ttc = tr["ttc"][np.isfinite(tr["ttc"])]
    return {"verdict": ok, "robustness": rob,
            "min_gap": float(g.min()) if g.size else 999.0,
            "min_ttc": float(ttc.min()) if ttc.size else 99.0,
            "max_decel": float(max(0.0, -tr["ego_a"].min())),
            "aeb": bool(tr["aeb"].any()), "crashed": bool(tr["crashed"].any()),
            "lead_present": float(np.isfinite(tr["lead_gap"]).mean())}


def behav_key(f: dict) -> tuple:
    # Coarse behavioural signature (chosen on the development round, see experiments/README):
    # own-oracle verdict, collision, emergency braking used, and critical / close / comfortable minimum gap.
    return (f["verdict"], f["crashed"], f["aeb"], _bucket(f["min_gap"], [2, 15]))


def quantify(samples: list[dict], traces: dict) -> dict:
    """samples: valid samples (each with 'scenario'); traces: idx -> trace on the SUT under test."""
    idx = [s["idx"] for s in samples]
    feats = {i: behav_features(s["scenario"], traces[i]) for i, s in zip(idx, samples)}
    s_per, s_ent = _surprisal([struct_key(s["scenario"]) for s in samples])
    b_per, b_ent = _surprisal([behav_key(feats[i]) for i in idx])
    per = {}
    for k, s in enumerate(samples):
        i = s["idx"]
        own = feats[i]["verdict"]
        peers = [oracle.evaluate(o["scenario"]["oracle"], traces[i])[0] for o in samples if o["idx"] != i]
        u_orc = (sum(p != own for p in peers) / len(peers)) if peers else 0.0
        conf = s["scenario"].get("confidence")
        try:
            u_verb = 1.0 - min(1.0, max(0.0, float(conf)))
        except (TypeError, ValueError):
            u_verb = 1.0
        per[i] = {"U_verbal": u_verb, "U_struct": s_per[k], "U_behav": b_per[k], "U_oracle": u_orc,
                  "U_comb": (s_per[k] + b_per[k]) / 2.0, "features": feats[i]}
    return {"per_sample": per, "batch": {"struct_entropy": s_ent, "behav_entropy": b_ent,
                                          "n_struct_clusters": len(set(struct_key(s["scenario"]) for s in samples)),
                                          "n_behav_clusters": len(set(behav_key(feats[i]) for i in idx)),
                                          "mean_verbal_conf": float(np.mean([1 - per[i]["U_verbal"] for i in idx])) if idx else None}}


ACCEPT_T, REJECT_T = 0.25, 0.55


def triage(u: float | None, valid: bool) -> str:
    if not valid or u is None:
        return "reject"
    return "accept" if u <= ACCEPT_T else ("review" if u < REJECT_T else "reject")
