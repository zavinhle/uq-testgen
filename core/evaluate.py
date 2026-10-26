"""Ground truth for a generated test (research mode only - unavailable in practice).

A test is *correct* if its scenario exercises the requirement within its operational design domain
and its oracle gives the same verdict as the reference formalisation on the correct SUT and on every
mutant. *Effectiveness* is the fraction of mutants its own oracle kills (mutation score).
"""
from __future__ import annotations

from . import oracle, sim
from .requirements import activated, reference_oracle
from .sut import MUTANTS, VERSIONS


def ground_truth(rid: str, sc: dict, traces: dict | None = None) -> dict:
    traces = traces or {v: sim.run(sc, v) for v in VERSIONS}
    ref = reference_oracle(rid, sc)
    gen = sc["oracle"]
    act_ok, act_msg = activated(rid, sc, traces["correct"])
    ref_v = {v: oracle.evaluate(ref, tr)[0] for v, tr in traces.items()}
    gen_v = {v: oracle.evaluate(gen, tr)[0] for v, tr in traces.items()}
    in_odd = ref_v["correct"]
    false_alarm = not gen_v["correct"]
    agree = sum(ref_v[v] == gen_v[v] for v in VERSIONS) / len(VERSIONS)
    killed_ref = [m for m in MUTANTS if not ref_v[m]]
    killed_gen = [m for m in MUTANTS if not gen_v[m]] if not false_alarm else []
    correct = act_ok and in_odd and not false_alarm and agree == 1.0
    if not act_ok:
        label = "irrelevant scenario"
    elif not in_odd:
        label = "outside ODD"
    elif false_alarm:
        label = "false alarm oracle"
    elif agree < 1.0:
        label = "wrong oracle"
    elif not killed_gen:
        label = "correct but weak"
    else:
        label = "correct & effective"
    return {"activated": act_ok, "activation_note": act_msg, "in_odd": in_odd, "false_alarm": false_alarm,
            "oracle_agreement": agree, "correct": correct, "label": label,
            "ms_ref": len(killed_ref) / len(MUTANTS), "ms_gen": len(killed_gen) / len(MUTANTS),
            "killed_gen": killed_gen, "killed_ref": killed_ref,
            "ref_verdicts": ref_v, "gen_verdicts": gen_v,
            "sim_seconds_all": sum(float(t["sim_seconds"]) for t in traces.values())}
