"""Runs the full experiment: every requirement x model with N samples. Re-runnable (skips finished runs)."""
import os
import sys
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.llm import MODELS  # noqa: E402
from core.pipeline import RUNS, run_pipeline  # noqa: E402
from core.requirements import REQUIREMENTS  # noqa: E402

N = int(os.environ.get("N", 10))


def job(rid, model):
    run_id = f"exp_{rid}_{model}"
    if os.path.exists(os.path.join(RUNS, f"{run_id}.json")):
        return f"{run_id} (cached)"
    r = run_pipeline(rid, model, N, workers=4, run_id=run_id)
    ok = sum(1 for s in r["samples"] if s.get("gt", {}).get("correct"))
    return f"{run_id}: {ok}/{N} correct, {r['wall_seconds']:.0f}s"


if __name__ == "__main__":
    combos = [(r, m) for m in MODELS for r in REQUIREMENTS]
    with ThreadPoolExecutor(max_workers=3) as ex:
        for msg in ex.map(lambda c: job(*c), combos):
            print(msg, flush=True)
