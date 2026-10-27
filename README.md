# UQ-TestGen

Uncertainty-aware LLM test generation for autonomous cyber-physical systems (ACPS).

LLMs can turn a natural-language requirement into an executable system test, but a well-formed test can still be
wrong: it may not exercise the requirement, its oracle may be wrong, or it may catch no faults. UQ-TestGen
samples several tests per requirement, measures how much they disagree (uncertainty), and uses that uncertainty
to decide which generated tests can be auto-accepted and which need human review.

```
requirement (NL)
  → ① LLM generation: N independent samples (scenario + STL-style oracle + verbalised confidence)
  → ② rule-based gate: schema + physical feasibility
  → ③ execution on the SUT in highway-env (ACC + AEB + lane keeping)
  → ④ agreement (no ground truth): share of the N tests that behaved the same on the SUT
       (= behavioural entropy up to ranking); research signals: structural entropy, behavioural entropy,
       oracle disagreement, verbalised confidence
  → ⑤ risk-based triage: accept (agreement ≥ 60 %) / human review / reject (< 20 %)
  → ⑥ research mode: ground truth via fault injection (8 seeded controller faults)
  → ⑦ dashboard + experiment analysis
```

## Layout

| Path | What |
|---|---|
| `core/sut.py` | ACC/AEB controller (system under test) and 8 mutants (seeded faults) |
| `core/sim.py` | Scenario execution in highway-env, traces, GIF rendering |
| `core/oracle.py` | STL-style oracle templates with quantitative robustness |
| `core/requirements.py` | 7 requirements, reference formalisation, activation conditions |
| `core/scenario.py` | Rule-based validation gate |
| `core/llm.py`, `core/generate.py` | Codex CLI adapter (ChatGPT login), prompt, sampling |
| `core/uq.py` | Uncertainty metrics and triage |
| `core/evaluate.py` | Ground truth: relevance, oracle agreement, mutation score |
| `core/pipeline.py` | End-to-end pipeline with progress events |
| `api/server.py`, `ui/index.html` | FastAPI + dashboard (live SSE progress) |
| `experiments/run_all.py`, `analysis/analyze.py` | Full experiment and analysis/figures |
| `demo/record.mjs` | Playwright script recording the demo video |

## Run

```bash
python -m venv .venv && .venv/Scripts/pip install numpy scipy scikit-learn matplotlib fastapi "uvicorn[standard]" highway-env gymnasium pillow imageio
codex login                                   # ChatGPT account
.venv/Scripts/python -m uvicorn api.server:app --port 8765   # dashboard at http://127.0.0.1:8765
.venv/Scripts/python experiments/run_all.py   # 7 requirements x 3 models x N=10
.venv/Scripts/python analysis/analyze.py      # results/summary.json + results/figures/
```

## Method notes

- **Development / held-out split.** A first round (5 requirements, `results/dev/`) was used to fix the design of
  the behavioural signature and the combined score (oracle disagreement was uninformative there and is reported
  but not combined). All reported numbers come from a fresh held-out round (`results/runs/exp_*`).
- **Ground truth.** A test is *correct* if its scenario exercises the requirement inside its operational design
  domain and its oracle gives the same verdict as the reference formalisation on the reference controller and
  all 8 mutants. *Effectiveness* is the mutation score of the test's own oracle.
- **Limits.** Small benchmark (7 requirements, one SUT, simulated), black-box LLM access only (no token
  log-probabilities), coarse per-test signals, heuristic thresholds. These are the starting points for the
  research agenda, not results to generalise.
