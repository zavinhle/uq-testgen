# Interview Q&A prep: likely questions and grounded answers

Numbers are from the held-out experiment (`results/summary.json`): 210 tests, 7 requirements, 3 LLMs, N = 10.

## Demo 2 in three sentences (say this if you get lost)
1. I ask the AI for the same test 10 times and run all 10 in a driving simulator.
2. I group the tests by what happened (crash or not, emergency brake or not, pass or fail, how close the car got); the share of tests in a test's group is its **agreement**.
3. High agreement → accept, low → a human reviews, alone → reject. Agreement spots wrong tests far better than the AI's own confidence (0.90 vs 0.74).

## Questions the "vote" framing invites

**Q: Isn't this just majority voting? What is new?**
Sampling and comparing answers is known for chatbots (self-consistency, semantic entropy [Kuhn 2023, Farquhar 2024]). What is new is *what counts as the same answer*: for an executable test I compare **behaviour in the simulator**, not text. Two tests that look different but do the same thing agree; two that look alike but behave differently don't.

**Q: Agreement vs. behavioural entropy, which is your metric?**
The same signal in two forms. Entropy form −log(k/N)/log N and agreement 1 − k/N both depend only on the group size k and fall as k grows, so they rank tests identically (same AUROC, 0.90 / 0.71). I present agreement because it is easier to read; the backup slide shows the formal version.

**Q: What if the majority is wrong?**
It happens: the AI can be consistently wrong. In my live demo recording on 27 Oct, 4 of 8 tests for R6 formed the biggest group and all 4 were irrelevant scenarios; nothing was auto-accepted, but agreement alone could not tell them apart. In the 210-test experiment, 2 of 173 accepted tests were wrong. That is why agreement is a *risk signal*, not a proof, and why O1 (sharper signals, e.g. checking that the scenario actually exercises the requirement) and O2 (risk-aware thresholds) matter.

**Q: Why 60% and 20%?**
Hand-picked on the development round; they are not calibrated. Calibrating them against the severity of the requirement (risk = uncertainty × consequence) is O2.

**Q: Who built the prototypes?**
I designed the idea and the experiment and built the prototypes with AI coding assistance. That is also my motivation: I rely on AI-generated artefacts every day and need to know when to trust them.

## Likely from Prof. Shaukat Ali (uncertainty, foundation models, quantum)

**Q: How does your notion of uncertainty relate to the U-Taxonomy / uncertainty-wise testing?**
Those works model uncertainty *in the CPS and its environment* (sensor noise, environment conditions) and test the system under it. My focus is uncertainty *in the AI that produces the test artefact*. The two compose: a generated scenario could itself sample environment uncertainty. In the PhD I'd use uncertainty-wise test models as structured output targets for the LLM, which also gives better structural clusters.

**Q: Is this aleatoric or epistemic uncertainty?**
Sampling disagreement mixes both. Ambiguous requirements (e.g., R3's "cut-in starts") are closer to aleatoric (inherent in the spec). Model knowledge gaps are epistemic, and they shrink with stronger models: gpt-6-astra was 100% correct. Separating the two (e.g., prompt paraphrase vs. resampling) is part of O1.

**Q: Why not just use token log-probabilities?**
The models I used are only available through a ChatGPT login (no log-probs), which is realistic for industry. White-box metrics on open-weight models are planned for O1. A local open-weight model was the plan, but installing Ollama on this machine was blocked by a policy check, so I didn't run it.

**Q: How would this transfer to quantum ML (O4)?**
Repeated shots are natural samples, so disagreement-based metrics apply directly. Shot and hardware noise add a further source of uncertainty. Start from QML test oracles for mobile robots [Wang et al., ASE 2025] and elevator QoS prediction [FSE Companion 2024].

## Likely from Prof. Paolo Arcaini (ADS testing, falsification, formal methods)

**Q: Your oracles are STL-like; why not full STL with robustness?**
They are a restricted STL fragment (G, G over [t, end], final-state F) with quantitative robustness. A small DSL keeps LLM output checkable. Extending it to full STL, with oracle equivalence checked semantically, is a natural O1 step and connects to my falsification background.

**Q: Some scenarios with collisions may be unavoidable. How do you handle that?**
That's the "avoidable collision" issue [Calò et al., ICST 2020]. I label a test *outside ODD* if the reference oracle fails on the reference controller. This happened in 3 of 210 tests. A principled avoidability check (e.g., an optimal-braking oracle) is future work.

**Q: Could falsification replace LLM generation?**
They complement each other. LLMs give requirement-grounded, diverse seeds plus oracles. Falsification (MCTS or robustness-guided search) can refine them. Uncertainty could decide where to spend falsification budget.

**Q: Is highway-env realistic enough?**
No, it is a kinematic simulator chosen for speed (≈1 s to run one test on 9 versions). The PhD plan moves to higher-fidelity ADS simulators and to industrial cases. The metrics only need traces, so they are simulator-agnostic.

## Likely from Prof. Aitor Arrieta (oracles, metamorphic testing, industrial CPS, Mondragon)

**Q: Your ground truth uses 8 hand-made mutants. Is that valid?**
It is a construct threat (backup slide). The mutants mimic realistic faults: calibration, perception range/corridor, actuator saturation, logic. In industry I'd use real fault histories and expert-validated oracles. The secondment is exactly where that becomes possible.

**Q: How could metamorphic relations help?**
Two uses. First, MRs as an oracle-free ground-truth proxy for generated tests. Second, MR-based perturbation of prompts (paraphrases, unit changes) as an additional uncertainty signal: a trustworthy generator should give metamorphically consistent tests.

**Q: What would you do during the secondment at Mondragon?**
Apply the metrics and triage to an industrial CPS case (e.g., elevator dispatching, where oracles and uncertainty were studied before). Measure review-effort savings with engineers. Collect real fault data.

**Q: Within one requirement your AUROC is only 0.71. Is the approach useful?**
Honestly, it is mostly a requirement-level signal so far: batch-level behavioural entropy vs. correct-rate gives ρ = −0.85. That already helps route whole requirements to review. Per-test discrimination is the first research gap. Only 17 incorrect tests in 6 batches means the estimate is noisy too.

## General / motivation

**Q: Why a PhD, and why here?**
I have tested CPS in industry and researched falsification. This position combines both with AI trustworthiness. The team's work (uncertainty-wise CPS testing, ADS testing, industrial oracles) is exactly the foundation I need.

**Q: What did you build before vs. for this interview?**
The starting point was my LLM test generator with rule-based trust (schema, grounding, review flags, fault-injection gate). UQ-TestGen extends it with sampling, uncertainty metrics, triage and a fault-injection ground truth on an ACPS. Be ready to explain every design choice, and say plainly what was built when if asked.

**Q: Eligibility?**
No doctorate. M.S. completed Feb 2026. No residence in Norway in the last 36 months. English at full professional proficiency.

## Key numbers to remember
- AUROC for detecting incorrect tests: agreement (= behavioural entropy) 0.90, verbalised confidence 0.74
- Within-batch AUROC: agreement 0.71, verbalised 0.62
- Vote (≥ 60% accept, < 20% reject): 173 auto-accepted (171 correct, 98.8%), 30 review (19 correct, 63%), 7 reject (3 correct)
- Previous triage on the combined score (v2 deck): 162 accepted, 160 correct
- Overall correct: 92%. gpt-5.6-luna: 97% self-confidence vs 87% correct
- Cost: ~16 s and ~2,900 tokens per sample. ~1 s to simulate one test on 9 SUT versions
