# Interview Q&A prep (deck v4: Demo 1 + PhD plan)

Demo 1 numbers (QA-Flow run on ShopLite, gpt-5.6-luna): 34 requirements, 62 tests (23 web, 26 API, 13 security),
31 PASS / 18 FAIL / 13 NEEDS REVIEW, 5 of 7 planted defects found, 6 of the 18 FAILs were wrong tests.

## The talk in four sentences (say this if you get lost)
1. I built QA-Flow: AI reads a requirements document, writes web, API and security tests, runs them and reports bugs.
2. It finds real bugs, but 6 of its 18 "bugs" were mistakes in the AI's own tests, and they hid 2 real bugs.
3. Its trust rules only check the form of a test, not whether it really tests the requirement.
4. My PhD: measure when an AI-generated test can be trusted, and use that to decide accept / review / reject, first for autonomous systems.

## About Demo 1

**Q: How do you know those 6 FAILs were wrong tests?**
I re-ran every failing test on a version of ShopLite with all planted defects fixed. A test that still fails there cannot be detecting a planted defect, so the test itself is wrong. I then checked each one by hand (e.g. the SAVE10 test opened an empty cart).

**Q: Why not just improve the prompt or add a rule?**
You can fix each mistake once you know it, but you cannot write a rule for every possible wrong test. The rules check form; the mistakes are about meaning. I need a signal that works without knowing the right answer in advance.

**Q: Isn't 6 of 18 just a bad prompt or a weak model?**
Maybe partly, and stronger models make fewer mistakes. But even a 5% error rate in safety testing matters, and the rules cannot tell which 5%. That is the gap, independent of the model.

**Q: The fault-injection gate: what does it do?**
Every passing test is re-run on 13 deliberately broken builds. A test that never fails cannot detect anything, so it goes to NEEDS REVIEW. It is mutation testing in spirit. It only checks passing tests, so a wrong failing test slips through.

**Q: Why a web shop when the position is about autonomous systems?**
QA-Flow comes from my daily work, and the problem is easiest to see there. The same blind spot is worse in autonomous systems, where a wrong test means false safety. The PhD starts by transferring the question to driving simulators and the InnoGuard platform, and my Bosch background is in automotive testing.

**Q: Who built QA-Flow?**
I designed the pipeline and the experiment and built this version with AI coding assistance. That is also my motivation: I rely on AI-generated artefacts every day and need to know when to trust them.

## About the PhD idea

**Q: How would you measure uncertainty of a generated test?**
Ask the AI for the same test several times and compare what the tests do when they run. If they agree, trust more; if they scatter, a human checks. This is the idea of self-consistency and semantic entropy for chatbot answers [Manakul 2023, Kuhn 2023, Farquhar 2024]; for tests, "same meaning" should be "same behaviour when executed". With open-weight models, token probabilities give white-box signals too (O1).

**Q: Have you tried it?**
I prototyped it on a driving simulator: promising, but too early to present. Turning it into reliable, calibrated metrics is exactly O1 and O3.

**Q: What if all the AI's answers agree but are wrong?**
That can happen: the AI can be consistently wrong. So agreement is a risk signal, not a proof. That is why O1 also looks at whether a test really exercises the requirement, and O2 adds the severity of the requirement to the decision.

**Q: How does this relate to the U-Taxonomy / uncertainty-wise testing?** (Prof. Ali)
Those works model uncertainty in the CPS and its environment and test the system under it. My focus is uncertainty in the AI that produces the test. They combine: uncertainty-wise test models can be the structured output the AI must fill in.

**Q: Aleatoric or epistemic?**
Disagreement mixes both: an ambiguous requirement (aleatoric, in the spec) and gaps in the model's knowledge (epistemic). Separating them, e.g. paraphrasing the requirement vs. resampling, is part of O1.

**Q: How would falsification fit in?** (Prof. Arcaini)
The AI gives requirement-grounded seeds and checks; robustness-guided falsification can refine them, and uncertainty can decide where to spend the search budget.

**Q: Metamorphic relations?** (Prof. Arrieta)
Two uses: as an oracle-free check of generated tests, and as perturbations of the prompt (paraphrases, unit changes); a trustworthy generator should give consistent tests.

**Q: What would you do at Mondragon?**
Apply the methods to an industrial CPS case (e.g. elevator dispatching), measure how much review effort they save for engineers, and collect real fault data.

**Q: How would O4 (quantum ML) work?**
Repeated shots are natural samples, so disagreement-based metrics apply directly; quantum noise adds a further source of uncertainty. Start from QML oracles for mobile robots [Wang et al., ASE 2025] and elevator QoS prediction [FSE Companion 2024].

## General

**Q: Why a PhD, and why here?**
I have tested CPS in industry and researched falsification; this position combines both with AI trustworthiness, and the team's work is exactly the foundation I need.

**Q: Eligibility?**
No doctorate. M.S. completed Feb 2026. No residence in Norway in the last 36 months. English at full professional proficiency.
