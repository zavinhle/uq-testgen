// Builds the PhD interview deck: node build.js  ->  PhD_Proposal_ZaVinhLe.pptx
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require("./apply_theme.js");

const ROOT = path.resolve(__dirname, "..");
const S = JSON.parse(fs.readFileSync(path.join(ROOT, "results", "summary.json"), "utf8"));
const REFS = require("./refs.js");
const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "PhD_Proposal_ZaVinhLe.pptx");

// ---------- theme ----------
const HEX = { navy: "13213C", ink: "1E2533", teal: "0F8B8D", amber: "E39B2D", red: "C8463D", blue: "3567C8",
  green: "2E8B57", purple: "6E56CF", mist: "F2F5F9", line: "DDE3EC", muted: "5B6578", white: "FFFFFF", ice: "CFE3F1" };
const THEME = { name: "InnoGuard Interview", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.mist, accent1: HEX.teal, accent2: HEX.amber,
    accent3: HEX.red, accent4: HEX.blue, accent5: HEX.purple, accent6: HEX.green, hlink: HEX.blue, folHlink: HEX.purple } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Za Vinh Le";
pres.title = "When can we trust AI-generated tests? PhD research proposal";
const C = pres.SchemeColor;
const SC = { navy: C.text2, ink: C.text1, white: C.background1, mist: C.background2, teal: C.accent1, amber: C.accent2,
  red: C.accent3, blue: C.accent4, purple: C.accent5, green: C.accent6 };
const W = 13.333, H = 7.5, MX = 0.6;

// ---------- citations ----------
const order = REFS.map((r) => r[0]);
const cite = (...keys) => "[" + keys.map((k) => { const i = order.indexOf(k); if (i < 0) throw new Error("ref " + k); return i + 1; }).join(", ") + "]";

// ---------- icons ----------
async function icon(name, color = "FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "DARK", background: { color: HEX.navy },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: MX, y: 2.0, w: 11.5, h: 1.2, fontFace: "Cambria", fontSize: 40, bold: true, color: HEX.white, valign: "top", align: "left", margin: 0 }, text: "" } }],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: HEX.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.35, w: W - 2 * MX, h: 0.85, fontFace: "Cambria", fontSize: 30, bold: true, color: HEX.navy, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "Za Vinh Le  ·  PhD research proposal  ·  InnoGuard @ Simula", options: { x: MX, y: H - 0.42, w: 7, h: 0.3, fontFace: "Calibri", fontSize: 10, color: HEX.muted, margin: 0 } } },
  ],
  slideNumber: { x: W - MX - 0.6, y: H - 0.42, w: 0.6, h: 0.3, fontFace: "Calibri", fontSize: 10, color: HEX.muted, align: "right" },
});

// ---------- helpers ----------
function addTitle(s, text) { s.addText(text, { placeholder: "title" }); }
function card(s, x, y, w, h, fill = HEX.mist, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill }, objectName: name });
}
function txt(s, text, o) { s.addText(text, { isTextBox: true, margin: 0, fontFace: "Calibri", color: HEX.ink, valign: "top", ...o }); }
async function badge(s, x, y, d, iconName, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const pad = d * 0.25;
  s.addImage({ data: await icon(iconName), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
function bullets(items, size = 15, color = HEX.ink) {
  return items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, j) => ({ text: r.text, options: { fontSize: size, color, bold: !!r.bold, italic: !!r.italic,
      bullet: j === 0 ? { indent: 14 } : undefined, paraSpaceAfter: j === runs.length - 1 ? 6 : 0,
      breakLine: j === runs.length - 1 && i < items.length - 1, ...(r.color ? { color: r.color } : {}) } }));
  }).flat();
}
function source(s, text) { txt(s, text, { x: MX, y: H - 0.78, w: W - 2 * MX, h: 0.3, fontSize: 10, color: HEX.muted, italic: true }); }
const pct = (x) => Math.round(x * 100) + "%";
const f2 = (x) => (x == null ? "–" : x.toFixed(2));

// numbers from the held-out experiment
const O = S.overall, T = S.triage, AU = S.auroc_incorrect, AW = S.auroc_within_batch;
const nAcc = T.accept.n, accPrec = T.accept.precision, cov = nAcc / S.n_samples;
const luna = S.by_model["gpt-5.6-luna"], terra = S.by_model["gpt-5.6-terra"], astra = S.by_model["gpt-6-astra"];
const nCorrect = Math.round(O.correct_rate_all * S.n_samples);

(async () => {
  // ======================= 1. TITLE =======================
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" });
  txt(s, "PhD RESEARCH PROPOSAL  ·  InnoGuard Doctoral Network  ·  Simula Research Laboratory", { x: MX, y: 1.45, w: 11, h: 0.4, fontSize: 14, color: HEX.ice, bold: true, charSpacing: 1 });
  addTitle(s, "When can we trust AI-generated tests?");
  txt(s, "Quantifying the uncertainty of LLM-generated test artefacts for risk-based quality engineering of autonomous cyber-physical systems",
    { x: MX, y: 3.3, w: 9.5, h: 1.0, fontSize: 20, color: HEX.white });
  txt(s, [{ text: "Za Vinh Le", options: { bold: true, fontSize: 20, color: HEX.white, breakLine: true } },
    { text: "Interview with Prof. Shaukat Ali, Prof. Paolo Arcaini and Prof. Aitor Arrieta", options: { fontSize: 14, color: HEX.ice } }],
    { x: MX, y: 4.95, w: 9, h: 0.9 });
  s.addImage({ path: A("strip.png"), x: 0, y: H - 1.33, w: W, h: W * 160 / 1600 });
  s.addNotes("Good morning. Thank you for the invitation. In the next 20 minutes I will argue that LLM-generated tests need a trust signal, show a working prototype that quantifies their uncertainty for an autonomous-driving function, and propose how I would turn its gaps into the four InnoGuard objectives over three years.");

  // ======================= 2. ABOUT ME =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  addTitle(s, "From automotive test benches to AI-driven testing");
  const tl = [
    ["2019–23", "B.S. Electrical & Electronics Eng.", "HCMC Univ. of Technology, honors program · GPA 8.27/10", "FaMicrochip", HEX.muted],
    ["2022", "Embedded SW intern", "Drivers, firmware unit & functional testing", "FaGears", HEX.muted],
    ["2022–24", "Embedded SW test engineer · Bosch", "Automotive ECU testing for OEMs · AUTOSAR, MISRA-C, ASPICE, CAPL", "FaCar", HEX.blue],
    ["2024–26", "M.S. Computer Science · KNU", "Software Safety Eng. Lab · GPA 4.25/4.30 · falsification & property-based testing", "FaGraduationCap", HEX.teal],
    ["2026–", "Software test engineer · industry", "Web/API/system testing · LLM, RAG & MCP-based test generation", "FaRobot", HEX.amber],
  ];
  const ty = 2.15, tw = (W - 2 * MX) / tl.length;
  s.addShape(pres.shapes.LINE, { x: MX + tw / 2, y: ty + 0.32, w: tw * (tl.length - 1), h: 0, line: { color: HEX.line, width: 2 } });
  for (let i = 0; i < tl.length; i++) {
    const [yr, role, det, ic, col] = tl[i], cx = MX + i * tw;
    await badge(s, cx + tw / 2 - 0.32, ty, 0.64, ic, col);
    txt(s, yr, { x: cx, y: ty + 0.78, w: tw, h: 0.35, fontSize: 14, bold: true, color: col, align: "center" });
    txt(s, role, { x: cx + 0.08, y: ty + 1.12, w: tw - 0.16, h: 0.62, fontSize: 14, bold: true, align: "center", valign: "top" });
    txt(s, det, { x: cx + 0.1, y: ty + 1.74, w: tw - 0.2, h: 0.95, fontSize: 12, color: HEX.muted, align: "center" });
  }
  card(s, MX, 5.2, W - 2 * MX, 1.35, HEX.mist);
  txt(s, [{ text: "Thesis  ", options: { bold: true, color: HEX.teal } }, { text: "Towards Efficient Function Abstraction via Active Learning for Property Falsification (advisor: Prof. Yunja Choi)", options: { breakLine: true } },
    { text: "Papers  ", options: { bold: true, color: HEX.teal } }, { text: "Test generation for property checking (KCC 2025); stateful system, stateless techniques: test generation for controller software (KCSE 2026), with S. Hong and Y. Choi" }],
    { x: MX + 0.3, y: 5.38, w: W - 2 * MX - 0.6, h: 1.0, fontSize: 14, paraSpaceAfter: 6, valign: "middle" });
  s.addNotes("My path goes from hardware to automotive software testing at Bosch, where I tested ECU software against AUTOSAR and MISRA rules, to research on falsification and test generation for controller software at Kyungpook National University, and now to building LLM-based test generation in industry. The common thread: making automated testing of safety-relevant software trustworthy.");

  // ======================= 3. FIT =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  addTitle(s, "What I bring to the preferred profile");
  const fit = [
    ["FaCar", HEX.blue, "Testing cyber-physical systems", "Functional and system testing of automotive ECUs (CAN, CAPL, ETAS INCA); robotics (ROS/ROS2) and embedded domains"],
    ["FaWaveSquare", HEX.teal, "Verification research", "Falsification of safety and security properties, function abstraction via active learning, property-based testing — robustness-guided search like FalsifAI " + cite("zhang23falsifai")],
    ["FaBrain", HEX.amber, "LLMs for quality engineering", "LLM-assisted test-case generation, RAG over requirements, MCP-based agentic test workflows used in daily industrial testing"],
    ["FaFlask", HEX.purple, "Python & empirical studies", "Python, C/C++, TypeScript; controlled experiments on test generators (KCC'25, KCSE'26); the prototype and experiment in this talk"],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, d] = fit[i], x = MX + (i % 2) * 6.15, y = 1.55 + Math.floor(i / 2) * 2.45;
    card(s, x, y, 5.95, 2.2, HEX.mist);
    await badge(s, x + 0.3, y + 0.3, 0.7, ic, col);
    txt(s, h, { x: x + 1.2, y: y + 0.32, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, d, { x: x + 1.2, y: y + 0.85, w: 4.5, h: 1.2, fontSize: 14, color: HEX.ink });
  }
  s.addNotes("Mapping the advert's preferred profile to evidence: CPS testing from Bosch; verification research from my master's — my falsification work is close in spirit to Prof. Arcaini's robustness-guided falsification; LLMs for testing from my current work; and Python plus empirical research, which this talk also demonstrates.");

  // ======================= 4. MOTIVATION =======================
  pres.addSection({ title: "Problem" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Problem" });
  addTitle(s, "LLMs now write tests for safety-relevant systems — confidently");
  txt(s, bullets([
    [{ text: "Foundation models enter CPS engineering: ", bold: true }, { text: "scenario realism checks, oracles, digital twins, robot testing " + cite("wu24", "lu26", "ali25isola", "valleVLA") }],
    [{ text: "LLM test generation is now mainstream ", bold: true }, { text: "in software testing research " + cite("wang24survey") }],
    [{ text: "But a well-formed test can still be wrong: ", bold: true }, { text: "it may miss the requirement, carry a wrong oracle, or catch no fault" }],
    [{ text: "In ACPS a wrong test is a safety risk: ", bold: true }, { text: "false confidence in an emergency-braking or cut-in function" }],
  ], 18), { x: MX, y: 1.6, w: 6.8, h: 4.5 });
  card(s, 7.75, 1.55, 4.98, 4.5, HEX.navy);
  txt(s, "In my experiment (gpt-5.6-luna, 70 tests)", { x: 8.05, y: 1.8, w: 4.4, h: 0.4, fontSize: 14, color: HEX.ice, bold: true });
  txt(s, pct(luna.mean_verbal_conf), { x: 8.05, y: 2.25, w: 4.4, h: 1.0, fontSize: 60, bold: true, color: HEX.white, fontFace: "Cambria" });
  txt(s, "average self-reported confidence", { x: 8.05, y: 3.25, w: 4.4, h: 0.4, fontSize: 14, color: HEX.ice });
  txt(s, pct(luna.correct_rate), { x: 8.05, y: 3.85, w: 4.4, h: 1.0, fontSize: 60, bold: true, color: HEX.amber, fontFace: "Cambria" });
  txt(s, "of its tests are actually correct (fault-injection ground truth)", { x: 8.05, y: 4.85, w: 4.4, h: 0.7, fontSize: 14, color: HEX.ice });
  s.addNotes("Prof. Ali's group and colleagues already use foundation models across CPS engineering. LLM-based test generation is mainstream. The open problem is trust: a syntactically valid test can be irrelevant, can have a wrong oracle, or can be useless at finding faults. My own data shows the model says it is 97% sure while only 87% of its tests are correct — so self-reported confidence is not a usable trust signal.");

  // ======================= 5. INNOGUARD OBJECTIVES =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Problem" });
  addTitle(s, "The position's four objectives, as research questions");
  const obj = [
    ["O1", "FaGaugeHigh", HEX.teal, "New UQ metrics", "Which signals reveal that an AI-generated test artefact is untrustworthy — for black-box and white-box models?"],
    ["O2", "FaScaleBalanced", HEX.blue, "Risk & trustworthiness", "How should quantified uncertainty drive accept / review / reject decisions under safety risk?"],
    ["O3", "FaChartLine", HEX.amber, "Uncertainty vs. quality", "How does uncertainty relate to correctness, fault-detection power and cost of the generated tests?"],
    ["O4", "FaAtom", HEX.purple, "Future AI models", "Do the metrics transfer to newer large models and to quantum machine learning models?"],
  ];
  for (let i = 0; i < 4; i++) {
    const [k, ic, col, h, d] = obj[i], x = MX + i * 3.08;
    card(s, x, 1.6, 2.88, 4.3, HEX.mist);
    await badge(s, x + 0.3, 1.9, 0.75, ic, col);
    txt(s, k, { x: x + 1.2, y: 1.98, w: 1.4, h: 0.6, fontSize: 28, bold: true, color: col, fontFace: "Calibri" });
    txt(s, h, { x: x + 0.3, y: 2.9, w: 2.4, h: 0.7, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, d, { x: x + 0.3, y: 3.65, w: 2.35, h: 2.1, fontSize: 14 });
  }
  source(s, "InnoGuard: Hybrid and Generative Intelligence for Trustworthy Autonomous CPS, MSCA Doctoral Network, 2024–2028 " + cite("innoguard"));
  s.addNotes("This is how I read the four objectives of the position, phrased as questions. Everything that follows — the prototype, its gaps and the plan — is organised around O1 to O4.");

  // ======================= 6. RELATED: TEAM =======================
  pres.addSection({ title: "Background" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Background" });
  addTitle(s, "Building on uncertainty-wise CPS testing");
  const cols = [
    ["Uncertainty in CPS & DL", HEX.teal, "FaShieldHalved", [
      "U-Taxonomy & uncertainty-wise test modelling and generation " + cite("zhang16ecmfa", "zhang19sosym", "zhang19jss"),
      "Uncertainty-aware prediction validator for DL in CPS " + cite("catak22"),
      "Uncertainty-wise multi-level oracles; elevator robustness under uncertainty " + cite("valle25", "han23"),
      "Uncertainty & quality of VLA-enabled robots " + cite("valleVLA")]],
    ["ADS testing & falsification", HEX.blue, "FaRoad", [
      "Avoidable-collision scenario generation " + cite("calo20"),
      "Parameter coverage for ADS testing under uncertainty " + cite("laurent23"),
      "Robustness-guided falsification of AI-enabled hybrid systems " + cite("zhang21cav", "zhang23falsifai"),
      "LLMs assessing driving-scenario realism " + cite("wu24")]],
    ["Oracles & industrial CPS", HEX.amber, "FaIndustry", [
      "ML-based test oracles for elevator dispatching " + cite("arrieta21ast"),
      "Generated metamorphic relations for CPS " + cite("ayerdi21"),
      "Seeding strategies for search-based test selection " + cite("arrieta23seeds"),
      "Foundation models for CPS engineering: road ahead " + cite("lu26", "ali25isola")]],
  ];
  for (let i = 0; i < 3; i++) {
    const [h, col, ic, items] = cols[i], x = MX + i * 4.1;
    card(s, x, 1.55, 3.9, 4.65, HEX.mist);
    await badge(s, x + 0.25, 1.75, 0.6, ic, col);
    txt(s, h, { x: x + 1.0, y: 1.78, w: 2.75, h: 0.6, fontSize: 16, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, bullets(items, 13), { x: x + 0.25, y: 2.6, w: 3.45, h: 3.45 });
  }
  card(s, MX, 6.3, W - 2 * MX, 0.55, "FFF4E0");
  txt(s, [{ text: "Gap  ", options: { bold: true, color: HEX.amber } }, { text: "Uncertainty has been studied for CPS, DNNs and robot foundation models — not yet for LLMs that generate or assess test artefacts." }],
    { x: MX + 0.25, y: 6.33, w: W - 2 * MX - 0.5, h: 0.5, fontSize: 14, valign: "middle" });
  s.addNotes("Three lines of your work that this PhD builds on. First, uncertainty-wise testing of CPS, from the U-Taxonomy to uncertainty-aware oracles and, most recently, uncertainty and quality of VLA-enabled robots. Second, ADS scenario generation and robustness-guided falsification. Third, oracles and industrial CPS, especially the elevator case studies with Mondragon. To my knowledge none of these quantify the uncertainty of LLMs that generate the test artefacts themselves — that is the gap.");

  // ======================= 7. RELATED: LLM UQ =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Background" });
  addTitle(s, "Uncertainty of LLM outputs: what transfers to test artefacts");
  const uqrows = [
    ["Semantic entropy", "Cluster sampled answers by meaning; entropy over clusters " + cite("kuhn23", "farquhar24"), "Meaning of a test = what it does when executed → behavioural clusters"],
    ["Self-consistency", "Disagreement among sampled outputs flags hallucination " + cite("manakul23"), "Black-box: works with ChatGPT-class models without log-probs"],
    ["UQ for LLMs in SE", "Exploratory study of LLM uncertainty, incl. code tasks " + cite("huang25"), "Measures ≠ correctness of executable, safety-relevant tests"],
    ["UQ for DL supervision", "Uncertainty monitors for fail-safe DL systems " + cite("weiss21", "catak22"), "Risk-based reject option → triage of generated tests"],
  ];
  const tx = MX, tyy = 1.55;
  const head = ["Idea", "Evidence", "What I take for test artefacts"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([head, ...uqrows.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 0 ? HEX.navy : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: tx, y: tyy, w: W - 2 * MX, colW: [2.6, 5.2, 4.33], rowH: 0.8, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.1, valign: "middle" });
  card(s, MX, 5.75, W - 2 * MX, 0.9, "E6F3F3");
  txt(s, [{ text: "Key idea  ", options: { bold: true, color: HEX.teal } }, { text: "For an executable test, two samples mean the same thing if they make the system behave the same way. That gives a CPS-specific variant of semantic entropy: behavioural entropy." }],
    { x: MX + 0.25, y: 5.8, w: W - 2 * MX - 0.5, h: 0.8, fontSize: 15, valign: "middle" });
  s.addNotes("From the ML literature I take two ideas: semantic entropy, which clusters sampled answers by meaning, and self-consistency, which needs only black-box access. The key adaptation: for an executable test the meaning is its behaviour on the system — so I cluster sampled tests by what they do in simulation. That is behavioural entropy.");

  // ======================= 8. TOOL V1 =======================
  pres.addSection({ title: "Prototype" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Prototype" });
  addTitle(s, "My starting point: an LLM test generator that trusts blindly");
  const v1 = [["FaFileLines", "Requirements & specs"], ["FaDatabase", "RAG over docs"], ["FaBrain", "LLM agents"], ["FaCode", "Tests & scripts"], ["FaPlay", "Execution"], ["FaListCheck", "Report"]];
  const vx0 = MX, vw = (W - 2 * MX) / v1.length;
  for (let i = 0; i < v1.length; i++) {
    const [ic, lab] = v1[i], x = vx0 + i * vw;
    await badge(s, x + vw / 2 - 0.35, 1.6, 0.7, ic, HEX.blue);
    txt(s, lab, { x, y: 2.4, w: vw, h: 0.4, fontSize: 14, bold: true, align: "center", color: HEX.navy });
    if (i < v1.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + vw / 2 + 0.5, y: 1.83, w: vw - 1.0, h: 0.24, fill: { color: HEX.line }, line: { color: HEX.line } });
  }
  card(s, MX, 3.1, 5.95, 3.35, HEX.mist);
  txt(s, "Trust today: hand-written rules", { x: MX + 0.3, y: 3.28, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: HEX.navy });
  txt(s, bullets(["Schema validation and JSON repair of LLM output", "Grounding checks: references and quotes must exist in the source", "AI-only expectations are parked as NEEDS_REVIEW", "Fault-injection gate: a test must fail on a corrupted system"], 14), { x: MX + 0.3, y: 3.8, w: 5.4, h: 2.5 });
  card(s, 6.78, 3.1, 5.95, 3.35, "FDECEA");
  txt(s, "The blind spot", { x: 7.08, y: 3.28, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: HEX.red });
  txt(s, bullets(["Rules are binary and uncalibrated: pass the gate = trusted", "A valid, grounded test can still miss the requirement or carry a wrong oracle", "No signal says which tests a human must look at", "No link between confidence and safety risk"], 14), { x: 7.08, y: 3.8, w: 5.4, h: 2.5 });
  s.addNotes("This is the tool I built for daily testing work: requirement documents go through retrieval into LLM agents that produce test cases and scripts, which are executed and reported. Trust is handled by rules: schema checks, grounding checks, a review flag for AI-only expectations and a fault-injection gate. These catch malformed output, but they are binary: once a test passes the gate it is trusted. The tool cannot tell which of the well-formed tests are wrong — that is the blind spot this research targets.");

  // ======================= 9. TOOL V2 ARCH =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Prototype" });
  addTitle(s, "UQ-TestGen: an uncertainty-aware end-to-end pipeline");
  const st = [["FaFileLines", "Requirement", "natural language", HEX.muted], ["FaDice", "N samples", "LLM: scenario + oracle + confidence", HEX.blue], ["FaFilter", "Rule gate", "schema & physics (v1)", HEX.muted],
    ["FaCar", "Simulate", "highway-env, ACC/AEB SUT", HEX.blue], ["FaGaugeHigh", "Quantify", "4 uncertainty signals", HEX.teal], ["FaScaleBalanced", "Triage", "accept / review / reject", HEX.teal], ["FaBug", "Ground truth", "8 seeded faults (research)", HEX.amber]];
  const sw = (W - 2 * MX) / st.length;
  for (let i = 0; i < st.length; i++) {
    const [ic, h, d, col] = st[i], x = MX + i * sw;
    card(s, x + 0.06, 1.5, sw - 0.12, 2.15, HEX.mist);
    await badge(s, x + sw / 2 - 0.33, 1.68, 0.66, ic, col);
    txt(s, h, { x: x + 0.1, y: 2.45, w: sw - 0.2, h: 0.4, fontSize: 15, bold: true, align: "center", color: HEX.navy });
    txt(s, d, { x: x + 0.12, y: 2.85, w: sw - 0.24, h: 0.75, fontSize: 12, align: "center", color: HEX.muted });
  }
  s.addImage({ path: A("frame_c.png"), x: MX, y: 3.95, w: 7.3, h: 7.3 * 260 / 1600 });
  txt(s, "System under test: adaptive cruise control with emergency braking and lane keeping in highway-env " + cite("leurent18") + ". Green = ego, blue = scripted traffic.", { x: MX, y: 5.2, w: 7.3, h: 0.6, fontSize: 12, color: HEX.muted, italic: true });
  card(s, 8.2, 3.95, 4.53, 2.5, "E6F3F3");
  txt(s, "What is new vs. v1", { x: 8.45, y: 4.08, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: HEX.teal });
  txt(s, bullets(["Sample N tests instead of one", "Uncertainty without ground truth", "Risk-based triage replaces the binary gate", "Fault-injection ground truth to evaluate the signals"], 14), { x: 8.45, y: 4.55, w: 4.1, h: 1.85 });
  s.addNotes("I re-targeted the tool to an autonomous-driving function. For a requirement, the LLM is sampled N times; each sample is a driving scenario, an STL-style oracle and a self-reported confidence. The rule gate from v1 stays. Valid tests run in highway-env against an adaptive cruise control with emergency braking. Then four uncertainty signals are computed — no ground truth needed — and drive a triage decision. Only in research mode do we compute ground truth by running every test against eight seeded faults.");

  // ======================= 10. UQ METRICS =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Prototype" });
  addTitle(s, "Four uncertainty signals — no ground truth needed");
  const um = [
    ["FaUserCheck", HEX.muted, "Verbalised confidence", "U = 1 − c", "The LLM's own probability that the test is correct. Baseline."],
    ["FaLayerGroup", HEX.blue, "Structural entropy", "U = −log p(cluster) / log N", "Cluster samples by oracle templates, event types and traffic layout — semantic entropy on the artefact " + cite("kuhn23") + "."],
    ["FaCar", HEX.teal, "Behavioural entropy (new)", "U = −log p(behaviour) / log N", "Cluster samples by what they do on the SUT: own verdict, collision, emergency braking, minimum-gap band."],
    ["FaCircleNodes", HEX.purple, "Oracle disagreement", "U = disagreeing peers / (N − 1)", "Apply every peer's oracle to this sample's trace; disagreement means an ambiguous requirement."],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, f, d] = um[i], x = MX + (i % 2) * 6.15, y = 1.5 + Math.floor(i / 2) * 2.4;
    card(s, x, y, 5.95, 2.2, i === 2 ? "E6F3F3" : HEX.mist);
    await badge(s, x + 0.3, y + 0.3, 0.7, ic, col);
    txt(s, h, { x: x + 1.2, y: y + 0.25, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, f, { x: x + 1.2, y: y + 0.72, w: 4.5, h: 0.4, fontSize: 14, fontFace: "Courier New", color: col === HEX.muted ? HEX.ink : col, bold: true });
    txt(s, d, { x: x + 1.2, y: y + 1.15, w: 4.55, h: 0.95, fontSize: 14 });
  }
  txt(s, "Combined score = mean(structural, behavioural); oracle disagreement was uninformative on a development round and is reported only.", { x: MX, y: 6.38, w: W - 2 * MX, h: 0.4, fontSize: 12, color: HEX.muted, italic: true });
  s.addNotes("Four signals. Verbalised confidence is the baseline. Structural entropy clusters samples by what they look like. Behavioural entropy, my CPS-specific adaptation, clusters them by what they do when executed. Oracle disagreement cross-applies the oracles. I fixed the design on a development round and evaluated it on a fresh held-out round, to avoid tuning on the test data.");

  // ======================= 11. DEMO VIDEO =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Prototype" });
  addTitle(s, "Demo: one requirement, end to end, live LLM");
  const vw2 = 8.3, vh2 = vw2 * 900 / 1440;
  const poster = "image/png;base64," + fs.readFileSync(A("demo_50.png")).toString("base64");
  s.addMedia({ type: "video", path: A("demo.mp4"), x: MX, y: 1.4, w: vw2, h: vh2, cover: poster });
  card(s, MX + vw2 + 0.3, 1.4, W - 2 * MX - vw2 - 0.3, vh2, HEX.mist);
  txt(s, "What you see", { x: MX + vw2 + 0.55, y: 1.6, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["R6: a cut-out reveals a stopped car", "8 samples from gpt-5.6-luna", "Live stages: generate → gate → simulate → quantify → triage → ground truth", "Most uncertain test = the irrelevant one", "Replay on a seeded fault: the oracle catches it"], 13), { x: MX + vw2 + 0.55, y: 2.1, w: 2.75, h: vh2 - 0.9 });
  txt(s, "1 min 44 s · recorded from the running prototype", { x: MX, y: 1.4 + vh2 + 0.08, w: vw2, h: 0.3, fontSize: 11, color: HEX.muted, italic: true });
  s.addNotes("[Play video, about 1:44.] Requirement R6: the lead car changes lane and reveals a stopped car. Eight tests are sampled live. You can see each stage complete. The batch has low structural but some behavioural entropy. Five tests are auto-accepted, three go to review. The most uncertain test turns out to be the one that does not exercise the requirement. Finally the same test is replayed on a faulty variant and its oracle catches the fault.");

  // ======================= 12. SETUP =======================
  pres.addSection({ title: "Results" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
  addTitle(s, "Experiment: 210 generated tests, 1,890 simulations");
  const stats = [["7", "ACC requirements", "time gap, braking, cut-in, set speed, stop & go, cut-out, comfort"], ["3", "LLMs via ChatGPT login", "gpt-5.6-luna · gpt-5.6-terra · gpt-6-astra"], ["10", "samples per requirement", "independent sessions, N = 10"], ["8", "seeded controller faults", "gap, AEB, perception, actuator, logic"]];
  for (let i = 0; i < 4; i++) {
    const [n, h, d] = stats[i], x = MX + i * 3.08;
    card(s, x, 1.5, 2.88, 2.35, HEX.mist);
    txt(s, n, { x: x + 0.3, y: 1.6, w: 2.4, h: 0.95, fontSize: 48, bold: true, color: HEX.teal, fontFace: "Cambria" });
    txt(s, h, { x: x + 0.3, y: 2.55, w: 2.4, h: 0.4, fontSize: 15, bold: true, color: HEX.navy });
    txt(s, d, { x: x + 0.3, y: 2.95, w: 2.4, h: 0.8, fontSize: 12, color: HEX.muted });
  }
  txt(s, "Ground truth per generated test", { x: MX, y: 4.15, w: 6, h: 0.4, fontSize: 17, bold: true, color: HEX.navy });
  txt(s, bullets([[{ text: "Correct: ", bold: true }, { text: "scenario exercises the requirement inside its ODD, and its oracle agrees with the reference formalisation on the correct SUT and all 8 faulty versions" }],
    [{ text: "Effective: ", bold: true }, { text: "mutation score of the test's own oracle" }],
    [{ text: "Method: ", bold: true }, { text: "metric design fixed on a development round (5 requirements); all numbers here from a fresh held-out round" }]], 14), { x: MX, y: 4.6, w: W - 2 * MX, h: 1.9 });
  s.addNotes("Setup: seven requirements of an ACC, three ChatGPT-class models, ten samples each — 210 tests. Each test runs on the correct controller and eight seeded faults, so 1,890 simulations. A test is correct only if it exercises the requirement and its oracle agrees with my reference formalisation on all nine versions. I designed the metrics on a development round and report a fresh held-out round.");

  // ======================= 13. RESULT AUROC =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
  addTitle(s, "Behavioural entropy beats the LLM's self-confidence");
  const labs = ["Verbalised conf.", "Structural", "Behavioural", "Oracle disagr.", "Combined"];
  const ks = ["U_verbal", "U_struct", "U_behav", "U_oracle", "U_comb"];
  s.addChart(pres.charts.BAR, [
    { name: "All tests (pooled)", labels: labs, values: ks.map((k) => +AU[k].toFixed(2)) },
    { name: "Within one requirement × model", labels: labs, values: ks.map((k) => +AW[k].toFixed(2)) }],
  { x: MX, y: 1.45, w: 7.6, h: 4.9, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.teal, HEX.ice],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.ink, dataLabelFormatCode: "0.00",
    valAxisMinVal: 0, valAxisMaxVal: 1, valAxisMajorUnit: 0.25, valAxisLabelFormatCode: "0.00",
    valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    showTitle: true, title: "AUROC for detecting incorrect tests (0.5 = random)", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt" });
  card(s, 8.5, 1.45, 4.23, 4.9, HEX.mist);
  txt(s, f2(AU.U_behav), { x: 8.8, y: 1.6, w: 3.7, h: 0.9, fontSize: 48, bold: true, color: HEX.teal, fontFace: "Cambria" });
  txt(s, "AUROC of behavioural entropy vs. " + f2(AU.U_verbal) + " for the LLM's own confidence", { x: 8.8, y: 2.5, w: 3.7, h: 0.7, fontSize: 14 });
  txt(s, bullets([`Ranks the ${S.n_incorrect} incorrect tests (of ${S.n_samples}) mostly at the top`,
    `Batch level: behavioural entropy vs. share of correct tests, Spearman ρ = ${f2(S.batch_spearman.behav_entropy)}`,
    [{ text: "Honest caveat: ", bold: true }, { text: `within one requirement the signal drops to ${f2(AW.U_behav)} — it mostly flags hard requirements` }]], 13), { x: 8.8, y: 3.35, w: 3.75, h: 2.9 });
  source(s, `Held-out round, ${S.n_samples} tests; within-batch AUROC averaged over the ${S.n_batches_with_errors} batches that contain incorrect tests.`);
  s.addNotes(`On the held-out round, behavioural entropy detects incorrect tests with AUROC ${f2(AU.U_behav)}, against ${f2(AU.U_verbal)} for the model's own confidence. At batch level, the more behaviourally diverse the samples, the fewer of them are correct: rho ${f2(S.batch_spearman.behav_entropy)}. But I want to be honest: within a single requirement the signal is weaker, ${f2(AW.U_behav)}. Much of what it captures is that some requirements are hard. Per-test discrimination is my first research gap.`);

  // ======================= 14. RESULT TRIAGE =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
  addTitle(s, `Risk-based triage: ${pct(cov)} auto-accepted at ${(accPrec * 100).toFixed(1)}% precision`);
  const rcC = S.risk_coverage.U_comb, rcV = S.risk_coverage.U_verbal;
  const step = (rc) => { const xs = [], ys = []; for (let i = 0; i < rc.coverage.length; i += 3) { xs.push(+rc.coverage[i].toFixed(3)); ys.push(+rc.precision[i].toFixed(3)); } xs.push(1); ys.push(+rc.precision.at(-1).toFixed(3)); return [xs, ys]; };
  const [xc, yc] = step(rcC), [xv, yv] = step(rcV);
  s.addChart(pres.charts.SCATTER, [{ name: "X", values: xc }, { name: "Combined uncertainty", values: yc }],
    { x: MX, y: 1.45, w: 6.4, h: 4.9, lineSize: 2.5, lineDataSymbol: "none", chartColors: [HEX.teal],
      valAxisMinVal: 0.85, valAxisMaxVal: 1.0, valAxisMajorUnit: 0.05, valAxisLabelFormatCode: "0%", catAxisMinVal: 0, catAxisMaxVal: 1, catAxisMajorUnit: 0.2, catAxisLabelFormatCode: "0%",
      valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", showValAxisTitle: true, valAxisTitle: "Precision of accepted tests", showCatAxisTitle: true, catAxisTitle: "Coverage (share auto-accepted)",
      valAxisTitleFontSize: 12, catAxisTitleFontSize: 12, valAxisTitleColor: HEX.muted, catAxisTitleColor: HEX.muted, valAxisTitleFontFace: "+mn-lt", catAxisTitleFontFace: "+mn-lt",
      showLegend: false, showTitle: true, title: "Accept the least uncertain tests first", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt" });
  const tri = [["accept", "Auto-accept", HEX.green, "E7F4EC"], ["review", "Human review", HEX.amber, "FFF4E0"], ["reject", "Reject / regenerate", HEX.red, "FDECEA"]];
  for (let i = 0; i < 3; i++) {
    const [k, lab, col, bg] = tri[i], y = 1.45 + i * 1.32;
    card(s, 7.35, y, 5.38, 1.18, bg);
    txt(s, String(T[k].n), { x: 7.6, y: y + 0.12, w: 1.2, h: 0.9, fontSize: 40, bold: true, color: col, fontFace: "Cambria", valign: "middle" });
    txt(s, [{ text: lab, options: { bold: true, fontSize: 16, color: HEX.navy, breakLine: true } }, { text: `${T[k].correct} correct · precision ${pct(T[k].precision)}`, options: { fontSize: 14 } }],
      { x: 8.85, y: y + 0.15, w: 3.75, h: 0.9, valign: "middle" });
  }
  card(s, 7.35, 5.45, 5.38, 0.9, HEX.mist);
  txt(s, `Cost: ~${O.mean_latency_s.toFixed(0)} s and ~${Math.round(O.mean_tokens / 100) * 100} tokens per sample; ~${S.nonfunctional.sim_seconds_per_test_9_versions.toFixed(1)} s to simulate one test on 9 versions.`, { x: 7.6, y: 5.5, w: 4.95, h: 0.8, fontSize: 13, valign: "middle" });
  source(s, `Accepting all tests gives ${pct(O.correct_rate_valid)} correct. Thresholds (accept ≤ 0.25 < review < 0.55 ≤ reject) are heuristic; calibrating them against risk is objective O2.`);
  s.addNotes(`Turning uncertainty into decisions: accepting the least uncertain tests first keeps precision near 100% until about 70% coverage. With fixed thresholds, ${nAcc} of ${S.n_samples} tests are auto-accepted and ${T.accept.correct} of them are correct; the review bucket is where the errors concentrate. These thresholds are heuristics, so calibrating them against safety risk is objective 2. I also log cost: N samples multiply the LLM cost, which is a non-functional trade-off for objective 3.`);

  // ======================= 15. GAPS =======================
  pres.addSection({ title: "Research plan" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Research plan" });
  addTitle(s, "What the demo cannot do yet becomes the research agenda");
  const gaps = [
    ["Per-test signal is weak within a requirement (AUROC " + f2(AW.U_behav) + "); black-box only", "White-box (log-prob) and trajectory-level metrics; semantic oracle equivalence; calibration", "O1", HEX.teal],
    ["Thresholds are heuristic; risk ignores the severity of the requirement", "Risk = uncertainty × consequence; conformal / selective guarantees; cost-aware review budget", "O2", HEX.blue],
    ["Uncertainty vs. effectiveness only hinted (ρ = " + f2(S.spearman_ms.U_comb) + " with mutation score); cost of N samples", "Large study: correctness, fault detection, flakiness, latency and token cost vs. uncertainty", "O3", HEX.amber],
    ["3 hosted LLMs, 1 simulator, 1 SUT, seeded faults", "Open-source ACPS + industrial cases + InnoGuard platform; newer models; quantum ML models", "O4", HEX.purple],
  ];
  const gh = ["Gap seen in the demo", "Research direction", "Obj."].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([gh, ...gaps.map((g, i) => [
    { text: g[0], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[1], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[2], options: { fontSize: 18, bold: true, color: g[3], align: "center", fill: { color: i % 2 ? HEX.white : HEX.mist } } }])],
  { x: MX, y: 1.5, w: W - 2 * MX, colW: [5.4, 5.73, 1.0], rowH: 1.05, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.12, valign: "middle" });
  s.addNotes("Each limitation of the demo maps onto one objective. The per-test signal is weak and black-box only — new metrics, O1. Thresholds are heuristic and ignore severity — risk methods, O2. The relation between uncertainty, effectiveness and cost is only hinted — a large empirical study, O3. And everything ran on three hosted models and one simulator — generalisation to industrial cases, the InnoGuard platform, newer models and quantum ML, O4.");

  // ======================= 16. RQs & METHOD =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Research plan" });
  addTitle(s, "Research questions and how I will answer them");
  const rqs = [
    ["RQ1", HEX.teal, "Which uncertainty metrics best predict the correctness of AI-generated test artefacts?", "Metric suite: behavioural / semantic / oracle / white-box; AUROC, calibration error; LLMs and VLMs as generators and judges"],
    ["RQ2", HEX.blue, "How can uncertainty support risk-based acceptance of generated artefacts?", "Risk model linking uncertainty to hazard severity; selective-prediction guarantees; review-effort savings in user studies"],
    ["RQ3", HEX.amber, "How does uncertainty relate to effectiveness and non-functional properties?", "Mutation score, failure severity, flakiness, execution time and token cost; mixed-effects statistics"],
    ["RQ4", HEX.purple, "Do the metrics transfer to advanced and quantum ML models?", "Re-apply metrics to newer foundation models and to QML-based oracles/predictors " + cite("wang25ase", "wang24qelm")],
  ];
  for (let i = 0; i < 4; i++) {
    const [k, col, q, m] = rqs[i], y = 1.45 + i * 1.28;
    card(s, MX, y, W - 2 * MX, 1.12, HEX.mist);
    txt(s, k, { x: MX + 0.25, y: y + 0.18, w: 1.0, h: 0.75, fontSize: 24, bold: true, color: col, fontFace: "Cambria", valign: "middle" });
    txt(s, q, { x: MX + 1.35, y: y + 0.12, w: 5.0, h: 0.88, fontSize: 15, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, m, { x: MX + 6.55, y: y + 0.12, w: 5.35, h: 0.88, fontSize: 13, valign: "middle" });
  }
  s.addNotes("Four research questions, one per objective, each with its method. Evaluation will use open-source ACPS such as highway-env or more realistic simulators, industrial case studies through the network — for example the elevator systems studied with Mondragon — and the InnoGuard demonstration platform, always with fault-injection or expert ground truth and proper statistics.");

  // ======================= 17. GANTT =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Research plan" });
  addTitle(s, "36-month plan with secondment and publications");
  const gx = MX + 3.2, gw = W - MX - gx, gy = 1.5, rh = 0.46, months = 36;
  for (let yv2 = 0; yv2 < 3; yv2++) {
    card(s, gx + (yv2 * 12 / months) * gw + 0.02, gy - 0.05, (12 / months) * gw - 0.04, 0.4, yv2 % 2 ? HEX.mist : "E9EEF5");
    txt(s, `Year ${yv2 + 1}`, { x: gx + (yv2 * 12 / months) * gw, y: gy, w: (12 / months) * gw, h: 0.3, fontSize: 13, bold: true, align: "center", color: HEX.navy });
  }
  const wp = [["WP1 Survey & ACPS benchmark", 1, 9, HEX.muted], ["WP2 UQ metrics (O1)", 4, 18, HEX.teal], ["WP3 Uncertainty vs. quality study (O3)", 10, 28, HEX.amber],
    ["WP4 Risk & trust methods (O2)", 13, 27, HEX.blue], ["Secondment · Mondragon", 19, 23, HEX.navy], ["WP5 Industrial & InnoGuard eval.", 22, 32, HEX.green], ["WP6 Advanced & quantum ML (O4)", 26, 34, HEX.purple], ["Thesis", 31, 36, HEX.red]];
  for (let i = 0; i < wp.length; i++) {
    const [lab, a, b, col] = wp[i], y = gy + 0.5 + i * rh;
    txt(s, lab, { x: MX, y, w: 3.1, h: rh - 0.08, fontSize: 13, bold: true, color: HEX.navy, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + ((a - 1) / months) * gw, y: y + 0.06, w: ((b - a + 1) / months) * gw, h: rh - 0.2, rectRadius: 0.08, fill: { color: col }, line: { color: col } });
  }
  const pubs = [[9, "P1 mapping study"], [16, "P2 metrics · ICST/ISSTA"], [24, "P3 industrial study"], [29, "P4 journal · TOSEM/TSE"], [34, "P5 QML · ASE/ICSE"]];
  const py = gy + 0.5 + wp.length * rh + 0.12;
  for (const [pi, [m, lab]] of pubs.entries()) {
    const x = gx + ((m - 0.5) / months) * gw, dy = pi % 2 ? 0.32 : 0;
    s.addShape(pres.shapes.DIAMOND, { x: x - 0.12, y: py, w: 0.24, h: 0.24, fill: { color: HEX.amber }, line: { color: HEX.amber } });
    txt(s, lab, { x: x - 1.0, y: py + 0.28 + dy, w: 2.0, h: 0.3, fontSize: 11, align: "center", color: HEX.ink });
  }
  txt(s, "Publications", { x: MX, y: py - 0.04, w: 3.1, h: 0.32, fontSize: 13, bold: true, color: HEX.amber });
  s.addNotes("A 36-month plan. Year 1: a mapping study and an open ACPS benchmark, then the first metric suite. Year 2: the large empirical study on uncertainty versus quality and cost, risk methods, and the planned secondment at Mondragon to apply them to an industrial CPS. Year 3: evaluation on industrial cases and the InnoGuard platform, then the quantum ML extension with the adjacent project, and the thesis. Five planned papers; dates of the secondment are indicative and would follow the network's schedule.");

  // ======================= 18. CONTRIBUTIONS =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Research plan" });
  addTitle(s, "Expected contributions to InnoGuard");
  const contrib = [
    ["FaGaugeHigh", HEX.teal, "Metrics", "Uncertainty metrics for AI-generated test artefacts, incl. behavioural entropy for executable CPS tests"],
    ["FaScaleBalanced", HEX.blue, "Methods", "Risk-based acceptance of AI-generated artefacts with calibrated, safety-aware thresholds"],
    ["FaChartLine", HEX.amber, "Evidence", "Empirical laws linking uncertainty to correctness, fault detection and cost across models"],
    ["FaRocket", HEX.purple, "Open science", "Open benchmark, tool and replication packages; integration with the InnoGuard platform"],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, d] = contrib[i], y = 1.5 + i * 1.12;
    await badge(s, MX, y + 0.12, 0.7, ic, col);
    txt(s, h, { x: MX + 0.95, y: y + 0.08, w: 2.0, h: 0.8, fontSize: 18, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, d, { x: MX + 2.9, y: y + 0.08, w: 4.4, h: 0.8, fontSize: 14, valign: "middle" });
  }
  card(s, 8.1, 1.5, 4.63, 4.4, HEX.mist);
  txt(s, "Dissemination & network", { x: 8.35, y: 1.68, w: 4.2, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["Venues: ICSE, FSE, ASE, ISSTA, ICST; TOSEM, TSE, EMSE", "Joint papers across the network, building on the team's foundation-model road maps " + cite("lu26", "ali25isola"), "Secondment at Mondragon: industrial CPS case and oracle expertise", "Training, summer schools and outreach of the doctoral network"], 14), { x: 8.35, y: 2.15, w: 4.2, h: 3.6 });
  s.addNotes("Four contributions: metrics, risk methods, empirical evidence and open artefacts, published at the main SE venues and developed jointly across the network, with the Mondragon secondment as the industrial anchor.");

  // ======================= 19. CLOSING =======================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Research plan" });
  addTitle(s, "Why me, why this team");
  const why = [["FaCar", "CPS tester by training", "automotive ECU testing, standards and test benches"], ["FaWaveSquare", "Verification researcher", "falsification and test generation for controller software"], ["FaBrain", "Builder of AI testing tools", "a working uncertainty-aware prototype, today"]];
  for (let i = 0; i < 3; i++) {
    const [ic, h, d] = why[i], y = 3.75 + i * 0.95;
    await badge(s, MX, y, 0.62, ic, HEX.teal);
    txt(s, [{ text: h + "  ", options: { bold: true, color: HEX.white } }, { text: d, options: { color: HEX.ice } }], { x: MX + 0.85, y, w: 7.8, h: 0.62, fontSize: 16, valign: "middle" });
  }
  txt(s, [{ text: "Thank you", options: { bold: true, fontSize: 28, color: HEX.white, fontFace: "Cambria", breakLine: true } }, { text: "zavinhle@gmail.com", options: { fontSize: 16, color: HEX.ice } }], { x: 9.2, y: 3.75, w: 3.5, h: 1.2, align: "right" });
  s.addNotes("To close: I bring CPS testing practice, verification research and hands-on AI tooling, and the team brings the leading work on uncertainty-wise CPS testing, ADS testing and industrial oracles. I would be glad to turn this prototype into a PhD within InnoGuard. Thank you — I look forward to your questions.");

  // ======================= BACKUP =======================
  pres.addSection({ title: "Backup" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · Results per model and requirement");
  const bm = Object.entries(S.by_model);
  const mh = ["Model", "Valid", "Correct", "Mutation score", "Self-confidence", "Latency", "Tokens"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 13 } }));
  s.addTable([mh, ...bm.map(([m, v], i) => [m, pct(v.valid_rate), pct(v.correct_rate), pct(v.mean_ms_gen), pct(v.mean_verbal_conf), v.mean_latency_s.toFixed(1) + " s", String(Math.round(v.mean_tokens))].map((c) => ({ text: c, options: { fontSize: 13, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: MX, y: 1.45, w: W - 2 * MX, colW: [2.6, 1.3, 1.4, 1.9, 1.9, 1.5, 1.53], rowH: 0.42, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.08, valign: "middle" });
  const rq = Object.entries(S.by_req);
  s.addChart(pres.charts.BAR, [{ name: "Share of correct tests", labels: rq.map(([k]) => k), values: rq.map(([, v]) => +v.correct_rate.toFixed(2)) },
    { name: "Mean combined uncertainty", labels: rq.map(([k]) => k), values: rq.map(([, v]) => +v.mean_U_comb.toFixed(2)) }],
  { x: MX, y: 3.4, w: W - 2 * MX, h: 3.25, barDir: "col", chartColors: [HEX.teal, HEX.amber], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFormatCode: "0.00", dataLabelFontFace: "+mn-lt",
    valAxisMinVal: 0, valAxisMaxVal: 1.1, valAxisMajorUnit: 0.25, valAxisLabelFormatCode: "0.00", valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt" });
  s.addNotes("Per model: the strongest model, gpt-6-astra, produced only correct tests, so uncertainty matters most for cheaper models. Per requirement: the cut-in (R3) and cut-out (R6) requirements are hardest, and they also have the highest uncertainty.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · Threats to validity and mitigations");
  const thr = [["Construct", "Correctness = agreement with my reference formalisation; mutants approximate real faults", "Expert-validated oracles; real fault data from industrial partners"],
    ["Internal", "Metric design chosen on a development round; heuristic thresholds", "Held-out evaluation; pre-registered designs; calibration on separate data"],
    ["External", "One SUT, one simulator, 7 requirements, 3 hosted LLMs", "Open-source ACPS, industrial cases, InnoGuard platform, open-weight models"],
    ["Conclusion", "17 incorrect tests in 6 batches; small N per batch", "Larger N, repeated runs, effect sizes and confidence intervals"]];
  const th = ["Threat", "In this prototype", "Mitigation in the PhD"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([th, ...thr.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 0 ? HEX.navy : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: MX, y: 1.5, w: W - 2 * MX, colW: [1.8, 5.2, 5.13], rowH: 1.0, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.12, valign: "middle" });
  s.addNotes("Main threats: my reference formalisation and synthetic mutants define ground truth; the design was chosen on a development round; and the scale is small. Each has a concrete mitigation in the PhD plan.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · Towards uncertainty of quantum ML models (O4)");
  txt(s, bullets([[{ text: "Starting point: ", bold: true }, { text: "QML is already used for test oracles of autonomous mobile robots and QoS prediction of elevator software " + cite("wang25ase", "wang24qelm") }],
    [{ text: "Sources of uncertainty: ", bold: true }, { text: "shot noise, hardware noise and model (epistemic) uncertainty add to data uncertainty" }],
    [{ text: "Transfer test: ", bold: true }, { text: "do behavioural / sampling-based metrics still predict wrong oracle verdicts when the model is quantum? Repeated shots are natural samples" }],
    [{ text: "Coverage & mutation for quantum programs ", bold: true }, { text: "as ground-truth tools " + cite("ali21icst") }],
    [{ text: "Road map: ", bold: true }, { text: "quantum AI for software engineering " + cite("wang26qai") }]], 16), { x: MX, y: 1.55, w: 7.4, h: 4.8 });
  await badge(s, 9.4, 2.2, 2.4, "FaAtom", HEX.purple);
  s.addNotes("For objective 4, the adjacent project on quantum ML: shot-based sampling gives a natural analogue to sampling an LLM, so the same disagreement-based metrics can be tested, with quantum noise as an extra uncertainty source.");

  // references (2 slides)
  const half = Math.ceil(REFS.length / 2);
  for (let p = 0; p < 2; p++) {
    s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
    addTitle(s, `References (${p + 1}/2)`);
    const part = REFS.slice(p * half, (p + 1) * half);
    txt(s, part.map(([, r], i) => ({ text: `[${p * half + i + 1}] ${r}`, options: { fontSize: 10, paraSpaceAfter: 3, breakLine: i < part.length - 1 } })),
      { x: MX, y: 1.35, w: W - 2 * MX, h: 5.55, valign: "top" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
