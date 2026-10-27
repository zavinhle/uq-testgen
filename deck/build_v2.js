// Builds the v2 interview deck (Motivation 1 -> Demo 1 -> Motivation 2 -> Demo 2 -> Gaps -> Me):
//   node build_v2.js  ->  PhD_Proposal_ZaVinhLe_v2.pptx
// Carries over the user's edits of v1 (title slide, motivation wording, per-slide reference footnotes, order).
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
const QF = JSON.parse(fs.readFileSync(path.join(__dirname, "assets", "qaflow_run.json"), "utf8")); // Demo 1 summary + ground truth
const REFS = require("./refs.js");
const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "PhD_Proposal_ZaVinhLe_v2.pptx");

// ---------- theme ----------
const HEX = { navy: "13213C", ink: "1E2533", teal: "0F8B8D", amber: "E39B2D", red: "C8463D", blue: "3567C8",
  green: "2E8B57", purple: "6E56CF", mist: "F2F5F9", line: "DDE3EC", muted: "5B6578", white: "FFFFFF", ice: "CFE3F1", indigo: "4F46E5" };
const THEME = { name: "InnoGuard Interview", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.mist, accent1: HEX.teal, accent2: HEX.amber,
    accent3: HEX.red, accent4: HEX.blue, accent5: HEX.purple, accent6: HEX.green, hlink: HEX.blue, folHlink: HEX.purple } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Za Vinh Le";
pres.title = "When can we trust AI-generated tests? PhD research proposal";
const W = 13.333, H = 7.5, MX = 0.6;

// ---------- citations + per-slide footnotes (user's v1 style) ----------
const order = REFS.map((r) => r[0]);
let slideRefs = new Set();
const cite = (...keys) => "[" + keys.map((k) => { const i = order.indexOf(k); if (i < 0) throw new Error("ref " + k); slideRefs.add(k); return i + 1; }).join(", ") + "]";
function shortRef(key) {
  const full = REFS[order.indexOf(key)][1];
  const m = /^(.*?), "(.*?),?" (.*)$/.exec(full);
  if (!m) return full;
  const [, authors, title, rest] = m;
  const first = authors.split(/,| and /)[0].trim();
  const etal = /,| and /.test(authors) ? " et al." : "";
  const year = (rest.replace(/,? doi:.*$/, "").replace(/https?:\S+/, "").match(/\b(19|20)\d{2}\b/g) || []).slice(-1)[0] || "";
  let venue = rest.replace(/^in /, "").split(/, (vol\.|no\.|pp\.|art\.|doi:|LNCS|\d{4})/)[0];
  if (/arXiv/.test(rest)) venue = rest.split(",")[0];
  return `${first}${etal}, “${title},” ${venue}${year && !venue.includes(year) ? ", " + year : ""}.`;
}
function newSlide(opts) { slideRefs = new Set(); return pres.addSlide(opts); }
/** Footnote with every reference cited on the slide, sorted by number. */
function footnotes(s, y, size = 9, sep = "\n") {
  const keys = [...slideRefs].sort((a, b) => order.indexOf(a) - order.indexOf(b));
  if (!keys.length) return;
  const text = keys.map((k) => `[${order.indexOf(k) + 1}] ${shortRef(k)}`).join(sep);
  s.addText(text, { isTextBox: true, x: MX, y, w: W - 2 * MX, h: H - 0.47 - y, fontSize: size, color: HEX.muted, fontFace: "Calibri", margin: 0, valign: "bottom" });
}

// ---------- icons ----------
async function icon(name, color = "FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "DARK", background: { color: HEX.navy },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: MX, y: 2.0, w: 11.5, h: 1.2, fontFace: "Calibri", fontSize: 40, bold: true, color: HEX.white, valign: "top", align: "left", margin: 0 }, text: "" } }],
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
function source(s, text, y = H - 0.78) { txt(s, text, { x: MX, y, w: W - 2 * MX, h: 0.3, fontSize: 10, color: HEX.muted, italic: true }); }
/** Small part label above the title, e.g. "PART A · DEMO 1". */
function kicker(s, text, color = HEX.indigo) { txt(s, text, { x: MX, y: 0.12, w: 8, h: 0.25, fontSize: 11, bold: true, color, charSpacing: 1 }); }
const pct = (x) => Math.round(x * 100) + "%";
const f2 = (x) => (x == null ? "–" : x.toFixed(2));

// numbers: Demo 2 (UQ-TestGen held-out experiment)
const O = S.overall, T = S.triage, AU = S.auroc_incorrect, AW = S.auroc_within_batch;
const nAcc = T.accept.n, accPrec = T.accept.precision, cov = nAcc / S.n_samples;
const luna = S.by_model["gpt-5.6-luna"];
// numbers: Demo 1 (QA-Flow run + offline ground truth)
const Q = QF.summary, G = QF.groundTruth;
const nWrong = G.wrongTests.length, nFail = G.fails, nFound = G.defectsFound.length;

(async () => {
  // ======================= 1. TITLE (user's v1 version) =======================
  pres.addSection({ title: "Opening" });
  let s = newSlide({ masterName: "DARK", sectionTitle: "Opening" });
  txt(s, "PhD RESEARCH PROPOSAL", { x: MX, y: 1.45, w: 11, h: 0.4, fontSize: 14, color: HEX.ice, bold: true, charSpacing: 1 });
  s.addText([{ text: "When can we trust AI-generated tests?", options: { breakLine: true } },
    { text: "Uncertainty quantification for risk-based testing of autonomous CPS" }], { placeholder: "title" });
  txt(s, "Za Vinh Le", { x: MX, y: 4.95, w: 9, h: 0.5, fontSize: 20, bold: true, color: HEX.white });
  s.addImage({ path: A("strip.png"), x: 0, y: H - 1.33, w: W, h: W * 160 / 1600 });
  s.addNotes("Good morning, and thank you for the invitation. My talk has four parts: first the tool I started from and its demo, then why its trust rules are not enough and how I quantify uncertainty instead, with a second demo; then the gaps that become my PhD plan along the four objectives; and finally what I bring.");

  // ======================= PART A: MOTIVATION 1 -> DEMO 1 =======================
  pres.addSection({ title: "A · Motivation 1 → Demo 1" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · Motivation 1 → Demo 1" });
  kicker(s, "PART A · MOTIVATION 1");
  addTitle(s, "Turning specifications into tests by hand does not scale");
  const loop = [["FaFileLines", "Read the spec"], ["FaListCheck", "Write test cases"], ["FaCode", "Write scripts"], ["FaPlay", "Run & triage"], ["FaBug", "Report bugs"]];
  const lw = (W - 2 * MX) / loop.length;
  for (let i = 0; i < loop.length; i++) {
    const [ic, lab] = loop[i], x = MX + i * lw;
    await badge(s, x + lw / 2 - 0.35, 1.55, 0.7, ic, HEX.muted);
    txt(s, lab, { x, y: 2.35, w: lw, h: 0.35, fontSize: 15, bold: true, align: "center", color: HEX.navy });
    txt(s, "manual", { x, y: 2.7, w: lw, h: 0.3, fontSize: 12, align: "center", color: HEX.red, italic: true });
    if (i < loop.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + lw / 2 + 0.5, y: 1.78, w: lw - 1.0, h: 0.24, fill: { color: HEX.line }, line: { color: HEX.line } });
  }
  txt(s, bullets([
    [{ text: "Every release, every spec change: ", bold: true }, { text: "test cases and scripts are re-written for web, API and security testing" }],
    [{ text: "Test design is the bottleneck, ", bold: true }, { text: "not test execution" }],
    [{ text: "LLMs can now read specifications and write tests " + cite("wang24survey") + ", ", bold: true }, { text: "which makes end-to-end automation possible" }],
  ], 17), { x: MX, y: 3.35, w: 7.1, h: 2.6 });
  card(s, 8.0, 3.35, 4.73, 2.45, "EEF0FE");
  txt(s, "Goal of my tool", { x: 8.3, y: 3.55, w: 4.2, h: 0.4, fontSize: 16, bold: true, color: HEX.indigo });
  txt(s, "One pipeline from a requirements document to executed web, API and security tests and a pass/fail report — with the LLM doing the test design.", { x: 8.3, y: 4.0, w: 4.2, h: 1.7, fontSize: 15 });
  footnotes(s, 6.45);
  s.addNotes("The first motivation comes from my daily work as a tester. For every release we read the specification, write test cases, write scripts, run them and file bugs — and when the specification changes, we do it again. The bottleneck is test design, not execution. LLMs can now read specifications and write tests, so I asked: can one pipeline go from the document to the report?");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · Motivation 1 → Demo 1" });
  kicker(s, "PART A · DEMO 1");
  addTitle(s, "QA-Flow: from a requirements document to a pass/fail report");
  const v1 = [["FaFileLines", "Requirements & specs", "SRS + discovered app"], ["FaDatabase", "RAG over docs", "chunks + BM25"], ["FaBrain", "LLM agents", "analyst · designer · bug writer"], ["FaCode", "Tests & scripts", "step DSL → Playwright"], ["FaPlay", "Execution", "web · API · security"], ["FaListCheck", "Report", "traceability + bugs"]];
  const vw = (W - 2 * MX) / v1.length;
  for (let i = 0; i < v1.length; i++) {
    const [ic, lab, sub] = v1[i], x = MX + i * vw;
    await badge(s, x + vw / 2 - 0.35, 1.45, 0.7, ic, HEX.indigo);
    txt(s, lab, { x, y: 2.25, w: vw, h: 0.35, fontSize: 14, bold: true, align: "center", color: HEX.navy });
    txt(s, sub, { x, y: 2.6, w: vw, h: 0.3, fontSize: 12, align: "center", color: HEX.muted });
    if (i < v1.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + vw / 2 + 0.5, y: 1.68, w: vw - 1.0, h: 0.24, fill: { color: HEX.line }, line: { color: HEX.line } });
  }
  card(s, MX, 3.2, 5.95, 3.2, HEX.mist);
  txt(s, "Trust: hand-written rules", { x: MX + 0.3, y: 3.38, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: HEX.navy });
  txt(s, bullets(["Schema validation and JSON repair of every agent reply", "Grounding: refs must exist, quotes must be verbatim", "AI-only tests are parked as NEEDS REVIEW", `Fault-injection gate: every passing test is re-run on ${QF.faults} broken builds`], 15), { x: MX + 0.3, y: 3.88, w: 5.4, h: 2.4 });
  card(s, 6.78, 3.2, 5.95, 3.2, "EEF0FE");
  txt(s, "Demo application: ShopLite", { x: 7.08, y: 3.38, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: HEX.indigo });
  txt(s, bullets(["Web shop + REST API + OpenAPI spec + requirements document", "7 planted defects: 1 web, 2 API, 4 security (XSS, IDOR, CSP, user enumeration)", "Security tests run only against this local demo app", "LLM: gpt-5.6-luna through the Codex CLI"], 15), { x: 7.08, y: 3.88, w: 5.4, h: 2.4 });
  s.addNotes("This is QA-Flow, the tool I started from. Its six stages: the requirements document plus a discovery of the running application; retrieval over the documents; three LLM agents; a closed step language that code compiles into Playwright scripts; execution of web, API and security tests; and a report. Trust is handled by four hand-written rules. The demo application, ShopLite, contains seven planted defects so we know what should be found.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · Motivation 1 → Demo 1" });
  kicker(s, "PART A · DEMO 1");
  addTitle(s, "Demo 1: QA-Flow end to end on ShopLite");
  const vwA = 8.3, vhA = vwA * 900 / 1440;
  s.addMedia({ type: "video", path: A("qaflow_demo.mp4"), x: MX, y: 1.4, w: vwA, h: vhA, cover: "image/png;base64," + fs.readFileSync(A("qaflow_poster.png")).toString("base64") });
  card(s, MX + vwA + 0.3, 1.4, W - 2 * MX - vwA - 0.3, vhA, HEX.mist);
  txt(s, "What you see", { x: MX + vwA + 0.55, y: 1.6, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["Six stages run live", "Analyst extracts requirements as verbatim quotes", "Designer writes web, API and security tests", "Code compiles them to Playwright", `Gate re-runs tests on ${QF.faults} broken builds`, "Report with drafted bugs"], 13), { x: MX + vwA + 0.55, y: 2.1, w: 2.75, h: vhA - 0.9 });
  txt(s, `${QF.videoLabel} · recorded from the running tool (long waits sped up)`, { x: MX, y: 1.4 + vhA + 0.08, w: vwA, h: 0.3, fontSize: 11, color: HEX.muted, italic: true });
  s.addNotes("[Play video.] The document goes in, the application is discovered and the documents are indexed. The analyst agent extracts atomic requirements as verbatim quotes, the designer writes tests per requirement using retrieved context, code compiles them into Playwright, they run, and the fault-injection gate re-runs passing tests on broken builds. The report shows traceability and drafted bug reports.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · Motivation 1 → Demo 1" });
  kicker(s, "PART A · DEMO 1 RESULTS");
  addTitle(s, `It finds the bugs — but ${nWrong} of its ${nFail} FAILs are wrong tests`);
  const tiles = [[String(Q.requirements), "requirements", "extracted from the SRS"], [String(Q.tests), "tests", `web ${Q.byKind.web} · API ${Q.byKind.api} · security ${Q.byKind.security}`],
    [`${Q.pass} / ${Q.fail} / ${Q.review}`, "PASS / FAIL / REVIEW", "verdicts by the rules"], [`${nFound}/7`, "planted defects found", G.defectsMissed.length ? "missed: " + G.defectsMissed.join("; ") : "all planted defects detected"]];
  for (let i = 0; i < 4; i++) {
    const [n, h, d] = tiles[i], x = MX + i * 3.08;
    card(s, x, 1.45, 2.88, 1.95, HEX.mist);
    txt(s, n, { x: x + 0.25, y: 1.55, w: 2.5, h: 0.8, fontSize: i === 2 ? 30 : 40, bold: true, color: i === 3 ? HEX.green : HEX.indigo, fontFace: "Cambria", valign: "middle" });
    txt(s, h, { x: x + 0.25, y: 2.38, w: 2.5, h: 0.35, fontSize: 15, bold: true, color: HEX.navy });
    txt(s, d, { x: x + 0.25, y: 2.73, w: 2.5, h: 0.6, fontSize: 11, color: HEX.muted });
  }
  card(s, MX, 3.65, 5.95, 2.85, "E7F4EC");
  txt(s, "Real defects it found", { x: MX + 0.3, y: 3.8, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: HEX.green });
  txt(s, bullets(G.defectsFound.slice(0, 7).map((d) => d.defect), 13), { x: MX + 0.3, y: 4.25, w: 5.4, h: 2.2 });
  card(s, 6.78, 3.65, 5.95, 2.85, "FDECEA");
  txt(s, `The twist: ${nWrong} FAILs are mistakes in the AI's tests`, { x: 7.08, y: 3.8, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: HEX.red });
  txt(s, bullets([...QF.wrongExamples.slice(0, 2), "All of them passed every trust rule: valid JSON, real refs, verbatim quotes", "The report files them as bugs, exactly like the real ones"], 13), { x: 7.08, y: 4.25, w: 5.4, h: 2.2 });
  source(s, "Ground truth computed offline: every FAIL re-run on ShopLite builds with the planted defects fixed (a FAIL that persists = wrong test).");
  s.addNotes(`The result: QA-Flow extracted ${Q.requirements} requirements and ran ${Q.tests} tests, and it found ${nFound} of the 7 planted defects. But when I re-ran every failing test on a build with the planted defects fixed, ${nWrong} of the ${nFail} FAILs still failed: the bug was in the AI's test, not in the shop. Every one of them passed all four trust rules. So the rules cannot tell a wrong test from a real defect. That is my second motivation.`);

  // ======================= PART B: MOTIVATION 2 -> DEMO 2 =======================
  pres.addSection({ title: "B · Motivation 2 → Demo 2" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · MOTIVATION 2", HEX.teal);
  addTitle(s, "LLMs now write tests for safety-relevant systems");
  txt(s, bullets([
    [{ text: "AI is already in CPS engineering: ", bold: true }, { text: "LLMs check driving scenarios, write oracles, build digital twins " + cite("wu24", "lu26") }],
    [{ text: "LLMs now write the tests themselves: ", bold: true }, { text: "a major topic in software testing research " + cite("wang24survey") }],
    [{ text: "A test can run cleanly and still be wrong ", bold: true }, { text: "in three ways:" }],
  ], 18), { x: MX, y: 1.58, w: 6.81, h: 1.81 });
  const wrong = ["Wrong scenario", "Wrong oracle", "Catches no fault"];
  for (let i = 0; i < 3; i++) {
    const x = MX + i * 2.32;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.42, w: 2.17, h: 0.69, fill: { color: "FDECEA" }, line: { color: "FDECEA" } });
    txt(s, wrong[i], { x: x + 0.15, y: 3.42, w: 1.95, h: 0.69, fontSize: 15, bold: true, color: HEX.red, valign: "middle" });
  }
  txt(s, bullets([[{ text: "In ACPS a wrong test is a safety risk: ", bold: true }, { text: "false confidence in an emergency-braking or cut-in function" }],
    [{ text: "Demo 1 showed it: ", bold: true }, { text: "rules cannot tell a wrong test from a real defect" }]], 18), { x: MX, y: 4.4, w: 6.81, h: 1.5 });
  card(s, 7.75, 1.55, 4.98, 4.5, HEX.navy);
  txt(s, "In my experiment (gpt-5.6-luna, 70 tests)", { x: 8.05, y: 1.8, w: 4.4, h: 0.4, fontSize: 14, color: HEX.ice, bold: true });
  txt(s, pct(luna.mean_verbal_conf), { x: 8.05, y: 2.25, w: 4.4, h: 1.0, fontSize: 60, bold: true, color: HEX.white, fontFace: "Cambria" });
  txt(s, "average self-reported confidence", { x: 8.05, y: 3.25, w: 4.4, h: 0.4, fontSize: 14, color: HEX.ice });
  txt(s, pct(luna.correct_rate), { x: 8.05, y: 3.85, w: 4.4, h: 1.0, fontSize: 60, bold: true, color: HEX.amber, fontFace: "Cambria" });
  txt(s, "of its tests are actually correct", { x: 8.05, y: 4.85, w: 4.4, h: 0.7, fontSize: 14, color: HEX.ice });
  footnotes(s, 6.2);
  s.addNotes("The second motivation: in autonomous CPS the stakes are higher. Foundation models are already used across CPS engineering, and LLMs write the tests themselves. A test can run cleanly and still be wrong in three ways: the wrong scenario, the wrong oracle, or a test that catches no fault. Demo 1 showed that rules cannot see this, and the model's own confidence does not help either: it says 97% while only 87% of its tests are correct.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · STATE OF THE ART", HEX.teal);
  addTitle(s, "Building on uncertainty-wise CPS testing and LLM uncertainty");
  const cols = [
    ["Uncertainty in CPS & DL", HEX.teal, "FaShieldHalved", ["U-Taxonomy, uncertainty-wise test modelling & generation " + cite("zhang16ecmfa", "zhang19sosym", "zhang19jss"),
      "DL prediction validator; uncertainty-wise oracles " + cite("catak22", "valle25"), "Uncertainty & quality of VLA robots " + cite("valleVLA")]],
    ["ADS testing & falsification", HEX.blue, "FaRoad", ["Avoidable-collision scenarios " + cite("calo20"), "Parameter coverage under uncertainty " + cite("laurent23"),
      "Robustness-guided falsification " + cite("zhang21cav", "zhang23falsifai"), "LLMs judging scenario realism " + cite("wu24")]],
    ["Oracles & industrial CPS", HEX.amber, "FaIndustry", ["ML-based oracles for elevators " + cite("arrieta21ast"), "Metamorphic relations for CPS " + cite("ayerdi21"),
      "Elevator robustness under uncertainty " + cite("han23"), "Foundation models for CPS " + cite("lu26", "ali25isola")]],
    ["Uncertainty of LLM outputs", HEX.purple, "FaBrain", ["Semantic entropy " + cite("kuhn23", "farquhar24"), "Self-consistency (SelfCheckGPT) " + cite("manakul23"),
      "LLM uncertainty in SE tasks " + cite("huang25")]],
  ];
  for (let i = 0; i < 4; i++) {
    const [h, col, ic, items] = cols[i], x = MX + i * 3.08;
    card(s, x, 1.45, 2.92, 3.05, HEX.mist);
    await badge(s, x + 0.2, 1.6, 0.55, ic, col);
    txt(s, h, { x: x + 0.85, y: 1.6, w: 1.95, h: 0.55, fontSize: 14, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, bullets(items, 12), { x: x + 0.2, y: 2.3, w: 2.6, h: 2.15 });
  }
  card(s, MX, 4.62, W - 2 * MX, 0.55, "FFF4E0");
  txt(s, [{ text: "Gap  ", options: { bold: true, color: HEX.amber } }, { text: "Uncertainty is studied for CPS, DNNs, robot foundation models and LLM text — not yet for LLMs that generate executable test artefacts." }],
    { x: MX + 0.25, y: 4.65, w: W - 2 * MX - 0.5, h: 0.5, fontSize: 14, valign: "middle" });
  footnotes(s, 5.3, 7.5, "   ");
  s.addNotes("This builds on your work: uncertainty-wise CPS testing from the U-Taxonomy to uncertainty-aware oracles and VLA robots; ADS scenario generation and robustness-guided falsification; ML oracles and metamorphic testing for industrial elevators; and from ML, semantic entropy and self-consistency for LLM outputs. The gap: nobody quantifies the uncertainty of LLMs that generate executable test artefacts.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2", HEX.teal);
  addTitle(s, "UQ-TestGen: an uncertainty-aware end-to-end pipeline");
  const st = [["FaFileLines", "Requirement", "natural language", HEX.muted], ["FaDice", "N samples", "LLM: scenario + oracle + confidence", HEX.blue], ["FaFilter", "Rule gate", "schema & physics (from QA-Flow)", HEX.muted],
    ["FaCar", "Simulate", "highway-env, ACC/AEB SUT", HEX.blue], ["FaGaugeHigh", "Quantify", "4 uncertainty signals", HEX.teal], ["FaScaleBalanced", "Triage", "accept / review / reject", HEX.teal], ["FaBug", "Ground truth", "8 seeded faults (research)", HEX.amber]];
  const sw = (W - 2 * MX) / st.length;
  for (let i = 0; i < st.length; i++) {
    const [ic, h, d, col] = st[i], x = MX + i * sw;
    card(s, x + 0.06, 1.45, sw - 0.12, 2.15, HEX.mist);
    await badge(s, x + sw / 2 - 0.33, 1.63, 0.66, ic, col);
    txt(s, h, { x: x + 0.1, y: 2.4, w: sw - 0.2, h: 0.4, fontSize: 15, bold: true, align: "center", color: HEX.navy });
    txt(s, d, { x: x + 0.12, y: 2.8, w: sw - 0.24, h: 0.75, fontSize: 12, align: "center", color: HEX.muted });
  }
  s.addImage({ path: A("frame_c.png"), x: MX, y: 3.85, w: 7.3, h: 7.3 * 260 / 1600 });
  txt(s, "System under test: adaptive cruise control with emergency braking and lane keeping in highway-env " + cite("leurent18") + ". Green = ego, blue = scripted traffic.", { x: MX, y: 5.1, w: 7.3, h: 0.6, fontSize: 12, color: HEX.muted, italic: true });
  card(s, 8.2, 3.85, 4.53, 2.35, "E6F3F3");
  txt(s, "What is new vs. QA-Flow", { x: 8.45, y: 3.98, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: HEX.teal });
  txt(s, bullets(["Sample N tests instead of one", "Uncertainty without ground truth", "Risk-based triage replaces the binary gate", "Fault-injection ground truth to evaluate the signals"], 14), { x: 8.45, y: 4.45, w: 4.1, h: 1.7 });
  footnotes(s, 6.55);
  s.addNotes("UQ-TestGen keeps QA-Flow's rule gate but targets an autonomous-driving function. For each requirement the LLM is sampled N times; each sample is a driving scenario, an STL-style oracle and a confidence. Valid tests run in highway-env against an adaptive cruise control with emergency braking. Four uncertainty signals, computed without ground truth, drive the triage. Only in research mode is ground truth computed with eight seeded faults.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2", HEX.teal);
  addTitle(s, "Four uncertainty signals — no ground truth needed");
  const um = [
    ["FaUserCheck", HEX.muted, "Verbalised confidence", "U = 1 − c", "The LLM's own probability that the test is correct. Baseline."],
    ["FaLayerGroup", HEX.blue, "Structural entropy", "U = −log p(cluster) / log N", "Cluster samples by oracle templates, event types and traffic layout — semantic entropy on the artefact " + cite("kuhn23") + "."],
    ["FaCar", HEX.teal, "Behavioural entropy (new)", "U = −log p(behaviour) / log N", "Cluster samples by what they do on the SUT: own verdict, collision, emergency braking, minimum-gap band."],
    ["FaCircleNodes", HEX.purple, "Oracle disagreement", "U = disagreeing peers / (N − 1)", "Apply every peer's oracle to this sample's trace; disagreement means an ambiguous requirement."],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, f, d] = um[i], x = MX + (i % 2) * 6.15, y = 1.45 + Math.floor(i / 2) * 2.3;
    card(s, x, y, 5.95, 2.1, i === 2 ? "E6F3F3" : HEX.mist);
    await badge(s, x + 0.3, y + 0.3, 0.7, ic, col);
    txt(s, h, { x: x + 1.2, y: y + 0.22, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, f, { x: x + 1.2, y: y + 0.68, w: 4.5, h: 0.4, fontSize: 14, fontFace: "Courier New", color: col === HEX.muted ? HEX.ink : col, bold: true });
    txt(s, d, { x: x + 1.2, y: y + 1.08, w: 4.55, h: 0.95, fontSize: 14 });
  }
  txt(s, "Combined score = mean(structural, behavioural); oracle disagreement was uninformative on a development round and is reported only.", { x: MX, y: 6.12, w: W - 2 * MX, h: 0.3, fontSize: 12, color: HEX.muted, italic: true });
  footnotes(s, 6.6);
  s.addNotes("Four signals. Verbalised confidence is the baseline. Structural entropy clusters samples by what they look like. Behavioural entropy, my CPS-specific adaptation of semantic entropy, clusters them by what they do when executed. Oracle disagreement cross-applies the oracles. The design was fixed on a development round and evaluated on a fresh held-out round.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2", HEX.teal);
  addTitle(s, "Demo 2: one requirement, end to end, live LLM");
  const vw2 = 8.3, vh2 = vw2 * 900 / 1440;
  s.addMedia({ type: "video", path: A("demo.mp4"), x: MX, y: 1.4, w: vw2, h: vh2, cover: "image/png;base64," + fs.readFileSync(A("demo_50.png")).toString("base64") });
  card(s, MX + vw2 + 0.3, 1.4, W - 2 * MX - vw2 - 0.3, vh2, HEX.mist);
  txt(s, "What you see", { x: MX + vw2 + 0.55, y: 1.6, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["R6: a cut-out reveals a stopped car", "8 samples from gpt-5.6-luna", "Live stages: generate → gate → simulate → quantify → triage → ground truth", "Most uncertain test = the irrelevant one", "Replay on a seeded fault: the oracle catches it"], 13), { x: MX + vw2 + 0.55, y: 2.1, w: 2.75, h: vh2 - 0.9 });
  txt(s, "1 min 44 s · recorded from the running prototype", { x: MX, y: 1.4 + vh2 + 0.08, w: vw2, h: 0.3, fontSize: 11, color: HEX.muted, italic: true });
  s.addNotes("[Play video, about 1:44.] Requirement R6: the lead car changes lane and reveals a stopped car. Eight tests are sampled live and each stage completes. Five tests are auto-accepted and three go to review. The most uncertain test is the one that does not exercise the requirement. Finally the same test is replayed on a faulty variant and its oracle catches the fault.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2 RESULTS", HEX.teal);
  addTitle(s, "Behavioural entropy beats the LLM's self-confidence");
  txt(s, `${S.n_samples} generated tests · 7 ACC requirements · 3 LLMs × N = 10 · 8 seeded faults · 1,890 simulations · held-out round`, { x: MX, y: 1.3, w: W - 2 * MX, h: 0.3, fontSize: 13, color: HEX.muted, bold: true });
  const labs = ["Verbalised conf.", "Structural", "Behavioural", "Oracle disagr.", "Combined"];
  const ks = ["U_verbal", "U_struct", "U_behav", "U_oracle", "U_comb"];
  s.addChart(pres.charts.BAR, [
    { name: "All tests (pooled)", labels: labs, values: ks.map((k) => +AU[k].toFixed(2)) },
    { name: "Within one requirement × model", labels: labs, values: ks.map((k) => +AW[k].toFixed(2)) }],
  { x: MX, y: 1.7, w: 7.6, h: 4.65, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.teal, HEX.ice],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.ink, dataLabelFormatCode: "0.00",
    valAxisMinVal: 0, valAxisMaxVal: 1, valAxisMajorUnit: 0.25, valAxisLabelFormatCode: "0.00",
    valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    showTitle: true, title: "AUROC for detecting incorrect tests (0.5 = random)", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt" });
  card(s, 8.5, 1.7, 4.23, 4.65, HEX.mist);
  txt(s, f2(AU.U_behav), { x: 8.8, y: 1.85, w: 3.7, h: 0.9, fontSize: 48, bold: true, color: HEX.teal, fontFace: "Cambria" });
  txt(s, "AUROC of behavioural entropy vs. " + f2(AU.U_verbal) + " for the LLM's own confidence", { x: 8.8, y: 2.75, w: 3.7, h: 0.7, fontSize: 14 });
  txt(s, bullets([`Ranks the ${S.n_incorrect} incorrect tests (of ${S.n_samples}) mostly at the top`,
    `Batch level: behavioural entropy vs. share of correct tests, Spearman ρ = ${f2(S.batch_spearman.behav_entropy)}`,
    [{ text: "Honest caveat: ", bold: true }, { text: `within one requirement the signal drops to ${f2(AW.U_behav)} — it mostly flags hard requirements` }]], 13), { x: 8.8, y: 3.55, w: 3.75, h: 2.75 });
  source(s, `Ground truth: a test is correct if it exercises the requirement and its oracle agrees with the reference on the SUT and all 8 faulty versions. Within-batch AUROC averaged over the ${S.n_batches_with_errors} batches with incorrect tests.`, 6.55);
  s.addNotes(`Setup: 210 tests from three ChatGPT-class models on seven ACC requirements, each run on the correct controller and eight seeded faults. On the held-out round behavioural entropy detects incorrect tests with AUROC ${f2(AU.U_behav)}, against ${f2(AU.U_verbal)} for the model's own confidence, and at batch level rho is ${f2(S.batch_spearman.behav_entropy)}. Honestly, within a single requirement it drops to ${f2(AW.U_behav)}: it mostly flags hard requirements. Per-test discrimination is a research gap.`);

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2 RESULTS", HEX.teal);
  addTitle(s, `Risk-based triage: ${pct(cov)} auto-accepted at ${(accPrec * 100).toFixed(1)}% precision`);
  const rcC = S.risk_coverage.U_comb;
  const xs = [], ys = [];
  for (let i = 0; i < rcC.coverage.length; i += 3) { xs.push(+rcC.coverage[i].toFixed(3)); ys.push(+rcC.precision[i].toFixed(3)); }
  xs.push(1); ys.push(+rcC.precision.at(-1).toFixed(3));
  s.addChart(pres.charts.SCATTER, [{ name: "X", values: xs }, { name: "Combined uncertainty", values: ys }],
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
  s.addNotes(`Turning uncertainty into decisions: accepting the least uncertain tests first keeps precision near 100% until about 70% coverage. With fixed thresholds ${nAcc} of ${S.n_samples} tests are auto-accepted and ${T.accept.correct} of them are correct; errors concentrate in the review bucket. Unlike QA-Flow's binary gate, this tells a human where to look. The thresholds are heuristic, which leads to objective 2.`);

  // ======================= PART C: GAPS -> NEXT (O1-O4) =======================
  pres.addSection({ title: "C · Gaps → research plan" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · Gaps → research plan" });
  kicker(s, "PART C · OBJECTIVES", HEX.amber);
  addTitle(s, "The position's four objectives, as research questions");
  const obj = [
    ["O1", "FaGaugeHigh", HEX.teal, "New UQ metrics", "Which signals reveal that an AI-generated test artefact is untrustworthy — for black-box and white-box models?"],
    ["O2", "FaScaleBalanced", HEX.blue, "Risk & trustworthiness", "How should quantified uncertainty drive accept / review / reject decisions under safety risk?"],
    ["O3", "FaChartLine", HEX.amber, "Uncertainty vs. quality", "How does uncertainty relate to correctness, fault-detection power and cost of the generated tests?"],
    ["O4", "FaAtom", HEX.purple, "Future AI models", "Do the metrics transfer to newer large models and to quantum machine learning models?"],
  ];
  for (let i = 0; i < 4; i++) {
    const [k, ic, col, h, d] = obj[i], x = MX + i * 3.08;
    card(s, x, 1.5, 2.88, 4.3, HEX.mist);
    await badge(s, x + 0.3, 1.8, 0.75, ic, col);
    txt(s, k, { x: x + 1.2, y: 1.88, w: 1.4, h: 0.6, fontSize: 28, bold: true, color: col, fontFace: "Calibri" });
    txt(s, h, { x: x + 0.3, y: 2.8, w: 2.4, h: 0.7, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, d, { x: x + 0.3, y: 3.55, w: 2.35, h: 2.1, fontSize: 14 });
  }
  cite("innoguard");
  footnotes(s, 6.3);
  s.addNotes("This is how I read the four objectives of the position, phrased as questions. The gaps from both demos map onto them.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · Gaps → research plan" });
  kicker(s, "PART C · GAPS → NEXT IMPLEMENTATION", HEX.amber);
  addTitle(s, "What both demos cannot do yet becomes the research agenda");
  const gaps = [
    ["Demo 1", `Rules cannot tell a wrong test from a real defect (${nWrong} of ${nFail} FAILs were wrong tests)`, "Uncertainty per generated artefact, incl. oracles and bug reports", "O1", HEX.teal],
    ["Demo 2", `Per-test signal weak within a requirement (AUROC ${f2(AW.U_behav)}); black-box only`, "White-box (log-prob) and trajectory-level metrics; calibration", "O1", HEX.teal],
    ["Demo 2", "Thresholds are heuristic; risk ignores the severity of the requirement", "Risk = uncertainty × consequence; selective guarantees; review budget", "O2", HEX.blue],
    ["Both", `Uncertainty vs. effectiveness only hinted (ρ = ${f2(S.spearman_ms.U_comb)} with mutation score); cost of N samples`, "Large study: correctness, fault detection, flakiness, latency and token cost", "O3", HEX.amber],
    ["Both", "One app or one simulator, 3 hosted LLMs, seeded faults", "Open-source ACPS, industrial cases, InnoGuard platform; newer and quantum ML models", "O4", HEX.purple],
  ];
  const gh = ["From", "Gap seen in the demo", "Next implementation", "Obj."].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([gh, ...gaps.map((g, i) => [
    { text: g[0], options: { fontSize: 13, bold: true, color: g[0] === "Demo 1" ? HEX.indigo : g[0] === "Demo 2" ? HEX.teal : HEX.muted, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[1], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[2], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[3], options: { fontSize: 18, bold: true, color: g[4], align: "center", fill: { color: i % 2 ? HEX.white : HEX.mist } } }])],
  { x: MX, y: 1.45, w: W - 2 * MX, colW: [1.05, 5.0, 5.08, 1.0], rowH: 0.9, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.1, valign: "middle" });
  s.addNotes("Each limitation maps onto an objective. From Demo 1: rules cannot tell a wrong test from a real defect, which calls for uncertainty on every generated artefact. From Demo 2: the per-test signal is weak and black-box only — new metrics, O1. Heuristic thresholds that ignore severity — risk methods, O2. The relation between uncertainty, effectiveness and cost — a large study, O3. And generalisation beyond one app, one simulator and hosted models — O4.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · Gaps → research plan" });
  kicker(s, "PART C · RESEARCH QUESTIONS", HEX.amber);
  addTitle(s, "Research questions and how I will answer them");
  const rqs = [
    ["RQ1", HEX.teal, "Which uncertainty metrics best predict the correctness of AI-generated test artefacts?", "Metric suite: behavioural / semantic / oracle / white-box; AUROC, calibration error; LLMs and VLMs as generators and judges"],
    ["RQ2", HEX.blue, "How can uncertainty support risk-based acceptance of generated artefacts?", "Risk model linking uncertainty to hazard severity; selective-prediction guarantees; review-effort savings in user studies"],
    ["RQ3", HEX.amber, "How does uncertainty relate to effectiveness and non-functional properties?", "Mutation score, failure severity, flakiness, execution time and token cost; mixed-effects statistics"],
    ["RQ4", HEX.purple, "Do the metrics transfer to advanced and quantum ML models?", "Re-apply metrics to newer foundation models and to QML-based oracles/predictors " + cite("wang25ase", "wang24qelm")],
  ];
  for (let i = 0; i < 4; i++) {
    const [k, col, q, m] = rqs[i], y = 1.4 + i * 1.2;
    card(s, MX, y, W - 2 * MX, 1.06, HEX.mist);
    txt(s, k, { x: MX + 0.25, y: y + 0.15, w: 1.0, h: 0.75, fontSize: 24, bold: true, color: col, fontFace: "Cambria", valign: "middle" });
    txt(s, q, { x: MX + 1.35, y: y + 0.09, w: 5.0, h: 0.88, fontSize: 15, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, m, { x: MX + 6.55, y: y + 0.09, w: 5.35, h: 0.88, fontSize: 13, valign: "middle" });
  }
  footnotes(s, 6.3);
  s.addNotes("Four research questions, one per objective, each with its method, evaluated on open-source ACPS, industrial cases through the network such as the Mondragon elevator systems, and the InnoGuard platform, with fault-injection or expert ground truth.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · Gaps → research plan" });
  kicker(s, "PART C · WORK PLAN", HEX.amber);
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
  s.addNotes("A 36-month plan. Year 1: a mapping study, an open benchmark that includes both QA-Flow-style web/API artefacts and ACPS tests, and the first metric suite. Year 2: the empirical study, risk methods and the planned secondment at Mondragon. Year 3: industrial and InnoGuard-platform evaluation, the quantum ML extension, and the thesis. Secondment dates are indicative.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · Gaps → research plan" });
  kicker(s, "PART C · CONTRIBUTIONS", HEX.amber);
  addTitle(s, "Expected contributions to InnoGuard");
  const contrib = [
    ["FaGaugeHigh", HEX.teal, "Metrics", "Uncertainty metrics for AI-generated test artefacts, incl. behavioural entropy for executable CPS tests"],
    ["FaScaleBalanced", HEX.blue, "Methods", "Risk-based acceptance of AI-generated artefacts with calibrated, safety-aware thresholds"],
    ["FaChartLine", HEX.amber, "Evidence", "Empirical laws linking uncertainty to correctness, fault detection and cost across models"],
    ["FaRocket", HEX.purple, "Open science", "Open benchmark, both tools and replication packages; integration with the InnoGuard platform"],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, d] = contrib[i], y = 1.45 + i * 1.08;
    await badge(s, MX, y + 0.12, 0.7, ic, col);
    txt(s, h, { x: MX + 0.95, y: y + 0.08, w: 2.0, h: 0.8, fontSize: 18, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, d, { x: MX + 2.9, y: y + 0.08, w: 4.4, h: 0.8, fontSize: 14, valign: "middle" });
  }
  card(s, 8.1, 1.45, 4.63, 4.3, HEX.mist);
  txt(s, "Dissemination & network", { x: 8.35, y: 1.63, w: 4.2, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["Venues: ICSE, FSE, ASE, ISSTA, ICST; TOSEM, TSE, EMSE", "Joint papers across the network, building on the team's foundation-model road maps " + cite("lu26", "ali25isola"), "Secondment at Mondragon: industrial CPS case and oracle expertise", "Training, summer schools and outreach of the doctoral network"], 14), { x: 8.35, y: 2.1, w: 4.2, h: 3.5 });
  footnotes(s, 6.2);
  s.addNotes("Four contributions — metrics, risk methods, empirical evidence and open artefacts — published at the main SE venues and developed jointly across the network, with the Mondragon secondment as the industrial anchor.");

  // ======================= PART D: WHAT I CAN CONTRIBUTE =======================
  pres.addSection({ title: "D · What I can contribute" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "D · What I can contribute" });
  kicker(s, "PART D · MY BACKGROUND", HEX.green);
  addTitle(s, "From automotive test benches to AI-driven testing");
  const tl = [
    ["2019–23", "B.S. Electrical & Electronics Eng.", "HCMC Univ. of Technology, honors program · GPA 8.27/10", "FaMicrochip", HEX.muted],
    ["2022", "Embedded SW intern", "Drivers, firmware unit & functional testing", "FaGears", HEX.muted],
    ["2022–24", "Embedded SW test engineer · Bosch", "Automotive ECU testing for OEMs · AUTOSAR, MISRA-C, ASPICE, CAPL", "FaCar", HEX.blue],
    ["2024–26", "M.S. Computer Science · KNU", "Software Safety Eng. Lab · GPA 4.25/4.30 · falsification & property-based testing", "FaGraduationCap", HEX.teal],
    ["2026–", "Software test engineer · industry", "Web/API/system testing · LLM, RAG & MCP-based test generation", "FaRobot", HEX.amber],
  ];
  const ty = 2.0, tw = (W - 2 * MX) / tl.length;
  s.addShape(pres.shapes.LINE, { x: MX + tw / 2, y: ty + 0.32, w: tw * (tl.length - 1), h: 0, line: { color: HEX.line, width: 2 } });
  for (let i = 0; i < tl.length; i++) {
    const [yr, role, det, ic, col] = tl[i], cx = MX + i * tw;
    await badge(s, cx + tw / 2 - 0.32, ty, 0.64, ic, col);
    txt(s, yr, { x: cx, y: ty + 0.78, w: tw, h: 0.35, fontSize: 14, bold: true, color: col, align: "center" });
    txt(s, role, { x: cx + 0.08, y: ty + 1.12, w: tw - 0.16, h: 0.62, fontSize: 14, bold: true, align: "center", valign: "top" });
    txt(s, det, { x: cx + 0.1, y: ty + 1.74, w: tw - 0.2, h: 0.95, fontSize: 12, color: HEX.muted, align: "center" });
  }
  card(s, MX, 5.05, W - 2 * MX, 1.35, HEX.mist);
  txt(s, [{ text: "Thesis  ", options: { bold: true, color: HEX.teal } }, { text: "Towards Efficient Function Abstraction via Active Learning for Property Falsification (advisor: Prof. Yunja Choi)", options: { breakLine: true } },
    { text: "Papers  ", options: { bold: true, color: HEX.teal } }, { text: "Test generation for property checking (KCC 2025); stateful system, stateless techniques: test generation for controller software (KCSE 2026), with S. Hong and Y. Choi" }],
    { x: MX + 0.3, y: 5.23, w: W - 2 * MX - 0.6, h: 1.0, fontSize: 14, paraSpaceAfter: 6, valign: "middle" });
  s.addNotes("My path: from hardware to automotive software testing at Bosch, then research on falsification and test generation for controller software at Kyungpook National University, and now LLM-based testing in industry — which is where QA-Flow comes from. The common thread is making automated testing of safety-relevant software trustworthy.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "D · What I can contribute" });
  kicker(s, "PART D · WHAT I BRING", HEX.green);
  addTitle(s, "What I bring, and why this team");
  const fit = [
    ["FaCar", HEX.blue, "Testing cyber-physical systems", "Automotive ECU testing (CAN, CAPL, ETAS INCA); robotics (ROS/ROS2) and embedded domains"],
    ["FaWaveSquare", HEX.teal, "Verification research", "Falsification, function abstraction via active learning, property-based testing — close to robustness-guided falsification " + cite("zhang23falsifai")],
    ["FaBrain", HEX.amber, "LLMs for quality engineering", "Built QA-Flow and UQ-TestGen: RAG, LLM agents, MCP-based workflows, Playwright"],
    ["FaFlask", HEX.purple, "Python & empirical studies", "Controlled experiments on test generators (KCC'25, KCSE'26); held-out evaluation in this talk"],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, d] = fit[i], x = MX + (i % 2) * 4.1, y = 1.45 + Math.floor(i / 2) * 2.25;
    card(s, x, y, 3.95, 2.1, HEX.mist);
    await badge(s, x + 0.25, y + 0.25, 0.6, ic, col);
    txt(s, h, { x: x + 1.0, y: y + 0.22, w: 2.8, h: 0.65, fontSize: 15, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, d, { x: x + 0.25, y: y + 0.98, w: 3.5, h: 1.05, fontSize: 13 });
  }
  card(s, 8.85, 1.45, 3.88, 4.35, HEX.navy);
  txt(s, "Why this team", { x: 9.1, y: 1.65, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: HEX.white });
  txt(s, bullets(["Uncertainty-wise CPS testing (Simula)", "ADS testing & falsification (NII)", "Industrial CPS oracles & elevators (Mondragon)", "My demos sit exactly where the three meet"], 14, HEX.ice), { x: 9.1, y: 2.15, w: 3.45, h: 3.5 });
  footnotes(s, 6.3);
  s.addNotes("What I bring maps onto the preferred profile: CPS testing practice, verification research, hands-on LLM tooling — both tools in this talk — and empirical studies. And why this team: uncertainty-wise CPS testing at Simula, ADS testing and falsification with Prof. Arcaini, industrial oracles at Mondragon. My two demos sit exactly where these meet.");

  s = newSlide({ masterName: "DARK", sectionTitle: "D · What I can contribute" });
  addTitle(s, "Thank you");
  txt(s, bullets(["Demo 1 · QA-Flow: LLM tests end to end, trusted by rules — and fooled by wrong tests", "Demo 2 · UQ-TestGen: uncertainty tells which generated tests to trust", "PhD: metrics, risk methods, evidence and transfer — O1 to O4"], 18, HEX.ice), { x: MX, y: 3.4, w: 9.5, h: 1.9 });
  txt(s, "zavinhle@gmail.com", { x: MX, y: 5.45, w: 6, h: 0.4, fontSize: 16, color: HEX.white, bold: true });
  s.addImage({ path: A("strip.png"), x: 0, y: H - 1.33, w: W, h: W * 160 / 1600 });
  s.addNotes("To summarise: QA-Flow automates testing end to end but trusts by rules, and the rules are fooled by wrong tests; UQ-TestGen shows uncertainty can tell which generated tests to trust; and the PhD turns that into metrics, risk methods, evidence and transfer, along O1 to O4. Thank you — I look forward to your questions.");

  // ======================= BACKUP =======================
  pres.addSection({ title: "Backup" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · QA-Flow demo: planted defects, faults and gates");
  const dh = ["Planted defect (ShopLite)", "Type", "Found by test(s)"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 13 } }));
  const allDef = [["W1", "web"], ["A1", "API"], ["A2", "API"], ["S1", "security"], ["S2", "security"], ["S3", "security"], ["S4", "security"]];
  s.addTable([dh, ...allDef.map(([id, kind], i) => { const f = G.defectsFound.find((d) => d.id === id); return [QF.defectNames[id], kind, f ? f.tests.join(", ") : "not found"].map((c, j) => ({ text: c, options: { fontSize: 12, color: j === 2 && !f ? HEX.red : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })); })],
    { x: MX, y: 1.45, w: 7.3, colW: [3.6, 1.2, 2.5], rowH: 0.42, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.08, valign: "middle" });
  card(s, 8.2, 1.45, 4.53, 4.9, HEX.mist);
  txt(s, "Run details", { x: 8.45, y: 1.6, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets([`${Q.requirements} requirements, ${Q.tests} tests (${Q.tests - Q.byKindRule} LLM-designed, ${Q.byKindRule} rule-derived)`,
    `Schema gate: ${Q.gate.replies} agent replies, ${Q.gate.rejected} rejected, ${Q.gate.repaired} repaired`, `Grounding gate: ${Q.gate.ungrounded} tests ungrounded → NEEDS REVIEW`,
    `Fault gate: ${QF.faults} broken builds; ${Q.gate.cannotFail} passing tests never failed → NEEDS REVIEW`, `Wrong tests among FAILs: ${G.wrongTests.join(", ")}`, `Total run time ${Math.round(Q.seconds / 60)} min (gpt-5.6-luna)`], 13), { x: 8.45, y: 2.05, w: 4.1, h: 4.2 });
  s.addNotes("Details of the QA-Flow run behind Demo 1: which planted defects were found by which tests, and how many outputs each trust rule caught.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · UQ-TestGen results per model and requirement");
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
  s.addNotes("Per model: gpt-6-astra produced only correct tests, so uncertainty matters most for cheaper models. Per requirement: cut-in (R3) and cut-out (R6) are hardest and also have the highest uncertainty.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · Threats to validity and mitigations");
  const thr = [["Construct", "Correctness = agreement with my reference formalisation; mutants and planted defects approximate real faults", "Expert-validated oracles; real fault data from industrial partners"],
    ["Internal", "UQ metric design chosen on a development round; heuristic thresholds", "Held-out evaluation; pre-registered designs; calibration on separate data"],
    ["External", "One web app, one SUT, one simulator, 3 hosted LLMs", "Open-source ACPS, industrial cases, InnoGuard platform, open-weight models"],
    ["Conclusion", `${S.n_incorrect} incorrect tests in ${S.n_batches_with_errors} batches; one QA-Flow run`, "Larger N, repeated runs, effect sizes and confidence intervals"]];
  const th = ["Threat", "In the demos", "Mitigation in the PhD"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([th, ...thr.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 0 ? HEX.navy : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: MX, y: 1.5, w: W - 2 * MX, colW: [1.8, 5.2, 5.13], rowH: 1.0, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.12, valign: "middle" });
  s.addNotes("Main threats: my reference formalisation and synthetic faults define ground truth; the metric design was chosen on a development round; and the scale is small. Each has a mitigation in the plan.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · Towards uncertainty of quantum ML models (O4)");
  txt(s, bullets([[{ text: "Starting point: ", bold: true }, { text: "QML is already used for test oracles of autonomous mobile robots and QoS prediction of elevator software " + cite("wang25ase", "wang24qelm") }],
    [{ text: "Sources of uncertainty: ", bold: true }, { text: "shot noise, hardware noise and model (epistemic) uncertainty add to data uncertainty" }],
    [{ text: "Transfer test: ", bold: true }, { text: "do sampling-based metrics still predict wrong oracle verdicts when the model is quantum? Repeated shots are natural samples" }],
    [{ text: "Coverage & mutation for quantum programs ", bold: true }, { text: "as ground-truth tools " + cite("ali21icst") }],
    [{ text: "Road map: ", bold: true }, { text: "quantum AI for software engineering " + cite("wang26qai") }]], 16), { x: MX, y: 1.5, w: 7.4, h: 4.4 });
  await badge(s, 9.4, 2.1, 2.4, "FaAtom", HEX.purple);
  footnotes(s, 6.1);
  s.addNotes("For objective 4: shot-based sampling is a natural analogue to sampling an LLM, so the same disagreement-based metrics can be tested, with quantum noise as an extra source of uncertainty.");

  const half = Math.ceil(REFS.length / 2);
  for (let p = 0; p < 2; p++) {
    s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
    addTitle(s, `References (${p + 1}/2)`);
    const part = REFS.slice(p * half, (p + 1) * half);
    txt(s, part.map(([, r], i) => ({ text: `[${p * half + i + 1}] ${r}`, options: { fontSize: 10, paraSpaceAfter: 3, breakLine: i < part.length - 1 } })),
      { x: MX, y: 1.35, w: W - 2 * MX, h: 5.55, valign: "top" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
