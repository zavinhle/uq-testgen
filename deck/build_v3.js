// Builds the v3 interview deck (v2 with a simplified Demo 2: 'ask 10 times, let the tests vote')
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
const OUT = path.join(__dirname, "PhD_Proposal_ZaVinhLe_v3.pptx");

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
  addTitle(s, "UQ-TestGen: ask the AI 10 times, let the tests vote");
  const st = [["FaFileLines", "1 · Requirement", "one sentence in plain English", HEX.muted],
    ["FaDice", "2 · Ask 10 times", "the AI writes the same test 10 times, independently", HEX.blue],
    ["FaCar", "3 · Run all 10", "in a driving simulator, on a cruise-control car", HEX.blue],
    ["FaPeopleGroup", "4 · Group them", "tests that behaved the same go into one group", HEX.teal],
    ["FaScaleBalanced", "5 · Vote", "big group → accept · small → human review · alone → reject", HEX.teal]];
  const sw = (W - 2 * MX) / st.length;
  for (let i = 0; i < st.length; i++) {
    const [ic, h, d, col] = st[i], x = MX + i * sw;
    card(s, x + 0.06, 1.45, sw - 0.12, 2.3, i >= 3 ? "E6F3F3" : HEX.mist);
    await badge(s, x + sw / 2 - 0.35, 1.62, 0.7, ic, col);
    txt(s, h, { x: x + 0.12, y: 2.45, w: sw - 0.24, h: 0.4, fontSize: 17, bold: true, align: "center", color: HEX.navy });
    txt(s, d, { x: x + 0.18, y: 2.88, w: sw - 0.36, h: 0.8, fontSize: 13, align: "center", color: HEX.ink });
  }
  s.addImage({ path: A("frame_c.png"), x: MX, y: 4.0, w: 6.6, h: 6.6 * 260 / 1600 });
  txt(s, "The car under test (green) and the traffic the AI scripts (blue) in highway-env " + cite("leurent18") + ".", { x: MX, y: 5.15, w: 6.6, h: 0.5, fontSize: 12, color: HEX.muted, italic: true });
  card(s, 7.5, 4.0, 5.23, 2.15, HEX.navy);
  txt(s, [{ text: "Why it works  ", options: { bold: true, color: HEX.amber } },
    { text: "If the AI really knows how to test the requirement, its 10 tests behave alike. If it is guessing, they scatter.", options: { color: HEX.white, breakLine: true } },
    { text: "Same idea as semantic entropy for chatbots " + cite("kuhn23", "farquhar24") + " — here “same meaning” = “same behaviour in the simulator”.", options: { color: HEX.ice } }],
  { x: 7.75, y: 4.15, w: 4.8, h: 1.9, fontSize: 14, paraSpaceAfter: 6 });
  footnotes(s, 6.3);
  s.addNotes("Here is the whole idea of Demo 2 in five steps. One: a requirement in plain English. Two: instead of asking the AI once, I ask it for the same test ten times, in independent sessions. Three: I run all ten tests in a driving simulator, against a cruise-control car. Four: I group the tests by what actually happened, for example 'no crash and a big gap', or 'crash'. Five: the groups vote. If most tests behaved the same, the AI is consistent and I accept them. If a test behaves unlike the others, a human looks at it, and a test that is completely alone is rejected. Why does this work? If the AI really knows how to test something, its answers agree. If it is guessing, they scatter. This is the same idea as semantic entropy for chatbots; my twist is that for a test, the meaning is what it does when it runs.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2 · HOW THE VOTE WORKS", HEX.teal);
  addTitle(s, "Example: 10 tests for one requirement, 5 behaviours");
  txt(s, [{ text: "Requirement R3 · A car cuts in:  ", options: { bold: true } }, { text: "“If a car from the next lane cuts in at least 15 metres in front of us, our car must not hit it, and 6 seconds later it must again keep a safe distance: at least 28 metres at 100 km/h.”" },
    { text: "   (gpt-5.6-terra, 10 tests)", options: { color: HEX.muted } }],
    { x: MX, y: 1.22, w: W - 2 * MX, h: 0.58, fontSize: 14, color: HEX.ink, valign: "middle" });
  const VC = { x: MX, dots: 1.15, lab: 3.4, agr: 7.75, dec: 9.6, tru: 11.05 };
  const head = [["Tests", VC.dots, 2.1], ["What happened in the simulator", VC.lab, 4.2], ["Agreement", VC.agr, 1.8], ["Decision", VC.dec, 1.4], ["Truth (checked later)", VC.tru, 1.7]];
  for (const [t, x, w] of head) txt(s, t, { x, y: 1.85, w, h: 0.3, fontSize: 12, bold: true, color: HEX.muted });
  const DEC = { accept: ["Accept", HEX.green, "E7F4EC"], review: ["Human review", HEX.amber, "FFF4E0"], reject: ["Reject", HEX.red, "FDECEA"] };
  const vote = [["A", 4, "Drives safely, keeps a big gap", "review", "2 ok · 2 wrong", HEX.blue],
    ["B", 3, "Drives safely, gets closer", "review", "3 ok", "7C3AED"],
    ["C", 1, "Emergency brake, then crash", "reject", "wrong", "0891B2"],
    ["D", 1, "Crash", "reject", "wrong", "DB2777"],
    ["E", 1, "Emergency brake, no crash", "reject", "ok (unusual)", "EA580C"]];
  for (let i = 0; i < vote.length; i++) {
    const [g, k, lab, dec, truth, col] = vote[i], y = 2.22 + i * 0.66;
    card(s, MX, y, W - 2 * MX, 0.56, i % 2 ? HEX.white : HEX.mist);
    s.addShape(pres.shapes.OVAL, { x: MX + 0.12, y: y + 0.1, w: 0.36, h: 0.36, fill: { color: col }, line: { color: col } });
    txt(s, g, { x: MX + 0.12, y: y + 0.1, w: 0.36, h: 0.36, fontSize: 13, bold: true, color: HEX.white, align: "center", valign: "middle" });
    for (let d = 0; d < k; d++) s.addShape(pres.shapes.OVAL, { x: VC.dots + d * 0.27, y: y + 0.17, w: 0.22, h: 0.22, fill: { color: col }, line: { color: col } });
    txt(s, lab, { x: VC.lab, y, w: 4.2, h: 0.56, fontSize: 15, valign: "middle" });
    txt(s, `${k} of 10 = ${k * 10}%`, { x: VC.agr, y, w: 1.8, h: 0.56, fontSize: 15, bold: true, valign: "middle", color: DEC[dec][1] });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: VC.dec, y: y + 0.11, w: 1.3, h: 0.34, rectRadius: 0.08, fill: { color: DEC[dec][2] }, line: { color: DEC[dec][2] } });
    txt(s, DEC[dec][0], { x: VC.dec, y: y + 0.11, w: 1.3, h: 0.34, fontSize: 12, bold: true, color: DEC[dec][1], align: "center", valign: "middle" });
    txt(s, truth, { x: VC.tru, y, w: 1.7, h: 0.56, fontSize: 13, valign: "middle", color: truth.startsWith("wrong") || truth.includes("wrong") ? HEX.red : HEX.green });
  }
  card(s, MX, 5.6, W - 2 * MX, 0.5, "E6F3F3");
  txt(s, [{ text: "Vote rule:  ", options: { bold: true, color: HEX.teal } }, { text: "≥ 60% agree → accept   ·   20–60% → human review   ·   below 20% (alone) → reject" }],
    { x: MX + 0.25, y: 5.6, w: W - 2 * MX - 0.5, h: 0.5, fontSize: 15, valign: "middle" });
  txt(s, "No group reached 60%: the AI disagrees with itself, so nothing is auto-accepted. Both crash tests are rejected — and they really were wrong.",
    { x: MX, y: 6.25, w: W - 2 * MX, h: 0.5, fontSize: 14, bold: true, color: HEX.navy });
  s.addNotes("One real example. The requirement: a car cuts in ahead; we must not crash, and six seconds later we must again keep a safe distance, about 28 metres at 100 km/h. The AI wrote ten tests. I ran them and grouped them by what happened. Four tests: the car drives safely and keeps a big gap. Three: safe, but closer. Then three tests that are each alone: emergency brake then crash, a crash, and an emergency brake without a crash. The vote: no group has 60 percent, so the AI is not consistent here and nothing is accepted automatically. The two big groups go to a human, and the three lonely tests are rejected. The last column is the truth, which I checked afterwards with deliberately broken controllers: both crash tests really were wrong. It is not perfect: group A still hides two wrong tests, which a human would catch in review, and the emergency-brake test was correct but rejected because it was unusual. That is exactly where my PhD starts.");

  const D2 = fs.existsSync(A("demo2.json")) ? JSON.parse(fs.readFileSync(A("demo2.json"), "utf8")) : null;
  const vidFile = D2 ? "demo2.mp4" : "demo.mp4", posterFile = D2 ? "demo2_poster.png" : "demo_50.png";
  const vidLen = D2 ? `${Math.floor(D2.seconds / 60)} min ${String(Math.round(D2.seconds % 60)).padStart(2, "0")} s` : "1 min 44 s";
  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2", HEX.teal);
  addTitle(s, "Demo 2: one requirement, end to end, live AI");
  const vw2 = 8.3, vh2 = vw2 * 900 / 1440;
  s.addMedia({ type: "video", path: A(vidFile), x: MX, y: 1.4, w: vw2, h: vh2, cover: "image/png;base64," + fs.readFileSync(A(posterFile)).toString("base64") });
  card(s, MX + vw2 + 0.3, 1.4, W - 2 * MX - vw2 - 0.3, vh2, HEX.mist);
  txt(s, "What you see", { x: MX + vw2 + 0.55, y: 1.6, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["R6 · A hidden stopped car: the car ahead changes lane, and a stopped car is behind it", `The AI writes ${D2 ? D2.n : 8} tests live`, "The simulator runs them; tests are grouped by what happened",
    D2 ? `Vote: ${D2.accept} accept · ${D2.review} review · ${D2.reject} reject` : "The vote: accept / review / reject", "Replay on a broken controller: the test catches the bug"], 13), { x: MX + vw2 + 0.55, y: 2.1, w: 2.75, h: vh2 - 0.9 });
  txt(s, `${vidLen} · recorded from the running prototype`, { x: MX, y: 1.4 + vh2 + 0.08, w: vw2, h: 0.3, fontSize: 11, color: HEX.muted, italic: true });
  s.addNotes(D2 ? D2.notes : "[Play video.] Requirement R6: the car ahead changes lane and reveals a stopped car. Eight tests are written live, run in the simulator and grouped by behaviour; the vote decides. Finally a test is replayed on a broken controller and catches the bug.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2 RESULTS", HEX.teal);
  addTitle(s, "Agreement spots wrong tests better than AI self-confidence");
  txt(s, `${S.n_samples} AI-written tests · 7 cruise-control requirements · 3 AI models × 10 tries · 8 broken controllers · 1,890 simulations`, { x: MX, y: 1.3, w: W - 2 * MX, h: 0.3, fontSize: 13, color: HEX.muted, bold: true });
  const bx0 = 3.9, bw = 3.9;
  const barsD = [["Asking the AI how sure it is", AU.U_verbal, HEX.muted], ["Checking if its 10 tests agree (ours)", AU.U_disagree, HEX.teal]];
  for (let i = 0; i < 2; i++) {
    const [lab, v, col] = barsD[i], y = 2.2 + i * 1.45;
    txt(s, lab, { x: MX, y, w: bx0 - MX - 0.2, h: 0.9, fontSize: 16, bold: i === 1, align: "right", valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: bx0, y: y + 0.1, w: bw * v, h: 0.7, fill: { color: col }, line: { color: col } });
    txt(s, f2(v), { x: bx0 + bw * v + 0.1, y, w: 1.0, h: 0.9, fontSize: 30, bold: true, valign: "middle", color: col === HEX.muted ? HEX.ink : HEX.teal, fontFace: "Cambria" });
  }
  s.addShape(pres.shapes.LINE, { x: bx0 + bw * 0.5, y: 1.95, w: 0, h: 2.95, line: { color: HEX.muted, width: 1, dashType: "dash" } });
  txt(s, "coin flip = 0.50", { x: bx0 + bw * 0.5 - 0.8, y: 4.95, w: 1.6, h: 0.3, fontSize: 11, color: HEX.muted, align: "center" });
  txt(s, "Score: AUROC (0.5 = coin flip, 1.0 = perfect)", { x: bx0, y: 5.3, w: 4.5, h: 0.3, fontSize: 12, color: HEX.muted, italic: true });
  card(s, 8.75, 1.85, 3.98, 4.45, HEX.mist);
  txt(s, "What 0.90 means", { x: 9.0, y: 2.0, w: 3.5, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, "Pick one wrong test and one correct test at random. In 9 of 10 cases the wrong one has the lower agreement.", { x: 9.0, y: 2.45, w: 3.5, h: 1.3, fontSize: 14 });
  txt(s, [{ text: "Honest limit  ", options: { bold: true, color: HEX.amber } }, { text: `comparing tests of the same requirement it drops to ${f2(AW.U_disagree)}: it mostly spots hard requirements. → objective O1` }],
    { x: 9.0, y: 3.85, w: 3.5, h: 1.4, fontSize: 14 });
  txt(s, `The AI also over-trusts itself: gpt-5.6-luna said ${pct(luna.mean_verbal_conf)} sure, but only ${pct(luna.correct_rate)} of its tests were correct.`, { x: 9.0, y: 5.25, w: 3.5, h: 0.95, fontSize: 12, color: HEX.muted });
  source(s, "Truth = each test also run on 8 deliberately broken controllers and compared with a hidden reference check. Design fixed on a first round; numbers from a fresh second round.", 6.55);
  s.addNotes(`Does the vote actually work? I generated ${S.n_samples} tests with three AI models for seven cruise-control requirements, and checked every test against eight deliberately broken controllers to know which tests were really correct. Then I asked: which signal spots the wrong tests? Asking the AI how sure it is scores ${f2(AU.U_verbal)}. Checking whether its ten tests agree scores ${f2(AU.U_disagree)}. In plain words: if you pick one wrong and one correct test at random, nine times out of ten the wrong one has the lower agreement. Honest limit: inside a single requirement the score drops to ${f2(AW.U_disagree)}, so today it mostly tells you which requirements are hard. Making it sharper per test is my first objective.`);

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · Motivation 2 → Demo 2" });
  kicker(s, "PART B · DEMO 2 RESULTS", HEX.teal);
  addTitle(s, `${nAcc} of ${S.n_samples} tests auto-accepted — ${T.accept.correct} of them correct`);
  const tk = ["accept", "review", "reject"];
  s.addChart(pres.charts.BAR, [
    { name: "Test was correct", labels: ["Accept", "Human review", "Reject"], values: tk.map((k) => T[k].correct) },
    { name: "Test was wrong", labels: ["Accept", "Human review", "Reject"], values: tk.map((k) => T[k].n - T[k].correct) }],
  { x: MX, y: 1.45, w: 6.4, h: 4.9, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55, chartColors: [HEX.green, HEX.red],
    showValue: false,
    valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted,
    catAxisLabelFontSize: 13, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    showTitle: true, title: `What the vote did with all ${S.n_samples} tests`, titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt" });
  const tri = [["accept", "Accept", "≥ 60% agree", HEX.green, "E7F4EC"], ["review", "Human review", "20–60% agree", HEX.amber, "FFF4E0"], ["reject", "Reject", "below 20% (alone)", HEX.red, "FDECEA"]];
  for (let i = 0; i < 3; i++) {
    const [k, lab, rule, col, bg] = tri[i], y = 1.45 + i * 1.32;
    card(s, 7.35, y, 5.38, 1.18, bg);
    txt(s, String(T[k].n), { x: 7.6, y: y + 0.12, w: 1.2, h: 0.9, fontSize: 40, bold: true, color: col, fontFace: "Cambria", valign: "middle" });
    txt(s, [{ text: `${lab}  `, options: { bold: true, fontSize: 16, color: HEX.navy } }, { text: rule, options: { fontSize: 13, color: HEX.muted, breakLine: true } },
      { text: `${T[k].correct} correct (${pct(T[k].precision)})${k === "review" ? " — the mistakes gather here" : ""}`, options: { fontSize: 14 } }],
    { x: 8.85, y: y + 0.15, w: 3.75, h: 0.9, valign: "middle" });
  }
  card(s, 7.35, 5.45, 5.38, 0.9, HEX.mist);
  txt(s, `Accepting everything blindly: ${pct(O.correct_rate_valid)} correct. The vote lifts it to ${(accPrec * 100).toFixed(1)}% and tells a human where to look.`, { x: 7.6, y: 5.5, w: 4.95, h: 0.8, fontSize: 13, valign: "middle" });
  source(s, "The 60% / 20% thresholds are hand-picked; linking them to safety risk is objective O2.");
  s.addNotes(`What does the vote decide for all ${S.n_samples} tests? ${nAcc} are accepted automatically, and ${T.accept.correct} of those are really correct, that is ${(accPrec * 100).toFixed(1)} percent. ${T.review.n} go to a human, and this is where the mistakes gather: ${T.review.n - T.review.correct} of them are wrong. ${T.reject.n} are rejected. If you simply trusted every test you would get ${pct(O.correct_rate_valid)} correct, so the vote is better and, more importantly, it tells the human which tests to look at. The thresholds are hand-picked, and connecting them to safety risk is my second objective. Cost: about ${O.mean_latency_s.toFixed(0)} seconds and ${Math.round(O.mean_tokens / 100) * 100} tokens per AI answer, and about one second to simulate a test on nine controller versions.`);

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
    ["Demo 2", `The vote is coarse inside one requirement (AUROC ${f2(AW.U_disagree)}); black-box only`, "White-box (log-prob) and trajectory-level metrics; calibration", "O1", HEX.teal],
    ["Demo 2", "60% / 20% thresholds are hand-picked; risk ignores how severe a failure is", "Risk = uncertainty × consequence; selective guarantees; review budget", "O2", HEX.blue],
    ["Both", `Agreement vs. bug-finding power only hinted (ρ = ${f2(S.spearman_ms.U_disagree)} with mutation score); cost of 10 samples`, "Large study: correctness, fault detection, flakiness, latency and token cost", "O3", HEX.amber],
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
    ["FaGaugeHigh", HEX.teal, "Metrics", "Uncertainty metrics for AI-generated test artefacts, starting from behaviour agreement (behavioural entropy) for executable CPS tests"],
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
  txt(s, bullets(["Demo 1 · QA-Flow: LLM tests end to end, trusted by rules — and fooled by wrong tests", "Demo 2 · UQ-TestGen: when the AI's tests agree, trust them; when they don't, a human checks", "PhD: metrics, risk methods, evidence and transfer — O1 to O4"], 18, HEX.ice), { x: MX, y: 3.4, w: 9.5, h: 1.9 });
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
  addTitle(s, "Backup · The four research signals behind the vote");
  const um = [
    ["FaUserCheck", HEX.muted, "Verbalised confidence", "U = 1 − c", "The AI's own probability that the test is correct. Baseline."],
    ["FaLayerGroup", HEX.blue, "Structural entropy", "U = −log p(cluster) / log N", "Cluster samples by oracle templates, event types and traffic layout — semantic entropy on the artefact " + cite("kuhn23") + "."],
    ["FaCar", HEX.teal, "Behavioural entropy = the vote", "U = −log(k/N) / log N  ⇔  1 − k/N", "Cluster samples by what they do: own verdict, collision, emergency braking, minimum-gap band. k = size of the sample's group."],
    ["FaCircleNodes", HEX.purple, "Oracle disagreement", "U = disagreeing peers / (N − 1)", "Apply every peer's oracle to this sample's trace; disagreement means an ambiguous requirement."],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, col, h, f, d] = um[i], x = MX + (i % 2) * 6.15, y = 1.45 + Math.floor(i / 2) * 2.3;
    card(s, x, y, 5.95, 2.1, i === 2 ? "E6F3F3" : HEX.mist);
    await badge(s, x + 0.3, y + 0.3, 0.7, ic, col);
    txt(s, h, { x: x + 1.2, y: y + 0.22, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: HEX.navy });
    txt(s, f, { x: x + 1.2, y: y + 0.68, w: 4.6, h: 0.4, fontSize: 13, fontFace: "Courier New", color: col === HEX.muted ? HEX.ink : col, bold: true });
    txt(s, d, { x: x + 1.2, y: y + 1.08, w: 4.55, h: 0.95, fontSize: 13 });
  }
  txt(s, "Both forms of the behavioural signal fall as the group size k grows, so they rank tests identically (same AUROC); the agreement form 1 − k/N is the one presented. Oracle disagreement was uninformative on the development round.", { x: MX, y: 6.12, w: W - 2 * MX, h: 0.5, fontSize: 12, color: HEX.muted, italic: true });
  footnotes(s, 6.65);
  s.addNotes("If asked for the formal version: the vote is behavioural entropy. The surprisal form and the agreement form are both monotone in the size of the test's behaviour group, so they rank tests identically. Structural entropy groups tests by what they look like; oracle disagreement cross-applies the tests' checks; verbalised confidence is the baseline.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "Backup" });
  addTitle(s, "Backup · AUROC of every research signal");
  const labs = ["Verbalised conf.", "Structural", "Oracle disagr.", "Combined (struct.+behav.)", "Agreement / behavioural"];
  const ks = ["U_verbal", "U_struct", "U_oracle", "U_comb", "U_disagree"];
  s.addChart(pres.charts.BAR, [
    { name: "All tests (pooled)", labels: labs, values: ks.map((k) => +AU[k].toFixed(2)) },
    { name: "Within one requirement × model", labels: labs, values: ks.map((k) => +AW[k].toFixed(2)) }],
  { x: MX, y: 1.45, w: W - 2 * MX, h: 4.9, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.teal, HEX.ice],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.ink, dataLabelFormatCode: "0.00",
    valAxisMinVal: 0, valAxisMaxVal: 1, valAxisMajorUnit: 0.25, valAxisLabelFormatCode: "0.00",
    valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    showTitle: true, title: "AUROC for detecting incorrect tests (0.5 = random)", titleFontSize: 14, titleColor: HEX.navy, titleFontFace: "+mn-lt" });
  source(s, `Within-batch AUROC averaged over the ${S.n_batches_with_errors} requirement × model batches that contain incorrect tests. Batch level: behavioural entropy vs. share of correct tests, Spearman ρ = ${f2(S.batch_spearman.behav_entropy)}.`, 6.5);
  s.addNotes(`All signals: agreement and behavioural entropy are identical at ${f2(AU.U_disagree)} pooled and ${f2(AW.U_disagree)} within a batch; the combined score adds structural entropy and does not help. Verbalised confidence ${f2(AU.U_verbal)}, oracle disagreement is near random.`);

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
