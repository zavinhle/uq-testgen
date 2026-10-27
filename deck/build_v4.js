// Builds the v4 interview deck: Demo 1 (QA-Flow) only -> the problem -> the PhD plan -> me
//   node build_v4.js  ->  PhD_Proposal_ZaVinhLe_v4.pptx
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
const QF = JSON.parse(fs.readFileSync(path.join(__dirname, "assets", "qaflow_run.json"), "utf8")); // Demo 1 summary + ground truth
const REFS = require("./refs.js");
const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "PhD_Proposal_ZaVinhLe_v4.pptx");

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
  s.addNotes("Good morning, and thank you for the invitation. My talk has three parts: first a real problem I hit with my own AI testing tool, shown in a demo; then how I want to solve it in the PhD, along the four objectives of the position; and finally what I bring.");

  // ======================= PART A: MOTIVATION 1 -> DEMO 1 =======================
  pres.addSection({ title: "A · The problem I hit (Demo 1)" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
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

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
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

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
  kicker(s, "PART A · DEMO 1");
  addTitle(s, "Demo 1: QA-Flow end to end on ShopLite");
  const vwA = 8.3, vhA = vwA * 900 / 1440;
  s.addMedia({ type: "video", path: A("qaflow_demo.mp4"), x: MX, y: 1.4, w: vwA, h: vhA, cover: "image/png;base64," + fs.readFileSync(A("qaflow_poster.png")).toString("base64") });
  card(s, MX + vwA + 0.3, 1.4, W - 2 * MX - vwA - 0.3, vhA, HEX.mist);
  txt(s, "What you see", { x: MX + vwA + 0.55, y: 1.6, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: HEX.navy });
  txt(s, bullets(["Six stages run live", "Analyst extracts requirements as verbatim quotes", "Designer writes web, API and security tests", "Code compiles them to Playwright", `Gate re-runs tests on ${QF.faults} broken builds`, "Report with drafted bugs"], 13), { x: MX + vwA + 0.55, y: 2.1, w: 2.75, h: vhA - 0.9 });
  txt(s, `${QF.videoLabel} · recorded from the running tool (long waits sped up)`, { x: MX, y: 1.4 + vhA + 0.08, w: vwA, h: 0.3, fontSize: 11, color: HEX.muted, italic: true });
  s.addNotes("[Play video.] The document goes in, the application is discovered and the documents are indexed. The analyst agent extracts atomic requirements as verbatim quotes, the designer writes tests per requirement using retrieved context, code compiles them into Playwright, they run, and the fault-injection gate re-runs passing tests on broken builds. The report shows traceability and drafted bug reports.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
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
  s.addNotes(`The result: QA-Flow extracted ${Q.requirements} requirements and ran ${Q.tests} tests, and it found ${nFound} of the 7 planted defects. But when I re-ran every failing test on a build with the planted defects fixed, ${nWrong} of the ${nFail} FAILs still failed: the bug was in the AI's test, not in the shop. Every one of them passed all four trust rules. So the rules cannot tell a wrong test from a real defect. Let me show you one up close.`);

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
  kicker(s, "PART A · A CLOSER LOOK");
  addTitle(s, "Inside one wrong test: the SAVE10 discount code");
  txt(s, [{ text: "Requirement:  ", options: { bold: true } }, { text: "“The discount code SAVE10 reduces the subtotal by 10 percent.”" },
    { text: "   ShopLite has a planted bug here: SAVE10 takes off €10 instead of 10%.", options: { color: HEX.muted } }],
  { x: MX, y: 1.3, w: W - 2 * MX, h: 0.45, fontSize: 15, valign: "middle" });
  const flows = [
    ["What a correct test would do", HEX.green, "E7F4EC", [["FaCartPlus", "Put shoes (€89.00) in the cart"], ["FaTag", "Enter the code SAVE10"], ["FaMagnifyingGlass", "Expect €80.10 → sees €79.00"]],
      "FAIL → the real discount bug is found"],
    ["What the AI's test did", HEX.red, "FDECEA", [["FaCartShopping", "Opened the cart directly — it was empty"], ["FaTag", "Never reached the discount"], ["FaMagnifyingGlass", "Saw €0.00 → stopped"]],
      "FAIL → filed as a bug, but the bug is in the test"],
  ];
  for (let c = 0; c < 2; c++) {
    const [head, col, bg, steps, out] = flows[c], x = MX + c * 6.17, w = 5.95;
    card(s, x, 1.95, w, 3.75, bg);
    txt(s, head, { x: x + 0.3, y: 2.07, w: w - 0.6, h: 0.4, fontSize: 18, bold: true, color: col });
    for (let i = 0; i < steps.length; i++) {
      const [ic, lab] = steps[i], y = 2.62 + i * 0.75;
      await badge(s, x + 0.3, y, 0.55, ic, col);
      txt(s, `${i + 1}. ${lab}`, { x: x + 1.05, y, w: w - 1.3, h: 0.55, fontSize: 16, valign: "middle" });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 4.95, w: w - 0.6, h: 0.55, rectRadius: 0.08, fill: { color: HEX.white }, line: { color: col, width: 1.5 } });
    txt(s, out, { x: x + 0.45, y: 4.95, w: w - 0.9, h: 0.55, fontSize: 15, bold: true, color: col, valign: "middle" });
  }
  txt(s, [{ text: "Same story for privacy: ", options: { bold: true } }, { text: "the tests for “a customer must not see other customers' orders” logged in the wrong way, got “not logged in”, and failed. The real data leak was never checked." }],
    { x: MX, y: 5.9, w: W - 2 * MX, h: 0.6, fontSize: 14, color: HEX.ink });
  s.addNotes("Let me show one of these wrong tests up close, because it explains the whole problem. The requirement says the code SAVE10 gives ten percent off. On the left is what a correct test would do: put a pair of shoes for 89 euros in the cart, enter SAVE10, and expect 80 euros 10. ShopLite has a planted bug and shows 79 euros, so the correct test fails and finds the bug. On the right is what the AI's test actually did: it opened the cart without adding anything, so the cart was empty, it never reached the discount, it saw zero euros and failed. The report files this as a bug, but the bug is in the test, and the real discount bug was never tested. The privacy tests made the same kind of mistake: they logged in the wrong way, got 'not logged in', and never checked whether one customer can see another customer's orders.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "A · The problem I hit (Demo 1)" });
  kicker(s, "PART A · WHY THE RULES MISS IT");
  addTitle(s, "The rules check the form of a test, not its meaning");
  const rh2 = ["Trust rule in QA-Flow", "What it checks", "The wrong SAVE10 test"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  const rules = [["Schema check", "The AI's answer is valid and uses only allowed steps", "✓ passed"],
    ["Grounding check", "Pages and fields exist; the quote is word-for-word from the requirement", "✓ passed"],
    ["NEEDS REVIEW flag", "Tests without a source in the document go to a human", "✓ not flagged"],
    ["Fault-injection gate", "A passing test must fail on a broken build", "– only checks passing tests"]];
  s.addTable([rh2, ...rules.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j !== 1, color: j === 2 ? (c.startsWith("✓") ? HEX.green : HEX.muted) : j === 0 ? HEX.navy : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: MX, y: 1.45, w: 7.4, colW: [2.1, 3.4, 1.9], rowH: 0.62, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.1, valign: "middle" });
  txt(s, "No rule asks: does this test really test the requirement?", { x: MX, y: 4.98, w: 7.4, h: 0.45, fontSize: 17, bold: true, color: HEX.red });
  txt(s, "A test can run cleanly and still be wrong in three ways:", { x: MX, y: 5.45, w: 7.4, h: 0.35, fontSize: 14 });
  const wrong = ["Wrong scenario", "Wrong check", "Catches no bug"];
  for (let i = 0; i < 3; i++) {
    const x = MX + i * 2.5;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 5.85, w: 2.35, h: 0.48, fill: { color: "FDECEA" }, line: { color: "FDECEA" } });
    txt(s, wrong[i], { x: x + 0.15, y: 5.85, w: 2.1, h: 0.48, fontSize: 15, bold: true, color: HEX.red, valign: "middle" });
  }
  card(s, 8.35, 1.45, 4.38, 4.88, HEX.navy);
  txt(s, "Why it matters more for autonomous systems", { x: 8.6, y: 1.65, w: 3.95, h: 0.75, fontSize: 16, bold: true, color: HEX.amber });
  txt(s, bullets([[{ text: "AI already helps engineer cyber-physical systems: ", bold: true }, { text: "checking driving scenarios, writing checks, building digital twins " + cite("wu24", "lu26") }],
    [{ text: "AI writes the tests themselves " + cite("wang24survey") }],
    [{ text: "In a self-driving car, a wrong test means ", bold: true }, { text: "false safety — or a team that stops trusting its test reports" }]], 14, HEX.ice), { x: 8.6, y: 2.45, w: 3.95, h: 3.65 });
  footnotes(s, 6.42);
  s.addNotes("Why did none of QA-Flow's rules catch it? Look at what each rule checks. The schema check: the answer is well-formed, passed. Grounding: the pages and fields exist and the quote is word-for-word, passed. The review flag: the test has a source, so it is not flagged. The fault-injection gate only looks at passing tests, and this one failed. So every rule checks the form of the test, and no rule asks whether the test really tests the requirement. A test can be wrong in three ways: the wrong scenario, the wrong check, or a test that catches no bug. In a web shop that wastes time. In an autonomous car, a wrong test means false safety, and AI is already being used to engineer exactly these systems.");

  pres.addSection({ title: "B · The PhD plan" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · THE PHD QUESTION", HEX.teal);
  addTitle(s, "When can we trust an AI-generated test?");
  card(s, MX, 1.4, W - 2 * MX, 1.05, "E6F3F3");
  txt(s, [{ text: "My idea  ", options: { bold: true, color: HEX.teal } },
    { text: "Ask the AI for the same test several times and compare what the tests do. If they agree, trust more; if they scatter, a human checks. This works for chatbot answers " + cite("kuhn23", "farquhar24", "manakul23") + " — but has not been studied for executable tests." }],
  { x: MX + 0.3, y: 1.45, w: W - 2 * MX - 0.6, h: 0.95, fontSize: 15, valign: "middle" });
  const cols = [
    ["Uncertainty in CPS testing", HEX.teal, "FaShieldHalved", ["U-Taxonomy, uncertainty-wise test modelling & generation " + cite("zhang16ecmfa", "zhang19sosym", "zhang19jss"),
      "DL prediction validator; uncertainty-wise oracles " + cite("catak22", "valle25"), "Uncertainty & quality of VLA robots " + cite("valleVLA")]],
    ["Driving-system testing", HEX.blue, "FaRoad", ["Avoidable-collision scenarios " + cite("calo20"), "Parameter coverage under uncertainty " + cite("laurent23"),
      "Robustness-guided falsification " + cite("zhang21cav", "zhang23falsifai")]],
    ["Oracles & industrial CPS", HEX.amber, "FaIndustry", ["ML-based oracles for elevators " + cite("arrieta21ast"), "Metamorphic relations for CPS " + cite("ayerdi21"),
      "Elevator robustness under uncertainty " + cite("han23")]],
    ["AI in testing", HEX.purple, "FaBrain", ["LLMs judging scenario realism " + cite("wu24"), "Foundation models for CPS " + cite("lu26", "ali25isola"),
      "LLM uncertainty in SE tasks " + cite("huang25")]],
  ];
  for (let i = 0; i < 4; i++) {
    const [h, col, ic, items] = cols[i], x = MX + i * 3.08;
    card(s, x, 2.6, 2.92, 2.55, HEX.mist);
    await badge(s, x + 0.2, 2.72, 0.5, ic, col);
    txt(s, h, { x: x + 0.8, y: 2.72, w: 2.0, h: 0.5, fontSize: 14, bold: true, color: HEX.navy, valign: "middle" });
    txt(s, bullets(items, 12), { x: x + 0.2, y: 3.32, w: 2.6, h: 1.78 });
  }
  card(s, MX, 5.25, W - 2 * MX, 0.48, "FFF4E0");
  txt(s, [{ text: "Gap  ", options: { bold: true, color: HEX.amber } }, { text: "Uncertainty is studied for CPS, neural networks, robots and AI text — not yet for AI that writes executable tests." }],
    { x: MX + 0.25, y: 5.25, w: W - 2 * MX - 0.5, h: 0.48, fontSize: 14, valign: "middle" });
  footnotes(s, 5.8, 7.5, "   ");
  s.addNotes("This leads to the question of my PhD: when can we trust an AI-generated test? My idea is simple: ask the AI for the same test several times and compare what the tests do. If they agree, we can trust them more; if they scatter, a human should look. This idea works for chatbot answers, in work on semantic entropy and self-consistency, but nobody has studied it for executable tests. It builds directly on your work: uncertainty-wise testing of cyber-physical systems, testing of driving systems and falsification, industrial oracles for elevators, and the use of foundation models in CPS engineering.");

  // ======================= PART C: GAPS -> NEXT (O1-O4) =======================
    s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · OBJECTIVES", HEX.amber);
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
  s.addNotes("This is how I read the four objectives of the position, phrased as questions. What Demo 1 cannot do yet maps onto them.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · GAPS → RESEARCH AGENDA", HEX.amber);
  addTitle(s, "What Demo 1 cannot do yet becomes the research agenda");
  const gaps = [
    ["Demo 1", `Rules cannot tell a wrong test from a real defect (${nWrong} of ${nFail} FAILs were wrong tests)`, "Uncertainty signals per generated test: agreement across repeated generations, behaviour of the test, white-box signals", "O1", HEX.teal],
    ["Demo 1", `Every FAIL looks the same, so a human must check all ${nFail}`, "Risk-based accept / review / reject, weighted by how severe a failure would be", "O2", HEX.blue],
    ["Demo 1", `Wrong tests hid ${G.defectsMissed.length} real bugs (discount, data leak); no measure of what a test is worth`, "Study how uncertainty relates to correctness, bug-finding power and cost", "O3", HEX.amber],
    ["Next", "One web shop and one AI model; safety matters most in autonomous systems", "Driving simulators and open-source ACPS, industrial cases, InnoGuard platform; newer and quantum ML models", "O4", HEX.purple],
  ];
  const gh = ["From", "Gap seen in the demo", "Next implementation", "Obj."].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([gh, ...gaps.map((g, i) => [
    { text: g[0], options: { fontSize: 13, bold: true, color: g[0] === "Demo 1" ? HEX.indigo : HEX.muted, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[1], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[2], options: { fontSize: 14, fill: { color: i % 2 ? HEX.white : HEX.mist } } },
    { text: g[3], options: { fontSize: 18, bold: true, color: g[4], align: "center", fill: { color: i % 2 ? HEX.white : HEX.mist } } }])],
  { x: MX, y: 1.45, w: W - 2 * MX, colW: [1.05, 5.0, 5.08, 1.0], rowH: 1.05, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.1, valign: "middle" });
  s.addNotes("Each thing Demo 1 cannot do becomes an objective. The rules cannot tell a wrong test from a real defect: I need uncertainty signals for every generated test, that is O1. Every FAIL looks the same, so a human checks all of them: I need risk-based decisions, O2. Wrong tests hid real bugs, and we have no measure of what a generated test is worth: a study of uncertainty against correctness, bug-finding power and cost, O3. And everything so far is one web shop and one model: the PhD moves to autonomous systems, industrial cases and newer models, O4.");

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · RESEARCH QUESTIONS", HEX.amber);
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

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · WORK PLAN", HEX.amber);
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

  s = newSlide({ masterName: "CONTENT", sectionTitle: "B · The PhD plan" });
  kicker(s, "PART B · CONTRIBUTIONS", HEX.amber);
  addTitle(s, "Expected contributions to InnoGuard");
  const contrib = [
    ["FaGaugeHigh", HEX.teal, "Metrics", "Uncertainty metrics for AI-generated tests, e.g. agreement between repeated generations, measured on what the tests do"],
    ["FaScaleBalanced", HEX.blue, "Methods", "Risk-based acceptance of AI-generated artefacts with calibrated, safety-aware thresholds"],
    ["FaChartLine", HEX.amber, "Evidence", "Empirical laws linking uncertainty to correctness, fault detection and cost across models"],
    ["FaRocket", HEX.purple, "Open science", "QA-Flow, an open ACPS benchmark and replication packages; integration with the InnoGuard platform"],
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
  pres.addSection({ title: "C · What I can contribute" });
  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · What I can contribute" });
  kicker(s, "PART C · MY BACKGROUND", HEX.green);
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

  s = newSlide({ masterName: "CONTENT", sectionTitle: "C · What I can contribute" });
  kicker(s, "PART C · WHAT I BRING", HEX.green);
  addTitle(s, "What I bring, and why this team");
  const fit = [
    ["FaCar", HEX.blue, "Testing cyber-physical systems", "Automotive ECU testing (CAN, CAPL, ETAS INCA); robotics (ROS/ROS2) and embedded domains"],
    ["FaWaveSquare", HEX.teal, "Verification research", "Falsification, function abstraction via active learning, property-based testing — close to robustness-guided falsification " + cite("zhang23falsifai")],
    ["FaBrain", HEX.amber, "LLMs for quality engineering", "Built QA-Flow end to end: RAG, LLM agents, MCP-based workflows, Playwright"],
    ["FaFlask", HEX.purple, "Python & empirical studies", "Controlled experiments on test generators (KCC'25, KCSE'26); ground-truth check of QA-Flow in this talk"],
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
  txt(s, bullets(["Uncertainty-wise CPS testing (Simula)", "ADS testing & falsification (NII)", "Industrial CPS oracles & elevators (Mondragon)", "My work sits exactly where the three meet"], 14, HEX.ice), { x: 9.1, y: 2.15, w: 3.45, h: 3.5 });
  footnotes(s, 6.3);
  s.addNotes("What I bring maps onto the preferred profile: CPS testing practice, verification research, hands-on LLM tooling — QA-Flow in this talk — and empirical studies. And why this team: uncertainty-wise CPS testing at Simula, ADS testing and falsification with Prof. Arcaini, industrial oracles at Mondragon. My work sits exactly where these meet.");

  s = newSlide({ masterName: "DARK", sectionTitle: "C · What I can contribute" });
  addTitle(s, "Thank you");
  txt(s, bullets(["QA-Flow: AI writes and runs tests end to end — and finds real bugs", "But its rules cannot spot the AI's own wrong tests", "PhD: measure when AI-generated tests can be trusted — O1 to O4"], 18, HEX.ice), { x: MX, y: 3.4, w: 9.5, h: 1.9 });
  txt(s, "zavinhle@gmail.com", { x: MX, y: 5.45, w: 6, h: 0.4, fontSize: 16, color: HEX.white, bold: true });
  s.addImage({ path: A("strip.png"), x: 0, y: H - 1.33, w: W, h: W * 160 / 1600 });
  s.addNotes("To summarise: QA-Flow lets AI write and run tests end to end, and it finds real bugs. But its rules cannot spot the AI's own wrong tests. My PhD is about measuring when AI-generated tests can be trusted, with metrics, risk methods, evidence and transfer to autonomous systems, along O1 to O4. Thank you — I look forward to your questions.");

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
  addTitle(s, "Backup · Threats to validity and mitigations");
  const thr = [["Construct", "Wrong test = still fails after the planted defects are fixed; planted defects approximate real faults", "Expert-validated oracles; real fault data from industrial partners"],
    ["Internal", "Ground truth and prompts designed by me", "Independent labelling; several prompt variants; pre-registered designs"],
    ["External", `One web shop, one AI model, ${Q.tests} tests in one run`, "Open-source ACPS, industrial cases, InnoGuard platform, open-weight models"],
    ["Conclusion", `${nWrong} wrong tests out of ${nFail} FAILs: small numbers`, "Repeated runs, larger benchmarks, effect sizes and confidence intervals"]];
  const th = ["Threat", "In the demos", "Mitigation in the PhD"].map((t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 14 } }));
  s.addTable([th, ...thr.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 0 ? HEX.navy : HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))],
    { x: MX, y: 1.5, w: W - 2 * MX, colW: [1.8, 5.2, 5.13], rowH: 1.0, border: { type: "solid", color: HEX.line, pt: 0.75 }, margin: 0.12, valign: "middle" });
  s.addNotes("Main threats: planted defects and my own ground truth define what counts as wrong; one shop, one model and one run; and small numbers. Each has a mitigation in the plan.");

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
