// Records the end-to-end demo video of the dashboard (live LLM run).
// Usage: node record.mjs [rid] [model] [n]   -> writes video/*.webm and video/marks.json
import { chromium } from "playwright";
import fs from "fs";

const [rid = "R6", model = "gpt-5.6-luna", n = "8"] = process.argv.slice(2);
const W = 1440, H = 900;
const t0 = Date.now();
const marks = [];
const mark = (name) => { marks.push({ name, t: (Date.now() - t0) / 1000 }); console.log(name, marks.at(-1).t); };

const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL || "chrome" });
const ctx = await browser.newContext({ viewport: { width: W, height: H } });
const page = await ctx.newPage();
// Capture frames with the Chrome DevTools screencast (encoded later with ffmpeg).
fs.mkdirSync("video/frames", { recursive: true });
const cdp = await ctx.newCDPSession(page);
const frames = [];
cdp.on("Page.screencastFrame", async (f) => {
  const file = `video/frames/f${String(frames.length).padStart(6, "0")}.jpg`;
  fs.writeFileSync(file, Buffer.from(f.data, "base64"));
  frames.push({ file, t: f.metadata.timestamp });
  await cdp.send("Page.screencastFrameAck", { sessionId: f.sessionId }).catch(() => {});
});
await cdp.send("Page.startScreencast", { format: "jpeg", quality: 85, maxWidth: W, maxHeight: H, everyNthFrame: 1 });
const sleep = (ms) => page.waitForTimeout(ms);

async function caption(text) {
  await page.evaluate((t) => {
    let el = document.getElementById("demo-cap");
    if (!el) {
      el = document.createElement("div");
      el.id = "demo-cap";
      el.style.cssText = "position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:9999;background:rgba(18,26,46,.92);color:#fff;padding:10px 18px;border-radius:10px;font:600 16px Inter,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.25);max-width:1100px;text-align:center;transition:opacity .3s";
      document.body.appendChild(el);
    }
    el.textContent = t; el.style.opacity = t ? 1 : 0;
  }, text);
}
async function smoothScroll(y, ms = 1200) {
  await page.evaluate(async ([y, ms]) => {
    const s = window.scrollY, d = y - s, st = performance.now();
    await new Promise(r => { const f = (n) => { const k = Math.min(1, (n - st) / ms); window.scrollTo(0, s + d * (0.5 - Math.cos(Math.PI * k) / 2)); k < 1 ? requestAnimationFrame(f) : r(); }; requestAnimationFrame(f); });
  }, [y, ms]);
}

await page.goto("http://127.0.0.1:8765/");
await sleep(1500);
mark("intro");
await caption("UQ-TestGen: can we trust tests that an AI writes for a self-driving feature?");
await sleep(3500);
await page.selectOption("#rid", rid);
await caption("1 · A requirement in plain English");
await sleep(3000);
await page.selectOption("#model", model);
await page.fill("#n", n); await page.dispatchEvent("#n", "input");
await caption(`2 · Ask the AI for the same test ${n} times (${model})`);
await sleep(2500);
await page.click("#run");
mark("gen_start");
await caption("Each answer = a driving scenario + a pass/fail check + how sure the AI says it is");
await page.waitForFunction(() => document.querySelectorAll("#samples .sample").length >= 1, null, { timeout: 240000 });
await page.waitForFunction(() => typeof state !== "undefined" && state.batch && state.batch.groups, null, { timeout: 300000 });
mark("gen_end");
await caption("3 · Run every test in the simulator and group the tests by what happened");
await sleep(3500);
await page.waitForFunction(() => !document.querySelector("#run").disabled, null, { timeout: 300000 });
mark("pipeline_done");
await caption("4 · The vote: big group → accept · small group → human review · alone → reject");
await smoothScroll(470);
await sleep(4500);

// open the test that agrees least with the others, then the one that agrees most; captions state what really happened
const info = await page.evaluate(() => {
  const ss = Object.values(state.samples).filter(s => s.uq && s.uq.group);
  const n = ss.length, tri = { accept: "accepted", review: "human review", reject: "rejected" };
  return ss.map(s => ({ idx: s.idx, k: Math.round(s.uq.agreement * n), n, dec: tri[s.triage], truth: s.gt?.label || "", killed: s.gt?.killed_gen || [] }))
    .sort((a, b) => a.k - b.k || a.idx - b.idx);
});
if (info.length) {
  const lo = info[0], hi = info[info.length - 1];
  await page.evaluate((i) => { selIdx = i; render(); }, lo.idx);
  await caption(`Test #${lo.idx}: only ${lo.k} of ${lo.n} tests behaved like it → ${lo.dec} · truth checked later: ${lo.truth}`);
  await sleep(5000);
  await smoothScroll(470);
  await sleep(3500);
  await page.evaluate((i) => { selIdx = i; render(); }, hi.idx);
  await caption(`Test #${hi.idx}: ${hi.k} of ${hi.n} tests behaved like it → ${hi.dec} · truth checked later: ${hi.truth}`);
  await sleep(6000);
  const rp = hi.killed.length ? hi : info.slice().reverse().find(x => x.killed.length);
  if (rp) {
    if (rp.idx !== hi.idx) { await page.evaluate((i) => { selIdx = i; render(); }, rp.idx); await sleep(800); }
    await page.selectOption("#gifv", rp.killed[0]);
    await page.dispatchEvent("#gifv", "change");
    await caption(`Test #${rp.idx} on a deliberately broken controller (${rp.killed[0]}) → its check catches the bug`);
    await sleep(7000);
  }
}
mark("experiment");
await page.click("#tabs button[data-tab=experiment]");
await smoothScroll(0, 400);
await caption("Experiment: 210 AI-written tests (7 requirements × 3 AIs × 10 tries)");
await sleep(5500);
await caption("Agreement spots wrong tests better than the AI's own confidence (0.90 vs 0.74) · 173 accepted, 171 correct");
await sleep(6000);
await smoothScroll(900);
await caption("Gap: the vote is still coarse inside one requirement, thresholds are hand-picked → PhD research");
await sleep(6500);
await caption("");
await sleep(800);
mark("end");
await cdp.send("Page.stopScreencast");
await ctx.close();
await browser.close();
// ffmpeg concat list with real frame durations
const lines = [];
for (let i = 0; i < frames.length; i++) {
  const d = i + 1 < frames.length ? Math.max(0.001, frames[i + 1].t - frames[i].t) : 1;
  lines.push(`file '${frames[i].file.replace("video/", "")}'`, `duration ${d.toFixed(3)}`);
}
lines.push(`file '${frames.at(-1).file.replace("video/", "")}'`);
fs.writeFileSync("video/frames.txt", lines.join(String.fromCharCode(10)));
fs.writeFileSync("video/marks.json", JSON.stringify({ marks, frames: frames.length, t0: frames[0]?.t }, null, 1));
console.log("frames", frames.length);
