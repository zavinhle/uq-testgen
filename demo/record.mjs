// Records the end-to-end demo video of the dashboard (live LLM run).
// Usage: node record.mjs [rid] [model] [n]   -> writes video/*.webm and video/marks.json
import { chromium } from "playwright";
import fs from "fs";

const [rid = "R6", model = "gpt-5.6-luna", n = "8"] = process.argv.slice(2);
const W = 1440, H = 900;
const t0 = Date.now();
const marks = [];
const mark = (name) => { marks.push({ name, t: (Date.now() - t0) / 1000 }); console.log(name, marks.at(-1).t); };

const browser = await chromium.launch({ channel: "chrome" });
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
await caption("UQ-TestGen: LLM-generated tests for an autonomous-driving ACC, with quantified uncertainty");
await sleep(3500);
await page.selectOption("#rid", rid);
await caption("1 · Pick a natural-language requirement");
await sleep(3000);
await page.selectOption("#model", model);
await page.fill("#n", n); await page.dispatchEvent("#n", "input");
await caption(`2 · Sample N=${n} independent tests from the LLM (${model}, via Codex CLI)`);
await sleep(2500);
await page.click("#run");
mark("gen_start");
await caption("Generating… each sample = scenario + STL-style oracle + self-reported confidence");
await page.waitForFunction(() => document.querySelectorAll("#samples .sample").length >= 1, null, { timeout: 240000 });
await page.waitForFunction(() => document.querySelector("#batch")?.innerText.includes("clusters") && !document.querySelector("#batch").innerText.includes("– structural"), null, { timeout: 300000 });
mark("gen_end");
await caption("3 · Rule gate → simulation in highway-env → uncertainty from sample disagreement (no ground truth)");
await sleep(3500);
await page.waitForFunction(() => !document.querySelector("#run").disabled, null, { timeout: 300000 });
mark("pipeline_done");
await caption("4 · Risk-based triage: accept / human review / reject — then ground truth via 8 seeded faults");
await smoothScroll(470);
await sleep(4500);

// open the most uncertain test, then the least uncertain one
const order = await page.evaluate(() => Object.values(state.samples).filter(s => s.uq).sort((a, b) => b.uq.U_comb - a.uq.U_comb).map(s => s.idx));
if (order.length) {
  await page.evaluate((i) => { selIdx = i; render(); }, order[0]);
  await caption("Most uncertain test: inspect its uncertainty signals and its ground-truth label");
  await sleep(4500);
  await smoothScroll(470);
  await sleep(4000);
  await page.evaluate((i) => { selIdx = i; render(); }, order[order.length - 1]);
  await caption("Least uncertain test: auto-accept candidate — replay of the scenario on the SUT");
  await sleep(6500);
  const killed = await page.evaluate((i) => (state.samples[i].gt?.killed_gen || [])[0], order[order.length - 1]);
  if (killed) {
    await page.selectOption("#gifv", killed);
    await page.dispatchEvent("#gifv", "change");
    await caption(`Same test on a faulty variant (${killed}) — its oracle detects the fault`);
    await sleep(7000);
  }
}
mark("experiment");
await page.click("#tabs button[data-tab=experiment]");
await smoothScroll(0, 400);
await caption("Experiment: 7 requirements × 3 LLMs × 10 samples — does uncertainty predict incorrect tests?");
await sleep(5500);
await smoothScroll(380);
await caption("Behavioural uncertainty separates correct from incorrect tests better than the LLM's own confidence");
await sleep(6000);
await smoothScroll(900);
await caption("Gap: per-test signals are coarse, ground truth is costly, LLM self-confidence is overconfident → PhD research");
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
