// Takes a screenshot of a dashboard tab: node shot.mjs <tab> <out.png> [runId]
import { chromium } from "playwright";
const [tab = "pipeline", out = "shot.png", runId] = process.argv.slice(2);
const b = await chromium.launch({ channel: "chrome" });
const p = await b.newPage({ viewport: { width: 1500, height: 950 }, deviceScaleFactor: 1 });
await p.goto("http://127.0.0.1:8765/");
await p.waitForTimeout(1500);
if (runId) { await p.evaluate(id => loadRun(id), runId); await p.waitForTimeout(800); await p.click(".sample >> nth=0"); await p.waitForTimeout(2500); }
if (tab !== "pipeline") { await p.click(`#tabs button[data-tab=${tab}]`); await p.waitForTimeout(2000); }
await p.screenshot({ path: out, fullPage: true });
await b.close();
