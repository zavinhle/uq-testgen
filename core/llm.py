"""LLM adapters. Default: OpenAI Codex CLI signed in with a ChatGPT account (black-box access:
no token log-probabilities, so uncertainty must be estimated from sampled outputs)."""
from __future__ import annotations

import json
import os
import re
import subprocess
import tempfile
import time
from dataclasses import dataclass

MODELS = ["gpt-5.6-luna", "gpt-5.6-terra", "gpt-6-astra"]


@dataclass
class LlmResult:
    text: str
    latency_s: float
    tokens: int | None
    model: str
    error: str | None = None


def codex_generate(prompt: str, model: str = "gpt-5.6-luna", effort: str = "low", timeout: int = 240) -> LlmResult:
    with tempfile.TemporaryDirectory(prefix="uqtg_") as wd:
        out = os.path.join(wd, "last.txt")
        cmd = ["codex", "exec", "-m", model, "-c", f'model_reasoning_effort="{effort}"', "-s", "read-only",
               "--ephemeral", "--skip-git-repo-check", "--ignore-user-config", "--color", "never", "-o", out, "-"]
        t0 = time.perf_counter()
        try:
            p = subprocess.run(cmd, input=prompt, capture_output=True, text=True, encoding="utf-8",
                               errors="replace", cwd=wd, timeout=timeout, shell=os.name == "nt")
        except subprocess.TimeoutExpired:
            return LlmResult("", time.perf_counter() - t0, None, model, "timeout")
        dt = time.perf_counter() - t0
        log = (p.stdout or "") + (p.stderr or "")
        m = re.search(r"tokens used\s*\n\s*([\d,]+)", log)
        tokens = int(m.group(1).replace(",", "")) if m else None
        text = open(out, encoding="utf-8").read() if os.path.exists(out) else ""
        err = None if text else (re.findall(r"ERROR: (.*)", log)[-1:] or [f"exit {p.returncode}"])[0][:300]
        return LlmResult(text, dt, tokens, model, err)


def extract_json(text: str):
    """Parses the first JSON object in a model reply (tolerates code fences / prose)."""
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t)
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        s, e = t.find("{"), t.rfind("}")
        if s >= 0 and e > s:
            return json.loads(t[s:e + 1])
        raise
