"""Aggregates experiment runs: does uncertainty predict test correctness / effectiveness?
Writes results/summary.json and slide-ready figures in results/figures/."""
from __future__ import annotations

import glob
import json
import os
import sys

import numpy as np
from scipy.stats import spearmanr
from sklearn.metrics import roc_auc_score

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from core.uq import ACCEPT_T, REJECT_T, triage  # noqa: E402

METRICS = ["U_verbal", "U_struct", "U_behav", "U_oracle", "U_comb"]
NAMES = {"U_verbal": "Verbalised conf.", "U_struct": "Structural entropy", "U_behav": "Behavioural entropy",
         "U_oracle": "Oracle disagreement", "U_comb": "Combined (ours)"}
LABELS = ["correct & effective", "correct but weak", "wrong oracle", "false alarm oracle", "irrelevant scenario",
          "outside ODD", "invalid (rule gate)"]


def load_rows(pattern="exp_*.json"):
    rows, batches = [], []
    for f in sorted(glob.glob(os.path.join(ROOT, "results", "runs", pattern))):
        r = json.load(open(f, encoding="utf-8"))
        corr = []
        for s in r["samples"]:
            row = {"rid": r["rid"], "model": r["model"], "idx": s["idx"], "latency_s": s["latency_s"],
                   "tokens": s.get("tokens"), "valid": not s["validation_errors"], "triage": s["triage"]}
            if s.get("gt"):
                g = s["gt"]
                row.update(correct=g["correct"], useful=g["correct"] and g["ms_gen"] > 0, label=g["label"],
                           ms_gen=g["ms_gen"], ms_ref=g["ms_ref"], sim_s=g["sim_seconds_all"])
                corr.append(g["correct"])
            else:
                row.update(correct=False, useful=False, label="invalid (rule gate)", ms_gen=0.0, ms_ref=0.0, sim_s=0.0)
            if s.get("uq"):
                row.update({m: s["uq"][m] for m in METRICS})
            rows.append(row)
        batches.append({"rid": r["rid"], "model": r["model"], **(r.get("batch_uq") or {}),
                        "correct_rate": float(np.mean(corr)) if corr else 0.0, "n": r["n"],
                        "valid_rate": float(np.mean([not s["validation_errors"] for s in r["samples"]]))})
    return rows, batches


def auroc(y, s):
    y, s = np.asarray(y), np.asarray(s, dtype=float)
    if len(set(y.tolist())) < 2:
        return None
    return float(roc_auc_score(y, s))


def risk_coverage(rows, metric):
    v = sorted([r for r in rows if r["valid"] and metric in r], key=lambda r: r[metric])
    cov, prec = [], []
    for k in range(1, len(v) + 1):
        cov.append(k / len(v))
        prec.append(float(np.mean([r["correct"] for r in v[:k]])))
    return cov, prec


def summarize(rows, batches):
    valid = [r for r in rows if r["valid"] and "U_comb" in r]
    models = sorted({r["model"] for r in rows})
    out = {"n_samples": len(rows), "n_valid": len(valid), "models": models,
           "thresholds": {"accept": ACCEPT_T, "reject": REJECT_T}}
    out["overall"] = {
        "valid_rate": float(np.mean([r["valid"] for r in rows])) if rows else 0,
        "correct_rate_all": float(np.mean([r["correct"] for r in rows])) if rows else 0,
        "correct_rate_valid": float(np.mean([r["correct"] for r in valid])) if valid else 0,
        "useful_rate_valid": float(np.mean([r["useful"] for r in valid])) if valid else 0,
        "mean_ms_gen_valid": float(np.mean([r["ms_gen"] for r in valid])) if valid else 0,
        "mean_latency_s": float(np.mean([r["latency_s"] for r in rows])) if rows else 0,
        "mean_tokens": float(np.nanmean([r["tokens"] or np.nan for r in rows])) if rows else 0,
    }
    out["labels"] = {lab: sum(r["label"] == lab for r in rows) for lab in LABELS}
    out["labels_by_model"] = {m: {lab: sum(r["label"] == lab for r in rows if r["model"] == m) for lab in LABELS} for m in models}
    out["auroc_incorrect"] = {m: auroc([not r["correct"] for r in valid], [r[m] for r in valid]) for m in METRICS}
    out["auroc_not_useful"] = {m: auroc([not r["useful"] for r in valid], [r[m] for r in valid]) for m in METRICS}
    out["auroc_by_model"] = {mo: {m: auroc([not r["correct"] for r in valid if r["model"] == mo],
                                           [r[m] for r in valid if r["model"] == mo]) for m in METRICS} for mo in models}
    # within-batch AUROC: does uncertainty rank tests WITHIN one requirement x model batch?
    out["auroc_within_batch"] = {}
    keys = sorted({(r["rid"], r["model"]) for r in valid})
    for m in METRICS:
        vals = [auroc([not r["correct"] for r in valid if (r["rid"], r["model"]) == k],
                      [r[m] for r in valid if (r["rid"], r["model"]) == k]) for k in keys]
        vals = [v for v in vals if v is not None]
        out["auroc_within_batch"][m] = float(np.mean(vals)) if vals else None
    out["n_batches_with_errors"] = sum(1 for k in keys if 0 < sum(not r["correct"] for r in valid if (r["rid"], r["model"]) == k)
                                       < sum(1 for r in valid if (r["rid"], r["model"]) == k))
    out["n_incorrect"] = sum(not r["correct"] for r in rows)
    out["spearman_ms"] = {}
    for m in METRICS:
        x, y = [r[m] for r in valid], [r["ms_gen"] for r in valid]
        rho = spearmanr(x, y).statistic if len(set(x)) > 1 and len(set(y)) > 1 else None
        out["spearman_ms"][m] = None if rho is None or np.isnan(rho) else float(rho)
    out["risk_coverage"] = {m: dict(zip(("coverage", "precision"), risk_coverage(rows, m))) for m in ("U_verbal", "U_comb")}
    tri = {}
    for t in ("accept", "review", "reject"):
        sel = [r for r in rows if r["triage"] == t]
        tri[t] = {"n": len(sel), "correct": sum(r["correct"] for r in sel),
                  "precision": float(np.mean([r["correct"] for r in sel])) if sel else None}
    out["triage"] = tri
    out["by_model"] = {}
    for mo in models:
        rr = [r for r in rows if r["model"] == mo]
        vv = [r for r in rr if r["valid"]]
        out["by_model"][mo] = {"n": len(rr), "valid_rate": float(np.mean([r["valid"] for r in rr])),
                               "correct_rate": float(np.mean([r["correct"] for r in rr])),
                               "mean_ms_gen": float(np.mean([r["ms_gen"] for r in vv])) if vv else 0,
                               "mean_latency_s": float(np.mean([r["latency_s"] for r in rr])),
                               "mean_tokens": float(np.nanmean([r["tokens"] or np.nan for r in rr])),
                               "mean_verbal_conf": float(np.mean([1 - r["U_verbal"] for r in vv if "U_verbal" in r])) if vv else None}
    out["by_req"] = {}
    for rid in sorted({r["rid"] for r in rows}):
        rr = [r for r in rows if r["rid"] == rid]
        out["by_req"][rid] = {"correct_rate": float(np.mean([r["correct"] for r in rr])),
                              "mean_U_comb": float(np.mean([r["U_comb"] for r in rr if "U_comb" in r] or [np.nan]))}
    out["batches"] = batches
    bx = [b for b in batches if "struct_entropy" in b]
    if len(bx) > 2:
        out["batch_spearman"] = {
            k: (lambda rho: None if np.isnan(rho) else float(rho))(spearmanr([b[k] for b in bx], [b["correct_rate"] for b in bx]).statistic)
            for k in ("struct_entropy", "behav_entropy", "mean_verbal_conf")}
    # non-functional: latency of correct vs incorrect, uncertainty vs latency
    out["nonfunctional"] = {
        "latency_correct": float(np.mean([r["latency_s"] for r in valid if r["correct"]] or [np.nan])),
        "latency_incorrect": float(np.mean([r["latency_s"] for r in valid if not r["correct"]] or [np.nan])),
        "spearman_Ucomb_latency": (lambda v: None if np.isnan(v) else float(v))(spearmanr([r["U_comb"] for r in valid], [r["latency_s"] for r in valid]).statistic) if len(valid) > 2 else None,
        "spearman_Ucomb_tokens": (lambda v: None if np.isnan(v) else float(v))(spearmanr([r["U_comb"] for r in valid if r["tokens"]], [r["tokens"] for r in valid if r["tokens"]]).statistic) if len(valid) > 2 else None,
        "sim_seconds_per_test_9_versions": float(np.mean([r["sim_s"] for r in valid] or [np.nan])),
    }
    out["points"] = [{k: r.get(k) for k in ("rid", "model", "idx", "U_comb", "U_verbal", "ms_gen", "correct", "label", "latency_s")} for r in valid]
    return out


# ---------- figures ----------
C = {"ink": "#1f2430", "muted": "#6b7280", "grid": "#e5e7eb", "ours": "#2563eb", "base": "#9ca3af",
     "good": "#16a34a", "bad": "#dc2626", "warn": "#d97706"}
SER = ["#2563eb", "#d97706", "#7c3aed", "#0891b2", "#db2777"]


def _style(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(C["grid"]); ax.spines["bottom"].set_color(C["grid"])
    ax.tick_params(colors=C["muted"], labelsize=10)
    ax.yaxis.grid(True, color=C["grid"], lw=0.8); ax.set_axisbelow(True)


def figures(S, rows, outdir):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11})
    os.makedirs(outdir, exist_ok=True)

    # 1. AUROC per metric
    fig, ax = plt.subplots(figsize=(7.2, 3.6), dpi=200)
    vals = [S["auroc_incorrect"][m] or 0 for m in METRICS]
    cols = [C["ours"] if m in ("U_comb", "U_behav", "U_oracle") else C["base"] for m in METRICS]
    b = ax.bar([NAMES[m] for m in METRICS], vals, color=cols, width=0.62)
    ax.axhline(0.5, color=C["muted"], ls="--", lw=1); ax.text(4.45, 0.51, "random", color=C["muted"], fontsize=9, ha="right")
    for bi, v in zip(b, vals):
        ax.text(bi.get_x() + bi.get_width() / 2, v + 0.015, f"{v:.2f}", ha="center", fontsize=10, color=C["ink"])
    ax.set_ylim(0, 1.05); ax.set_ylabel("AUROC (detect incorrect tests)", color=C["ink"])
    ax.tick_params(axis="x", labelsize=9); _style(ax)
    fig.tight_layout(); fig.savefig(os.path.join(outdir, "auroc.png")); plt.close(fig)

    # 2. risk-coverage
    fig, ax = plt.subplots(figsize=(6, 3.8), dpi=200)
    for m, col, lab in (("U_comb", C["ours"], "Combined uncertainty (ours)"), ("U_verbal", C["base"], "Verbalised confidence")):
        rc = S["risk_coverage"][m]
        ax.plot(rc["coverage"], rc["precision"], color=col, lw=2.2, label=lab)
    base = S["overall"]["correct_rate_valid"]
    ax.axhline(base, color=C["muted"], ls=":", lw=1.2); ax.text(0.02, base - 0.06, f"accept all: {base:.0%}", color=C["muted"], fontsize=9)
    ax.set_xlabel("Coverage (fraction of tests auto-accepted)"); ax.set_ylabel("Precision (accepted tests correct)")
    ax.set_ylim(min(0.8, base - 0.08), 1.01); ax.legend(frameon=False, fontsize=9, loc="lower left"); _style(ax)
    fig.tight_layout(); fig.savefig(os.path.join(outdir, "risk_coverage.png")); plt.close(fig)

    # 3. outcome breakdown by model (stacked)
    fig, ax = plt.subplots(figsize=(7.2, 3.4), dpi=200)
    pal = {"correct & effective": C["good"], "correct but weak": "#86efac", "wrong oracle": C["warn"],
           "false alarm oracle": "#fbbf24", "irrelevant scenario": C["bad"], "outside ODD": "#f87171", "invalid (rule gate)": "#6b7280"}
    models = S["models"]
    left = np.zeros(len(models))
    for lab in LABELS:
        v = np.array([S["labels_by_model"][m][lab] for m in models], dtype=float)
        if v.sum() == 0:
            continue
        ax.barh(models, v, left=left, color=pal[lab], label=lab, height=0.55)
        left += v
    ax.set_xlabel("Generated tests"); ax.legend(frameon=False, fontsize=8, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.25))
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(os.path.join(outdir, "outcomes.png")); plt.close(fig)

    # 4. scatter U_comb vs mutation score
    fig, ax = plt.subplots(figsize=(6, 3.8), dpi=200)
    rng = np.random.default_rng(0)
    for ok, col, lab in ((True, C["good"], "correct"), (False, C["bad"], "incorrect")):
        pts = [p for p in S["points"] if p["correct"] == ok]
        ax.scatter([p["U_comb"] + rng.normal(0, 0.01) for p in pts], [p["ms_gen"] + rng.normal(0, 0.008) for p in pts],
                   s=22, alpha=0.75, color=col, label=lab, edgecolors="none")
    ax.axvline(ACCEPT_T, color=C["muted"], ls="--", lw=1); ax.axvline(REJECT_T, color=C["muted"], ls="--", lw=1)
    ax.text(ACCEPT_T / 2, 1.02, "accept", ha="center", fontsize=9, color=C["muted"])
    ax.text((ACCEPT_T + REJECT_T) / 2, 1.02, "review", ha="center", fontsize=9, color=C["muted"])
    ax.text((REJECT_T + 1) / 2, 1.02, "reject", ha="center", fontsize=9, color=C["muted"])
    ax.set_xlabel("Combined uncertainty U_comb"); ax.set_ylabel("Mutation score (own oracle)")
    ax.set_xlim(-0.03, 1.0); ax.set_ylim(-0.05, 1.08); ax.legend(frameon=False, fontsize=9, loc="center right"); _style(ax)
    fig.tight_layout(); fig.savefig(os.path.join(outdir, "scatter.png")); plt.close(fig)

    # 5. batch entropy vs correctness
    bx = [b for b in S["batches"] if "struct_entropy" in b]
    if bx:
        fig, ax = plt.subplots(figsize=(6, 3.6), dpi=200)
        for i, mo in enumerate(models):
            pts = [b for b in bx if b["model"] == mo]
            ax.scatter([b["behav_entropy"] for b in pts], [b["correct_rate"] for b in pts], s=46, color=SER[i % 5], label=mo)
            for b in [b for b in pts if b["behav_entropy"] > 0.01 or b["correct_rate"] < 1]:
                ax.annotate(b["rid"], (b["behav_entropy"], b["correct_rate"]), fontsize=7, color=C["muted"], xytext=(3, 3), textcoords="offset points")
        ax.set_xlabel("Behavioural entropy of the batch (N samples)"); ax.set_ylabel("Fraction of correct tests")
        ax.set_ylim(-0.05, 1.05); ax.legend(frameon=False, fontsize=8); _style(ax)
        fig.tight_layout(); fig.savefig(os.path.join(outdir, "batch_entropy.png")); plt.close(fig)


if __name__ == "__main__":
    rows, batches = load_rows()
    S = summarize(rows, batches)
    with open(os.path.join(ROOT, "results", "summary.json"), "w", encoding="utf-8") as fh:
        json.dump(S, fh, indent=1)
    figures(S, rows, os.path.join(ROOT, "results", "figures"))
    print(json.dumps({k: S[k] for k in ("n_samples", "n_valid", "overall", "labels", "auroc_incorrect", "auroc_not_useful",
                                        "spearman_ms", "triage", "by_model", "by_req", "nonfunctional")} | {"batch_spearman": S.get("batch_spearman")}, indent=1))
