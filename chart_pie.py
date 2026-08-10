#!/usr/bin/env python3
"""Render monthly category pie chart as PNG."""

from __future__ import annotations

from pathlib import Path


def fmt_vnd(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def render_category_pie(
    by_category: dict[str, int],
    month: str,
    out_path: Path,
    total: int | None = None,
) -> Path:
    """Save a pie chart image. by_category: category -> amount (VND)."""
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError as e:
        raise SystemExit(
            "Missing matplotlib. Run: pip3 install --user matplotlib"
        ) from e

    # Prefer Vietnamese-capable fonts commonly on macOS
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [
        "Arial Unicode MS",
        "Hiragino Sans",
        "PingFang SC",
        "DejaVu Sans",
        "sans-serif",
    ]
    plt.rcParams["axes.unicode_minus"] = False

    items = sorted(
        ((k, v) for k, v in by_category.items() if v > 0),
        key=lambda x: -x[1],
    )
    if not items:
        raise ValueError("No category amounts to chart")

    labels = [k for k, _ in items]
    sizes = [v for _, v in items]
    total = total if total is not None else sum(sizes)

    colors = [
        "#4E79A7",
        "#F28E2B",
        "#E15759",
        "#76B7B2",
        "#59A14F",
        "#EDC948",
        "#B07AA1",
        "#FF9DA7",
        "#9C755F",
        "#BAB0AC",
    ]

    fig, ax = plt.subplots(figsize=(8, 6), dpi=150)
    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=None,
        autopct=lambda pct: f"{pct:.1f}%" if pct >= 3 else "",
        startangle=90,
        colors=colors[: len(sizes)],
        pctdistance=0.72,
        wedgeprops={"linewidth": 1, "edgecolor": "white"},
    )
    for t in autotexts:
        t.set_fontsize(9)
        t.set_color("white")
        t.set_fontweight("bold")

    legend_labels = [
        f"{lab} — {fmt_vnd(val)} ({val / total * 100:.1f}%)"
        for lab, val in items
    ]
    ax.legend(
        wedges,
        legend_labels,
        title="Danh mục",
        loc="center left",
        bbox_to_anchor=(1.0, 0.5),
        frameon=False,
        fontsize=9,
    )
    ax.set_title(
        f"Chi tiêu theo danh mục — {month}\nTổng: {fmt_vnd(total)} VND",
        fontsize=13,
        fontweight="bold",
        pad=16,
    )
    ax.axis("equal")

    out_path = Path(out_path).expanduser()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path
