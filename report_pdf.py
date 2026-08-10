#!/usr/bin/env python3
"""One-page monthly expense PDF report."""

from __future__ import annotations

from pathlib import Path


COLORS = [
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


def fmt_vnd(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _setup_fonts(plt) -> None:
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [
        "Arial Unicode MS",
        "Hiragino Sans",
        "PingFang SC",
        "DejaVu Sans",
        "sans-serif",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def render_month_pdf(
    *,
    month: str,
    total: int,
    by_category: dict[str, int],
    top_items: list[dict],
    vs_prev: dict | None,
    out_path: Path,
) -> Path:
    """
    Render 1-page A4 PDF:
    pie + category table + top expenses + vs previous month.
    """
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.gridspec import GridSpec
    except ImportError as e:
        raise SystemExit(
            "Missing matplotlib. Run: pip3 install --user matplotlib"
        ) from e

    _setup_fonts(plt)

    items = sorted(
        ((k, v) for k, v in by_category.items() if v > 0),
        key=lambda x: -x[1],
    )
    if not items and total == 0:
        raise ValueError("No data for PDF report")

    out_path = Path(out_path).expanduser()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    fig = plt.figure(figsize=(8.27, 11.69), dpi=150)  # A4
    fig.patch.set_facecolor("white")
    gs = GridSpec(
        4,
        2,
        figure=fig,
        height_ratios=[0.55, 2.2, 1.6, 1.8],
        hspace=0.45,
        wspace=0.25,
        left=0.08,
        right=0.92,
        top=0.94,
        bottom=0.05,
    )

    # --- Title ---
    ax_title = fig.add_subplot(gs[0, :])
    ax_title.axis("off")
    ax_title.text(
        0.5,
        0.75,
        f"Báo cáo chi tiêu — {month}",
        ha="center",
        va="center",
        fontsize=18,
        fontweight="bold",
        color="#1F4E79",
        transform=ax_title.transAxes,
    )
    ax_title.text(
        0.5,
        0.15,
        f"Tổng tháng: {fmt_vnd(total)} VND",
        ha="center",
        va="center",
        fontsize=13,
        transform=ax_title.transAxes,
    )

    # --- Pie ---
    ax_pie = fig.add_subplot(gs[1, 0])
    if items:
        sizes = [v for _, v in items]
        wedges, _, autotexts = ax_pie.pie(
            sizes,
            autopct=lambda pct: f"{pct:.0f}%" if pct >= 5 else "",
            startangle=90,
            colors=COLORS[: len(sizes)],
            pctdistance=0.7,
            wedgeprops={"linewidth": 1, "edgecolor": "white"},
        )
        for t in autotexts:
            t.set_fontsize(8)
            t.set_color("white")
            t.set_fontweight("bold")
        ax_pie.legend(
            wedges,
            [lab for lab, _ in items],
            loc="upper center",
            bbox_to_anchor=(0.5, -0.02),
            ncol=2,
            frameon=False,
            fontsize=8,
        )
    else:
        ax_pie.text(0.5, 0.5, "Không có dữ liệu", ha="center")
    ax_pie.set_title("Theo danh mục", fontsize=11, fontweight="bold", pad=8)
    ax_pie.axis("equal")

    # --- Category table ---
    ax_cat = fig.add_subplot(gs[1, 1])
    ax_cat.axis("off")
    ax_cat.set_title("Bảng danh mục", fontsize=11, fontweight="bold", pad=8)
    cat_rows = []
    for lab, val in items:
        pct = (val / total * 100) if total else 0
        cat_rows.append([lab, fmt_vnd(val), f"{pct:.1f}%"])
    if not cat_rows:
        cat_rows = [["—", "0", "0%"]]
    table = ax_cat.table(
        cellText=cat_rows,
        colLabels=["Danh mục", "Số tiền", "%"],
        loc="upper center",
        cellLoc="left",
        colLoc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.0, 1.45)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor("#DDDDDD")
        if r == 0:
            cell.set_facecolor("#1F4E79")
            cell.set_text_props(color="white", fontweight="bold")
            cell.set_height(0.08)
        elif r % 2 == 0:
            cell.set_facecolor("#F5F5F5")
        if c == 1:
            cell.set_text_props(ha="right")
        if c == 2:
            cell.set_text_props(ha="center")

    # --- Vs previous month ---
    ax_vs = fig.add_subplot(gs[2, :])
    ax_vs.axis("off")
    ax_vs.set_title("So với tháng trước", fontsize=11, fontweight="bold", pad=6)
    if not vs_prev or not vs_prev.get("prev_month"):
        ax_vs.text(
            0.5,
            0.5,
            "Chưa có dữ liệu tháng trước để so sánh.",
            ha="center",
            va="center",
            fontsize=10,
            color="#666666",
            transform=ax_vs.transAxes,
        )
    else:
        prev_total = int(vs_prev.get("prev_total") or 0)
        delta = int(vs_prev.get("delta") or (total - prev_total))
        pct = vs_prev.get("delta_pct")
        if pct is None and prev_total:
            pct = round(delta / prev_total * 100, 1)
        sign = "+" if delta > 0 else ""
        pct_txt = f" ({sign}{pct}%)" if pct is not None else ""
        tone = "#C0392B" if delta > 0 else "#27AE60" if delta < 0 else "#555555"
        lines = [
            f"Tháng trước ({vs_prev['prev_month']}): {fmt_vnd(prev_total)} VND",
            f"Chênh lệch: {sign}{fmt_vnd(delta)} VND{pct_txt}",
        ]
        movers = vs_prev.get("top_movers") or []
        if movers:
            lines.append("Danh mục biến động mạnh:")
            for m in movers[:3]:
                dlt = int(m["delta"])
                s = "+" if dlt > 0 else ""
                lines.append(
                    f"  • {m['category']}: {s}{fmt_vnd(dlt)} "
                    f"({fmt_vnd(m['prev'])} → {fmt_vnd(m['curr'])})"
                )
        ax_vs.text(
            0.05,
            0.85,
            "\n".join(lines),
            ha="left",
            va="top",
            fontsize=10,
            color=tone if len(lines) <= 2 else "#222222",
            family="sans-serif",
            transform=ax_vs.transAxes,
            linespacing=1.5,
        )
        # Color the delta line emphasis via a second pass if needed — keep simple

    # --- Top expenses ---
    ax_top = fig.add_subplot(gs[3, :])
    ax_top.axis("off")
    ax_top.set_title("Top khoản chi", fontsize=11, fontweight="bold", pad=6)
    top_rows = []
    for i, it in enumerate(top_items[:8], start=1):
        note = (it.get("note") or "").strip() or "—"
        if len(note) > 40:
            note = note[:37] + "…"
        top_rows.append(
            [
                str(i),
                it.get("date", ""),
                it.get("category", ""),
                note,
                fmt_vnd(int(it.get("amount") or 0)),
            ]
        )
    if not top_rows:
        top_rows = [["—", "—", "—", "—", "0"]]
    t2 = ax_top.table(
        cellText=top_rows,
        colLabels=["#", "Ngày", "Danh mục", "Ghi chú", "Số tiền"],
        loc="upper center",
        cellLoc="left",
    )
    t2.auto_set_font_size(False)
    t2.set_fontsize(8.5)
    t2.scale(1.0, 1.35)
    # Column widths roughly
    widths = [0.06, 0.16, 0.18, 0.40, 0.20]
    for (r, c), cell in t2.get_celld().items():
        cell.set_edgecolor("#DDDDDD")
        cell.set_width(widths[c])
        if r == 0:
            cell.set_facecolor("#1F4E79")
            cell.set_text_props(color="white", fontweight="bold", ha="center")
        elif r % 2 == 0:
            cell.set_facecolor("#F5F5F5")
        if c in (0, 1):
            cell.set_text_props(ha="center")
        if c == 4:
            cell.set_text_props(ha="right")

    fig.savefig(out_path, format="pdf", facecolor="white")
    plt.close(fig)
    return out_path
