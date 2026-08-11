#!/usr/bin/env python3
"""Summarize expenses for a day or a month from Excel."""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from datetime import date as date_cls
from datetime import datetime
from pathlib import Path

try:
    from workbook_io import DEFAULT_DIR, load_day_rows, month_path
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from workbook_io import DEFAULT_DIR, load_day_rows, month_path


def fmt_vnd(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def prev_month_key(year: int, month: int) -> tuple[int, int, str]:
    if month == 1:
        return year - 1, 12, f"{year - 1:04d}-12"
    return year, month - 1, f"{year:04d}-{month - 1:02d}"


def by_category(rows: list[dict]) -> dict[str, int]:
    out: dict[str, int] = defaultdict(int)
    for r in rows:
        out[r["category"]] += r["amount"]
    return dict(out)


def compare_months(
    curr_cat: dict[str, int],
    prev_cat: dict[str, int],
    curr_total: int,
    prev_total: int,
    prev_key: str,
) -> dict:
    cats = set(curr_cat) | set(prev_cat)
    movers = []
    for c in cats:
        a, b = curr_cat.get(c, 0), prev_cat.get(c, 0)
        movers.append(
            {"category": c, "curr": a, "prev": b, "delta": a - b}
        )
    movers.sort(key=lambda x: abs(x["delta"]), reverse=True)
    delta = curr_total - prev_total
    delta_pct = round(delta / prev_total * 100, 1) if prev_total else None
    return {
        "prev_month": prev_key,
        "prev_total": prev_total,
        "prev_total_fmt": fmt_vnd(prev_total),
        "delta": delta,
        "delta_fmt": fmt_vnd(delta),
        "delta_pct": delta_pct,
        "top_movers": movers[:5],
    }


def main() -> int:
    p = argparse.ArgumentParser(description="Summarize expenses")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--date", help="YYYY-MM-DD (day summary)")
    g.add_argument("--month", help="YYYY-MM (month summary)")
    p.add_argument("--data-dir", default=str(DEFAULT_DIR))
    p.add_argument(
        "--chart",
        action="store_true",
        help="Also write pie chart PNG (off by default; PDF is enough)",
    )
    p.add_argument(
        "--no-pdf",
        action="store_true",
        help="Skip 1-page PDF report for --month",
    )
    p.add_argument(
        "--chart-out",
        default="",
        help="Optional PNG path (only with --chart)",
    )
    p.add_argument(
        "--pdf-out",
        default="",
        help="Optional PDF path (default: ~/Documents/ChiTieu/reports/bao-cao-YYYY-MM.pdf)",
    )
    p.add_argument(
        "--top",
        type=int,
        default=8,
        help="How many top expenses to include (default 8)",
    )
    args = p.parse_args()

    data_dir = Path(args.data_dir).expanduser()
    vs_prev = None
    year = month = None

    if args.date:
        try:
            d = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"Invalid date: {args.date!r}", file=sys.stderr)
            return 1
        path = month_path(data_dir, d)
        rows = load_day_rows(path, d)
        scope = {"type": "day", "date": d.isoformat(), "sheet": f"{d.day:02d}"}
        want_chart = False
        want_pdf = False
        month_key = f"{d.year:04d}-{d.month:02d}"
    else:
        try:
            y, m = args.month.split("-")
            year, month = int(y), int(m)
            if not (1 <= month <= 12):
                raise ValueError
        except ValueError:
            print(f"Invalid month: {args.month!r} (use YYYY-MM)", file=sys.stderr)
            return 1
        path = month_path(data_dir, date_cls(year, month, 1))
        rows = load_day_rows(path, None)
        scope = {"type": "month", "month": f"{year:04d}-{month:02d}"}
        month_key = f"{year:04d}-{month:02d}"
        want_chart = bool(args.chart)
        want_pdf = not args.no_pdf

        py, pm, prev_key = prev_month_key(year, month)
        prev_path = month_path(data_dir, date_cls(py, pm, 1))
        prev_rows = load_day_rows(prev_path, None) if prev_path.exists() else []
        prev_total = sum(r["amount"] for r in prev_rows)
        prev_cat = by_category(prev_rows)
        # Always attach vs_prev when month mode (even if prev empty)
        curr_cat_tmp = by_category(rows)
        curr_total_tmp = sum(r["amount"] for r in rows)
        if prev_path.exists():
            vs_prev = compare_months(
                curr_cat_tmp, prev_cat, curr_total_tmp, prev_total, prev_key
            )
        else:
            vs_prev = {
                "prev_month": None,
                "prev_total": 0,
                "message": f"No file for {prev_key}",
            }

    cat = by_category(rows)
    total = sum(r["amount"] for r in rows)
    top_items = sorted(rows, key=lambda r: r["amount"], reverse=True)[: max(1, args.top)]

    chart_path = None
    if want_chart and cat:
        from chart_pie import render_category_pie

        out = (
            Path(args.chart_out).expanduser()
            if args.chart_out
            else data_dir / "charts" / f"bieu-do-{month_key}.png"
        )
        chart_path = str(render_category_pie(cat, month_key, out, total=total))

    pdf_path = None
    if want_pdf and (cat or total):
        from report_pdf import render_month_pdf

        out_pdf = (
            Path(args.pdf_out).expanduser()
            if args.pdf_out
            else data_dir / "reports" / f"bao-cao-{month_key}.pdf"
        )
        pdf_path = str(
            render_month_pdf(
                month=month_key,
                total=total,
                by_category=cat,
                top_items=top_items,
                vs_prev=vs_prev if vs_prev and vs_prev.get("prev_month") else None,
                out_path=out_pdf,
            )
        )

    out = {
        "ok": True,
        "file": str(path),
        "exists": path.exists(),
        "scope": scope,
        "count": len(rows),
        "total": total,
        "total_fmt": fmt_vnd(total),
        "by_category": {
            k: {"amount": v, "amount_fmt": fmt_vnd(v)}
            for k, v in sorted(cat.items(), key=lambda x: -x[1])
        },
        "top_expenses": [
            {**r, "amount_fmt": fmt_vnd(r["amount"])} for r in top_items
        ],
        "vs_previous_month": vs_prev,
        "items": [{**r, "amount_fmt": fmt_vnd(r["amount"])} for r in rows],
        "chart": chart_path,
        "pdf": pdf_path,
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
