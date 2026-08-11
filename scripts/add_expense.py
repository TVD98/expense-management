#!/usr/bin/env python3
"""Append one expense row to the monthly Excel file (one sheet per day)."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

try:
    from workbook_io import DEFAULT_DIR, append_expense, month_path, parse_amount
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from workbook_io import DEFAULT_DIR, append_expense, month_path, parse_amount


def main() -> int:
    p = argparse.ArgumentParser(description="Add one expense to monthly Excel")
    p.add_argument("--date", required=True, help="YYYY-MM-DD")
    p.add_argument("--amount", required=True, help="VND or shorthand (45k, 1M)")
    p.add_argument("--category", required=True)
    p.add_argument("--note", default="")
    p.add_argument("--data-dir", default=str(DEFAULT_DIR))
    args = p.parse_args()

    try:
        d = datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print(f"Invalid date: {args.date!r} (use YYYY-MM-DD)", file=sys.stderr)
        return 1

    try:
        amount = parse_amount(args.amount)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 1

    try:
        from openpyxl import Workbook  # noqa: F401
    except ImportError:
        print("Missing openpyxl. Run: pip3 install --user openpyxl", file=sys.stderr)
        return 1

    data_dir = Path(args.data_dir).expanduser()
    path = month_path(data_dir, d)
    result = append_expense(
        path,
        d,
        amount,
        args.category.strip(),
        args.note.strip(),
    )
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
