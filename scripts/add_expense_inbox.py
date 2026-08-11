#!/usr/bin/env python3
"""Append one expense to the JSON inbox (phone side of the sync flow).

Writes to ``<repo>/inbox/YYYY-MM-DD.json``. The desktop later imports these
into Excel with sync_inbox.py.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

try:
    from inbox_io import add_entry, inbox_dir_for, make_entry
    from workbook_io import parse_amount
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from inbox_io import add_entry, inbox_dir_for, make_entry
    from workbook_io import parse_amount


def main() -> int:
    p = argparse.ArgumentParser(description="Add one expense to the JSON inbox")
    p.add_argument("--date", required=True, help="YYYY-MM-DD")
    p.add_argument("--amount", required=True, help="VND or shorthand (45k, 1M)")
    p.add_argument("--category", required=True)
    p.add_argument("--note", default="")
    p.add_argument(
        "--repo",
        default=".",
        help="Path to the expense repo (contains inbox/). Default: cwd.",
    )
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

    entry = make_entry(d, amount, args.category, args.note)
    inbox_dir = inbox_dir_for(Path(args.repo))
    path = add_entry(inbox_dir, entry)

    print(
        json.dumps(
            {"ok": True, "file": str(path), "entry": entry},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
