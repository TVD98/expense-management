#!/usr/bin/env python3
"""Import JSON inbox entries into the monthly Excel file (desktop side).

For each ``<repo>/inbox/YYYY-MM-DD.json`` file: validate every entry, append the
valid ones to Excel, then delete the file only if the whole file imported
cleanly. Files containing an invalid entry are kept untouched for fixing.

This script only touches the filesystem/Excel. Committing and pushing the
deletions to git is done by the skill/agent afterwards.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from inbox_io import entry_date, list_inbox_files, load_entries, validate_entry
    from workbook_io import DEFAULT_DIR, append_expense, month_path
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from inbox_io import entry_date, list_inbox_files, load_entries, validate_entry
    from workbook_io import DEFAULT_DIR, append_expense, month_path


def main() -> int:
    p = argparse.ArgumentParser(description="Sync JSON inbox into Excel")
    p.add_argument(
        "--repo",
        default=".",
        help="Path to the expense repo (contains inbox/). Default: cwd.",
    )
    p.add_argument(
        "--data-dir",
        default=str(DEFAULT_DIR),
        help="Where the Excel files live. Default: ~/Documents/ChiTieu",
    )
    args = p.parse_args()

    try:
        from openpyxl import Workbook  # noqa: F401
    except ImportError:
        print("Missing openpyxl. Run: pip3 install --user openpyxl", file=sys.stderr)
        return 1

    repo = Path(args.repo).expanduser()
    inbox_dir = repo / "inbox"
    data_dir = Path(args.data_dir).expanduser()

    files = list_inbox_files(inbox_dir)
    imported = 0
    deleted: list[str] = []
    kept: list[dict] = []

    for f in files:
        try:
            entries = load_entries(f)
        except (ValueError, json.JSONDecodeError) as e:
            kept.append({"file": str(f), "error": f"cannot read: {e}"})
            continue

        errors = []
        for i, entry in enumerate(entries):
            try:
                validate_entry(entry)
            except ValueError as e:
                errors.append({"index": i, "error": str(e)})

        if errors:
            kept.append({"file": str(f), "errors": errors})
            continue

        for entry in entries:
            d = entry_date(entry)
            append_expense(
                month_path(data_dir, d),
                d,
                int(entry["amount"]),
                str(entry["category"]).strip(),
                str(entry.get("note", "")).strip(),
            )
            imported += 1

        f.unlink()
        deleted.append(str(f))

    print(
        json.dumps(
            {
                "ok": True,
                "files_seen": len(files),
                "imported": imported,
                "deleted": deleted,
                "kept": kept,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
