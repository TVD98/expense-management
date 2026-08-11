#!/usr/bin/env python3
"""Shared helpers for the JSON inbox used to sync expenses via a git repo.

Layout: one file per day at ``<repo>/inbox/YYYY-MM-DD.json`` whose content is a
JSON array of expense entries. The phone appends entries; the desktop imports
them into Excel and deletes the file.
"""

from __future__ import annotations

import json
import re
import secrets
from datetime import date, datetime
from pathlib import Path

INBOX_DIRNAME = "inbox"
INBOX_FILE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})\.json$")
ENTRY_FIELDS = ("id", "date", "amount", "category", "note", "created_at")


def inbox_dir_for(repo: Path) -> Path:
    return Path(repo).expanduser() / INBOX_DIRNAME


def inbox_file_for(inbox_dir: Path, d: date) -> Path:
    return Path(inbox_dir) / f"{d.isoformat()}.json"


def gen_id() -> str:
    return secrets.token_hex(4)


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def make_entry(
    d: date,
    amount: int,
    category: str,
    note: str,
    entry_id: str | None = None,
    created_at: str | None = None,
) -> dict:
    return {
        "id": entry_id or gen_id(),
        "date": d.isoformat(),
        "amount": int(amount),
        "category": category.strip(),
        "note": note.strip(),
        "created_at": created_at or now_iso(),
    }


def validate_entry(entry: dict) -> None:
    """Raise ValueError if the entry cannot be imported into Excel."""
    if not isinstance(entry, dict):
        raise ValueError("entry is not an object")

    raw_date = entry.get("date")
    if not raw_date:
        raise ValueError("missing 'date'")
    try:
        datetime.strptime(str(raw_date), "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"invalid 'date': {raw_date!r} (use YYYY-MM-DD)") from exc

    amount = entry.get("amount")
    if not isinstance(amount, int) or isinstance(amount, bool):
        raise ValueError(f"'amount' must be an integer VND, got {amount!r}")
    if amount <= 0:
        raise ValueError("'amount' must be positive")

    category = entry.get("category")
    if not isinstance(category, str) or not category.strip():
        raise ValueError("missing 'category'")


def entry_date(entry: dict) -> date:
    return datetime.strptime(str(entry["date"]), "%Y-%m-%d").date()


def load_entries(path: Path) -> list[dict]:
    """Load entries from one inbox file. Missing file → empty list."""
    path = Path(path)
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        return []
    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError(f"{path} must contain a JSON array")
    return data


def write_entries(path: Path, entries: list[dict]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(entries, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def add_entry(inbox_dir: Path, entry: dict) -> Path:
    """Append one entry to that day's inbox file, creating it if needed."""
    d = entry_date(entry)
    path = inbox_file_for(inbox_dir, d)
    entries = load_entries(path)
    entries.append(entry)
    write_entries(path, entries)
    return path


def list_inbox_files(inbox_dir: Path) -> list[Path]:
    """Return daily inbox files (YYYY-MM-DD.json), sorted by date."""
    inbox_dir = Path(inbox_dir)
    if not inbox_dir.exists():
        return []
    files = [p for p in inbox_dir.iterdir() if INBOX_FILE_RE.match(p.name)]
    return sorted(files, key=lambda p: p.name)
