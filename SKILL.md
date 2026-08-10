---
name: expense-tracker
description: >-
  Logs personal daily expenses into a monthly Excel file under ~/Documents/ChiTieu
  and summarizes by day or month (1-page PDF report). Use when the user asks to
  ghi chi tiêu, thêm chi tiêu, chi tiêu hôm nay, tóm tắt tháng, thống kê tháng,
  báo cáo PDF, expense tracker, or mentions amounts like 45k / 1M for spending.
---

# Expense tracker

Personal daily expense log → **one Excel file per month**, **one sheet per day**. Currency: **VND**.

## Paths

| What | Path |
|------|------|
| Data dir | `~/Documents/ChiTieu/` |
| Monthly file | `~/Documents/ChiTieu/chi-tieu-YYYY-MM.xlsx` |
| Day sheet | tab name `DD` (e.g. `10` for ngày 10) |
| Month PDF report | `~/Documents/ChiTieu/reports/bao-cao-YYYY-MM.pdf` |
| Scripts | `~/.cursor/skills/expense-tracker/scripts/` |

Create the data dir if missing. Never commit expense files to a git repo.

## Sheet layout

- Title row: `Chi tiêu ngày DD/MM/YYYY`
- Columns: `STT | Số tiền | Danh mục | Ghi chú`
- **Danh mục** = Excel dropdown + màu nền theo danh mục
- Footer row: `Tổng` with `=SUM(...)` on Số tiền
- Styled header, number format `#,##0`, freeze panes

Category colors (soft): peach / sky / teal / rose / green / violet / gray.  
Fonts: title `Avenir Next`, body `Helvetica Neue` (macOS).

Legacy single-sheet files (`Ngày | Số tiền | …`) are auto-migrated to day sheets on next write/summary.

## Categories (default)

`Ăn uống` · `Di chuyển` · `Nhà ở` · `Mua sắm` · `Sức khỏe` · `Giải trí` · `Khác`

Map free-text hints (cafe → Ăn uống, grab → Di chuyển, …). If unclear, ask or use `Khác`.

## Amount shorthand

Parse before writing (case-insensitive):

| Input | Value |
|-------|-------|
| `45k` / `45K` | 45000 |
| `1.5k` | 1500 |
| `1M` / `1m` | 1000000 |
| `1.5M` | 1500000 |
| `45000` / `45.000` / `45,000` | 45000 |

Rules: `k`/`K` = ×1_000, `m`/`M` = ×1_000_000. Round to nearest integer VND. Reject ≤ 0.

## When to use

- User wants to **add** an expense (chat natural language OK)
- User wants **today’s list + total** or **month summary by category**

## Add expense workflow

1. Resolve **date** (default today; accept “hôm qua”, `YYYY-MM-DD`, etc.).
2. Parse **amount**, **category**, **note** from the message. Ask only for missing required fields (amount; category if cannot infer).
3. Show **Final draft** — do **not** write yet:

```text
## Final draft — chi tiêu
- Ngày: YYYY-MM-DD
- Số tiền: N VND (from shorthand if any)
- Danh mục: …
- Ghi chú: …
- File: ~/Documents/ChiTieu/chi-tieu-YYYY-MM.xlsx
```

4. Write **only** after the user replies **confirm** (or clear equivalent: “ok”, “ghi đi”).
5. Run:

```bash
python3 ~/.cursor/skills/expense-tracker/scripts/add_expense.py \
  --date YYYY-MM-DD \
  --amount <int_or_shorthand> \
  --category "Ăn uống" \
  --note "cafe"
```

6. Confirm success with path + today’s running total (optional: run summary for that day).

## Summary workflow

**Today / a day:**

```bash
python3 ~/.cursor/skills/expense-tracker/scripts/summary.py --date YYYY-MM-DD
```

**Whole month** (auto 1-page PDF — pie + bảng danh mục + top chi + so tháng trước):

```bash
python3 ~/.cursor/skills/expense-tracker/scripts/summary.py --month YYYY-MM
```

- PDF: `~/Documents/ChiTieu/reports/bao-cao-YYYY-MM.pdf`
- Skip PDF: `--no-pdf` · Custom: `--pdf-out /path/to.pdf`
- PNG pie is **off by default** (only if user asks: `--chart`)

Present totals in VND with thousand separators (e.g. `45.000`).  
Tell the user the PDF path and optionally `open` it on macOS.

Triggers: “tóm tắt tháng”, “thống kê tháng”, “báo cáo PDF”, “xuất PDF”.

## Agent rules

- Do not invent expenses the user did not state.
- Do not write Excel until confirm.
- Prefer running the scripts over ad-hoc Python.
- If `openpyxl` is missing: `pip3 install --user openpyxl`, then retry.
- If `matplotlib` is missing (charts): `pip3 install --user matplotlib`, then retry.
