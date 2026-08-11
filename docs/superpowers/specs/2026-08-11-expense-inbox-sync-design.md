# Expense inbox sync — Design

## Goal

Cho phép ghi chi tiêu từ điện thoại (Cloud Agent) dưới dạng file JSON theo ngày
trên repo GitHub, sau đó máy tính chạy skill để nạp (sync) các file JSON đó vào
file Excel chi tiêu cục bộ và xóa các file draft đã nạp.

## Bối cảnh

Skill `expense-tracker` hiện lưu chi tiêu vào Excel tại `~/Documents/ChiTieu/`
trên máy chạy skill. Khi chạy trên điện thoại/Cloud Agent, file Excel nằm trên
máy ảo nên không dùng được trên điện thoại. Repo `expense-management` (PUBLIC)
được dùng làm nơi trung chuyển JSON giữa điện thoại và máy tính.

> Cảnh báo: repo là PUBLIC nên dữ liệu chi tiêu trong `inbox/*.json` sẽ công khai.
> Người dùng đã chấp nhận đánh đổi này.

## Kiến trúc

Hai phía dùng chung repo qua thư mục `inbox/`:

- **Điện thoại (ghi):** parse khoản chi → hiện Final draft → ghi thẳng (không cần
  confirm riêng) vào `inbox/YYYY-MM-DD.json` → commit & push.
- **Máy tính (sync):** `git pull` → đọc mọi `inbox/*.json` → nạp vào Excel qua
  `append_expense` → xóa các file JSON đã nạp trọn vẹn → commit & push.

## Dữ liệu

Mỗi ngày một file `inbox/YYYY-MM-DD.json`, nội dung là mảng các entry:

```json
[
  {
    "id": "a1b2c3d4",
    "date": "2026-08-11",
    "amount": 15000,
    "category": "Ăn uống",
    "note": "cf",
    "created_at": "2026-08-11T04:05:00+07:00"
  }
]
```

- `id`: chuỗi ngắn ngẫu nhiên, để tránh nạp trùng khi file chưa kịp xóa.
- `date`: `YYYY-MM-DD`, khớp tên file.
- `amount`: số nguyên VND (đã parse shorthand `45k`, `1M`).
- `category`: một trong danh mục mặc định của skill.
- `note`: tùy chọn.
- `created_at`: ISO-8601 có timezone.

Ghi thêm vào mảng nếu file trong ngày đã tồn tại.

## File script (trong `scripts/`)

- `inbox_io.py` — helper: đường dẫn inbox, đọc/ghi entry, sinh id, validate entry.
- `add_expense_inbox.py` — phía điện thoại: append 1 khoản vào
  `inbox/YYYY-MM-DD.json` (tham số giống `add_expense.py`).
- `sync_inbox.py` — phía máy tính: nạp toàn bộ inbox vào Excel, xóa file đã nạp.

## Luồng lỗi

- Sync đọc từng file, từng entry:
  - Entry lỗi (thiếu amount/category, sai định dạng) → bỏ qua, báo lại.
  - Chỉ xóa file khi **tất cả** entry trong file đã nạp thành công.
  - File còn entry lỗi → giữ nguyên để người dùng sửa.
- `add_expense_inbox.py`: parse amount lỗi → thoát mã khác 0, không ghi.
- Commit/push do agent thực hiện theo skill; lỗi mạng thì retry.

## Không làm (YAGNI)

- Không tự động watch/scheduler; sync chạy khi người dùng yêu cầu.
- Không hỗ trợ sửa/xóa khoản đã ghi qua inbox (chỉ thêm mới).
- Không mã hóa dữ liệu.
