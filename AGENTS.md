# AGENTS.md

## Nhập chi tiêu (fast path — dùng cho điện thoại / Cloud Agent)

Khi tin nhắn chỉ là một khoản chi tiêu (ví dụ `Bánh mì 12k`, `grab 45k`,
`cà phê 25000 hôm qua`), làm ngay theo mục này. **Không đọc `SKILL.md`, không
đọc skill nào khác, không brainstorm, không tạo todo list.** Mọi thứ cần biết
đã có ở dưới.

Chỉ đọc `SKILL.md` khi yêu cầu vượt ra ngoài việc nhập một khoản: tóm tắt /
thống kê tháng, báo cáo PDF, sync inbox vào Excel, hoặc sửa chính bộ script.

### Quy tắc

- Ngày mặc định là hôm nay; hiểu `hôm qua` và `YYYY-MM-DD`.
- Tiền: `k` = ×1.000, `M` = ×1.000.000 (`12k` → 12000, `1.5M` → 1500000).
  `45.000` / `45,000` / `45000` → 45000. Từ chối số ≤ 0.
- Danh mục: `Ăn uống` · `Di chuyển` · `Nhà ở` · `Mua sắm` · `Sức khỏe` ·
  `Giải trí` · `Khác`. Suy ra từ ghi chú (bánh mì/cà phê → `Ăn uống`,
  grab/xe buýt → `Di chuyển`); không rõ thì `Khác`.
- Ghi chú: giữ nguyên chữ người dùng viết, viết hoa chữ đầu.
- Không tự bịa khoản chi người dùng không nói.
- Repo này **public**: cảnh báo trước khi ghi ghi chú riêng tư.

### Chạy

Không cần hỏi xác nhận. In bản nháp rồi ghi luôn:

```text
## Final draft — chi tiêu
- Ngày: YYYY-MM-DD
- Số tiền: N VND
- Danh mục: …
- Ghi chú: …
- Inbox: inbox/YYYY-MM-DD.json
```

Trên Cloud Agent, repo nằm ở `/workspace`; trên máy cá nhân thay bằng đường dẫn
repo của bạn.

```bash
cd /workspace
git checkout main && git pull origin main
python3 scripts/add_expense_inbox.py --repo . \
  --date YYYY-MM-DD --amount 12k --category "Ăn uống" --note "Bánh mì"
git add inbox/
git commit -m "expense: add YYYY-MM-DD <ghi chú> <số tiền>"
git push origin main
```

Ghi chi tiêu đi **thẳng vào `main`**: không tạo branch, không mở PR, chỉ sửa
`inbox/*.json`. Trả lời gọn: khoản vừa ghi + tổng trong `inbox/YYYY-MM-DD.json`
của ngày đó.
