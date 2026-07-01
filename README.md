# OT Monitor 2026 — Dashboard tự động cập nhật từ Excel

## Cấu trúc (PHẲNG — khớp đúng với repo GitHub hiện tại của bạn)

```
repo/ (root)
├── .github/workflows/
│   ├── update-dashboard.yml     ← tự parse khi Excel thay đổi
│   └── validate-excel-pr.yml    ← check Excel hợp lệ ở PR
├── 2026_attendance and overtime management sheet-1 (từ tháng 7-12).xlsx
├── dashboard_data.json          ← workflow tự sinh, ĐỪNG sửa tay
├── index.html                   ← dashboard
├── parse_excel.py
└── requirements.txt
```

⚠️ **Chỉ giữ đúng 1 file `.xlsx` ở gốc repo.**

---

## VIỆC CẦN LÀM NGAY (chỉ 3 bước, làm 1 lần)

### Bước 1 — Đưa 2 file workflow vào đúng chỗ

Đây là **thư mục duy nhất bắt buộc phải có** (GitHub yêu cầu path này chính
xác để nhận diện Actions). Không dùng kéo-thả (hay bị lỗi mất cấu trúc thư
mục) — làm theo cách này, đảm bảo đúng 100%:

1. Trên GitHub, vào repo → **Add file → Create new file**.
2. Ở ô đặt tên file, gõ **nguyên dòng này** (có dấu `/`, GitHub sẽ tự tạo
   thư mục con):
   ```
   .github/workflows/update-dashboard.yml
   ```
3. Mở file `update-dashboard.yml` trong gói này bằng Notepad/VSCode → copy
   toàn bộ nội dung → dán vào khung soạn thảo trên GitHub → **Commit**.
4. Lặp lại y hệt bước 2–3 cho file thứ 2, gõ tên:
   ```
   .github/workflows/validate-excel-pr.yml
   ```

### Bước 2 — Xoá các file cũ nằm sai/dư ở root, upload các file đúng đè lên

Trên GitHub, xoá các file cũ này (nếu có, do lần upload trước bị lỗi):
`update-dashboard.yml`, `validate-excel-pr.yml` **(bản nằm NGOÀI `.github/workflows/`)**.

Sau đó vào **Add file → Upload files**, kéo thả đè lên các file: `index.html`,
`parse_excel.py`, `requirements.txt`, `dashboard_data.json`, file `.xlsx`
(dùng đúng các file trong gói này) → Commit vào `main`.

### Bước 3 — Bật quyền cho Actions (bắt buộc, làm 1 lần)

**Settings → Actions → General → Workflow permissions** → chọn
**Read and write permissions** → Save.

*(Kiểm tra **Settings → Pages**: Source phải là "Deploy from a branch",
Branch `main`, folder **/ (root)** — không phải `/docs`, vì mọi thứ đang ở
gốc repo.)*

---

## Sau khi làm xong 3 bước trên

- Vào tab **Actions**: chạy tay 1 lần để test → bấm vào workflow
  **"Update OT Dashboard Data"** → **Run workflow** → chờ ✔ xanh (~30–60s).
- Mở dashboard, refresh (F5) → dữ liệu sẽ hiện đầy đủ (68 nhân viên), tab
  Daily click được mọi ngày trong tháng.

## Từ nay về sau, mỗi khi có Excel mới

1. Vào GitHub, mở file `.xlsx` ở gốc repo → **Upload file** → kéo đè bản mới lên → Commit vào `main`.
2. Đợi ~30–60 giây, tab Actions tự chạy xong (✔ xanh).
3. Mọi người mở lại link dashboard (F5) → thấy số liệu mới, không cần làm gì thêm.

## Chạy thử ở máy local (tuỳ chọn)

```bash
pip install -r requirements.txt
python parse_excel.py "2026_attendance and overtime management sheet-1 (từ tháng 7-12).xlsx" dashboard_data.json
python -m http.server 8000
# mở http://localhost:8000
```
