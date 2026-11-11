# Cơ sở dữ liệu 2026 – Phòng Lãnh sự

Thư mục này chứa các file Excel theo dõi hồ sơ theo từng mảng việc và các báo cáo tuần (`Báo cáo tuần/`).

## Quy tắc chung
- **Dữ liệu gốc** (`DỮ LIỆU GỐC/` chứa 5 file Excel gốc, và `Báo cáo tuần/`): chỉ đọc, không sửa.
- **Kết quả** ghi vào `KẾT QUẢ/`. Đọc `_HE_THONG/TEMPLATE_MAP.md` để biết template nằm đâu, dữ liệu điền vào dòng/cột nào.
  - `KẾT QUẢ/MẪU/`: template rỗng, không điền trực tiếp.
  - `KẾT QUẢ/CƠ SỞ DỮ LIỆU/`: file Excel đã điền.
  - `KẾT QUẢ/BÁO CÁO TUẦN/`: báo cáo tuần đã điền.
- **Hồ sơ ABTC (thẻ APEC):** đọc và làm theo `_HE_THONG/ABTC_LOGIC.md` trước khi xử lý.
- Các mảng khác (Bảo hộ công dân, HNHTQT, Nhập cảnh, Đoàn ra/vào): chưa có file logic.
- Máy không có Python. Đọc file bằng `_HE_THONG/tools/read_xlsx.ps1` và `_HE_THONG/tools/read_docx.ps1`; ghi Excel/Word bằng COM qua PowerShell. Script có tiếng Việt phải lưu UTF-8 có BOM.
- Luôn cho người dùng xem bảng thay đổi trước khi ghi, và sao lưu file trước khi ghi.
