# BẢN ĐỒ TEMPLATE (KHUNG BÁO CÁO)

> **Phiên bản:** v1.0 (20/10/2026)
> **Mục đích:** mô tả các khung template trong `KẾT QUẢ/MẪU/`, cho biết dữ liệu được điền vào **dòng nào, cột nào**, và cách tạo file kết quả từ template.
> **Nguyên tắc:** không bao giờ sửa trực tiếp file trong `KẾT QUẢ/MẪU/`. Luôn sao chép ra file kết quả rồi mới điền.

---

## 1. Cấu trúc folder

```
CƠ SỞ DỮ LIỆU 2026/
├── DỮ LIỆU GỐC/                   ← 5 file Excel gốc (chỉ đọc, không sửa)
├── Báo cáo tuần/                  ← báo cáo tuần gốc (chỉ đọc)
├── Folder Hồ sơ ABTC Mẫu/   ← hồ sơ mẫu
├── _HE_THONG/                     ← logic + công cụ
│   ├── ABTC_LOGIC.md
│   ├── TEMPLATE_MAP.md            ← file này
│   ├── ABTC_Tools.ps1             ← TẤT CẢ công cụ (1 file)
│   ├── cache/                     ← chữ/ảnh đã trích từ folder hồ sơ (lệnh chuan-bi)
│   └── backup/, thay_doi/, thong_ke/
└── KẾT QUẢ/
    ├── MẪU/                       ← TEMPLATE RỖNG (không điền trực tiếp)
    ├── CƠ SỞ DỮ LIỆU/             ← file Excel đã điền dữ liệu (tạo từ template)
    └── BÁO CÁO TUẦN/              ← báo cáo tuần đã điền (tạo từ template)
```

**Quy ước tên file kết quả:**
- Cơ sở dữ liệu: bỏ tiền tố `MẪU - `. Ví dụ `KẾT QUẢ/CƠ SỞ DỮ LIỆU/ABTC 2026.xlsx`.
- Báo cáo tuần: `KẾT QUẢ/BÁO CÁO TUẦN/BC tuần {dd.mm} - {dd.mm.yyyy}.docx`, theo cách đặt tên cũ (ví dụ `BC tuần 19.01 - 23.01.2026.docx`).

---

## 2. Template Excel

**Quy ước chung cho mọi sheet:**
- **Dòng tiêu đề** (giữ nguyên): từ dòng 1 đến dòng ngay trên **dòng mẫu**.
- **Dòng mẫu** = dòng dữ liệu đầu tiên, để trống nhưng đã có sẵn viền, phông chữ và căn lề. Bản ghi đầu tiên ghi thẳng vào dòng mẫu.
- **Thêm bản ghi tiếp theo:** chèn dòng mới ngay dưới dòng dữ liệu cuối, chép định dạng từ dòng mẫu (Excel COM: `Rows(mẫu).Copy()` rồi `Rows(n).Insert()`, sau đó `ClearContents()` dòng mới trước khi ghi).
- Cột ngày / số văn bản: đặt `NumberFormat = "@"` trước khi ghi để giữ dạng chữ `dd/mm/yyyy`.
- Ô nhiều dòng: dùng xuống dòng trong ô và bật Wrap Text.
- Tiêu đề có năm (ví dụ "NĂM 2026"): khi sang năm mới thì sửa ở template, không sửa từng file.

### 2.1. `MẪU - ABTC 2026.xlsx` → logic chi tiết ở `ABTC_LOGIC.md`

| Sheet | Dòng tiêu đề | Dòng mẫu |
|---|---|---|
| `ABTC 2026` | 2–4 | **5** |

| A | B | C | D | E | F | G | H | I | J | K | L | M | N | O | P |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| STT | Số vb đến/ngày nộp | Cán bộ thụ lý | Doanh nghiệp | Địa chỉ | Số lượng đề nghị cấp thẻ | Tên cá nhân đề nghị được cấp | Trả hồ sơ | Bổ sung, giải trình | Xin ý kiến LN | Văn bản đôn đốc LN | Trình UB | Kết quả UB | Ghi chú (tình trạng) | Số thẻ đã cấp trước đó | Báo cáo sử dụng thẻ |

(J–M nằm dưới nhóm tiêu đề "Tiếp tục giải quyết".)

### 2.1b. `MẪU - ABTC 2026 (khung mới).xlsx` – khung sheet `ABTC 2026` của FILE THỐNG KÊ (từ 03/11/2026)

Tạo từ file `ABTC 2026 NEW.ods` do người dùng gửi (bỏ dữ liệu ví dụ và sheet `TEST ABTC`; sửa lỗi gõ tiêu đề "ĐÊ NGHỊ" → "ĐỀ NGHỊ", "BỔ SUNG. GIẢI TRÌNH" → "BỔ SUNG, GIẢI TRÌNH"). Dùng cho `New-ThongKe` (nút *Lập thống kê* trong ứng dụng). File Excel tổng hợp `ABTC 2026.xlsx` (ghi đề xuất) **vẫn theo khung cũ** mục 2.1.

| Sheet | Dòng tiêu đề | Dòng mẫu |
|---|---|---|
| `ABTC 2026` | 2 | **3** |

| Cột | Tiêu đề | Lấy từ (khung cũ) | Ghi |
|---|---|---|---|
| A | STT | | số thứ tự |
| B / C | SỐ VB ĐẾN / NGÀY VB ĐẾN | B cũ (văn bản đề nghị của DN) | B chỉ số/ký hiệu, C chỉ ngày `dd/mm/yyyy`; bỏ "(nhận ngày …)" |
| D | CÁN BỘ THỤ LÝ | C | |
| E | DOANH NGHIỆP | D | |
| F | ĐỊA CHỈ | E | |
| G | SỐ LƯỢNG ĐỀ NGHỊ CẤP THẺ | F | |
| H | TÊN CÁ NHÂN ĐỀ NGHỊ CẤP | G | **chỉ họ tên**, mỗi người một dòng (bỏ "Ông/Bà", bỏ chức vụ) |
| I / J | TRẢ HỒ SƠ / NGÀY TRẢ HỒ SƠ | H | số / ngày tách riêng |
| K / L | BỔ SUNG, GIẢI TRÌNH / NGÀY BỔ SUNG | I | |
| M / N | VB XIN Ý KIẾN LN / NGÀY XIN Ý KIẾN LN | J | |
| O / P | VB ĐÔN ĐỐC LN / NGÀY ĐÔN ĐỐC LN | K | nhiều lần đôn đốc: mỗi lần một dòng, số và ngày cùng thứ tự |
| Q / R | VB TRÌNH UB / NGÀY TRÌNH UB | L | |
| S / T | KẾT QUẢ UB / NGÀY CÓ KẾT QUẢ UB | M | |
| U | GHI CHÚ (TÌNH TRẠNG HỒ SƠ) | N | |
| V | SỐ THẺ ĐÃ CẤP TRƯỚC ĐÓ | O | |
| W | BÁO CÁO TÌNH HÌNH SỬ DỤNG THẺ | P | |

**Quy tắc tách (hàm `TachSoNgay`):** mẫu gốc ghi gộp "số … ngày …" vào một ô – **không làm như vậy**. Ô số chỉ có số/ký hiệu (bỏ "số", "văn bản số", "1.", "Lần 1:", ghi chú trong ngoặc); ô ngày chỉ có `dd/mm/yyyy`. Văn bản chưa có số (chỉ có ngày) → ô số để trống. Nhiều văn bản trong một ô (xuống dòng hoặc ";") → mỗi văn bản một dòng, số và ngày cùng thứ tự.

### 2.2. `MẪU - BẢO HỘ CÔNG DÂN.xlsx`

| Sheet | Dòng tiêu đề | Dòng mẫu | Cột |
|---|---|---|---|
| `BHCD VN ở nước ngoài` | 3 | **4** | A STT · B Số VB đến · C Số VB · D Đơn vị đề nghị · E Nội dung · F Số lượng đề nghị hỗ trợ · G Kết quả văn bản · H Cán bộ thụ lý |
| `BHCD NN tại VIỆT NAM` | 5 | **6** | A STT · B Số VB đến · C Số VB · D Đơn vị đề nghị · E Nội dung · F Quốc tịch · G Số lượng đề nghị hỗ trợ · H Kết quả văn bản · I Cán bộ thụ lý |

### 2.3. `MẪU - HNHTQT _ TTHC _ 2026.xlsx` (Mẫu 6 – thống kê hội nghị, hội thảo quốc tế)

| Sheet | Dòng tiêu đề | Dòng mẫu |
|---|---|---|
| `TTHC HNHTQT` | 1–7 (dòng 6 là số thứ tự cột (1)…(13); dòng 7 là mục "HỘI NGHỊ, HỘI THẢO ĐÃ TỔ CHỨC VÀ CHO PHÉP TỔ CHỨC NĂM 2026") | **8** |
| `Trang tính3` | sheet trống của file gốc, giữ nguyên | — |

Cột: A STT · B Tên/chủ đề · C Đơn vị tổ chức · D Đơn vị phối hợp · E Cấp cho phép · F Số đại biểu người Việt Nam · G Người nước ngoài ở trong nước · H Người nước ngoài từ nước ngoài vào · I Đến từ nước/tổ chức quốc tế · J Chủ đề, nội dung · K Thời gian · L Địa điểm · M Nguồn, tổng kinh phí · N Tình trạng báo cáo · O Văn bản chấp thuận.

> Tiêu đề có nhắc "DỰ KIẾN TỔ CHỨC… NĂM 2027" nhưng file gốc chưa có mục này. Khi cần, thêm một dòng mục cùng định dạng dòng 7 ở cuối bảng.

### 2.4. `MẪU - NHẬP CẢNH 2026.xlsx`

| Sheet | Dòng tiêu đề | Dòng mẫu |
|---|---|---|
| `GIA HẠN VISA NĂM 2026` | 2–4 (dòng 5 để trống như file gốc) | **6** |
| `NHẬP CẢNH 2026` | 2–4 | **6** |

Cột (cả 2 sheet): A STT · B Cán bộ thụ lý · C Số vb đến/ngày đến · D Đơn vị đề nghị · E Số lượng · F Số chuyên gia · G Số người thân · H Số đối tượng khác · I Quốc tịch · J Thời gian gia hạn visa / nhập cảnh · K Kết quả · L Ghi chú.

### 2.5. `MẪU - ĐOÀN RA, VÀO, CỬ CÔNG TÁC, VIỆC RIÊNG 2026.xlsx`

| Sheet | Dòng tiêu đề | Dòng mẫu |
|---|---|---|
| `ĐOÀN RA CẤP SỞ NGÀNH 2026` | 1–4 | **5** |
| `ĐOÀN VÀO CẤP SỞ NGÀNH 2026` | 1–4 | **5** |
| `ĐOÀN VÀO CẤP THÀNH PHỐ 2026` | 1–4 | **5** |
| `CỬ CÁN BỘ ĐI CÔNG TÁC 2026` | 2–3 | **4** |
| `CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN` | 2–4 | **5** |

- **Đoàn ra cấp sở ngành:** A STT · B Số đến/số VB · C Cán bộ thụ lý · D Cơ quan, đơn vị · E Thời gian · F Nước đến · G Số lượng · H Cán bộ diện BTVTU · I Nội dung đoàn · J Nước đến · K Thời gian · L Số thành viên · M Trưởng đoàn · N Cán bộ diện BTV TU · O NSNN · P Nguồn khác · Q Trong kế hoạch · R Ngoài kế hoạch · S Kết quả · T Báo cáo · U Số QĐ · V Báo cáo.
- **Đoàn vào cấp sở ngành:** A STT · B Số đến/số VB · C Cán bộ thụ lý · D Cơ quan, đơn vị · E Thời gian · F Số lượng · H Nội dung đoàn · I Nước đến · J Thời gian · K Số thành viên · L Trưởng đoàn · M Cán bộ diện BTV TU · N NSNN · O Nguồn khác · P Trong kế hoạch · Q Ngoài kế hoạch · R Kết quả · S Báo cáo · T Báo cáo.
- **Đoàn vào cấp thành phố:** A STT · B Tên đoàn · C Trưởng đoàn · D Đến từ nước · E Nội dung hoạt động · F Số người · G Thời gian · H Nguồn kinh phí · I Báo cáo · J Đoàn trong kế hoạch · K Ghi chú · L Kết quả · M Báo cáo.
- **Cử cán bộ đi công tác:** A STT · B Cán bộ thụ lý · C Đơn vị đề nghị · D Nội dung · E Nước đến · F Thời gian · G Tên cán bộ cử đi · H Tổng số người · I TU quản lý · J Công chức · K Viên chức · L Ngân sách 1 phần · M Ngân sách toàn phần · N Đài thọ 1 phần · O Đài thọ toàn phần · P Tự chi trả · Q VB gửi BTC TU · R QĐ/VB chấp thuận · S Báo cáo.
- **Cán bộ diện BTV TU đi việc riêng:** A STT · B Cán bộ xử lý · C Đơn vị đề nghị · D Tên cán bộ/chức vụ · E Nước đến · F Thời gian · G Kinh phí · H VB chấp thuận UBND/HĐND · I Báo cáo kết quả.

---

## 3. Template báo cáo tuần `MẪU - Báo cáo tuần.docx`

- Tạo từ bản gốc `BC tuần 19.01 - 23.01.2026.docx`. Khổ A4 ngang, giữ nguyên tiêu ngữ, phông chữ, độ rộng cột.
- **Chỗ cần thay:** `{{TU_NGAY}}`, `{{DEN_NGAY}}` ở dòng tiêu đề. Ví dụ `19/01` và `23/01/2026`.
- **Bảng chính** (bảng thứ 2) có 8 cột:

| 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
|---|---|---|---|---|---|---|---|
| Tt | Nội dung công việc | Thời hạn – Tại VB đến | Thời hạn – Lãnh đạo giao | Tiến độ triển khai / Kiến nghị | Hoàn thành đúng hạn (`Đã HT`, `Đang XL`, `Đã trình`…) | Chậm muộn, nguyên nhân | Cán bộ (`Đ/c {Tên}`) |

- **Dòng mục** (cột 1 hoặc 2 có chữ) giữ nguyên. Ngay dưới mỗi mục có **1 dòng trống mẫu**.
- **Cách điền:** mỗi cán bộ hoặc mỗi nhóm việc là một dòng. Dòng đầu điền vào dòng mẫu; cần thêm dòng thì **nhân bản dòng mẫu** (sao chép nguyên `w:tr`) và chèn ngay dưới. Mục không có việc trong tuần thì để nguyên dòng trống.

| Mục | Tên mục | Nguồn dữ liệu (sheet trong `KẾT QUẢ/CƠ SỞ DỮ LIỆU/`) |
|---|---|---|
| I | Nhiệm vụ trong tuần đã, đang thực hiện | (mục lớn, không điền) |
| 1 | Cử CB đi công tác của các sở ngành, phường, xã | `CỬ CÁN BỘ ĐI CÔNG TÁC 2026` |
| 2 / 2.1 | Đoàn ra | `ĐOÀN RA CẤP SỞ NGÀNH 2026` |
| 2.2 | Đoàn vào, tiếp khách làm việc | `ĐOÀN VÀO CẤP SỞ NGÀNH 2026` |
| 3 | Việc riêng cán bộ diện BTVTU / UBND TP quản lý | `CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN` |
| 4 | Hồ sơ cho phép sử dụng thẻ APEC | `ABTC 2026` + nhật ký `ABTC_nhat_ky.csv` (xem `ABTC_LOGIC.md` Mục 10) |
| 5 | Gia hạn, nhập cảnh | `GIA HẠN VISA NĂM 2026`, `NHẬP CẢNH 2026` |
| 6 | NVNONN, NNN trên địa bàn TP | Chưa có sheet, người dùng cung cấp |
| 7 – Văn bản phối hợp cho ý kiến | Hội nghị hội thảo quốc tế | Chưa có sheet riêng, người dùng cung cấp |
| 7 – Hồ sơ hành chính | TTHC hội nghị hội thảo | `TTHC HNHTQT` |
| 8 | Công tác NGOs / Viện trợ | Chưa có sheet |
| 9 | Công tác lễ tân đối ngoại | Chưa có sheet |
| 10 | Tiếp khách của Lãnh đạo TP | Chưa có sheet |
| 11 | Bảo hộ công dân / Di cư | `BHCD VN ở nước ngoài`, `BHCD NN tại VIỆT NAM` |
| 12 | Đoàn vào Lãnh đạo Thành phố | `ĐOÀN VÀO CẤP THÀNH PHỐ 2026` |
| 13 | Các nội dung công việc khác | Người dùng cung cấp |
| II | Phương hướng, nhiệm vụ tuần tới | Người dùng cung cấp / hồ sơ đang dở |

> **Tiêu ngữ cơ quan:** "SỞ NGOẠI VỤ" (dòng 1, chữ thường) / "**PHÒNG LÃNH SỰ**" (dòng 2, chữ đậm). Người dùng xác nhận ngày 20/10/2026. Bản gốc tháng 1/2026 ghi "VĂN PHÒNG UBND THÀNH PHỐ / PHÒNG LÃNH SỰ-LỄ TÂN ĐỐI NGOẠI" (trước khi có Sở Ngoại vụ).

---

## 4. Dữ liệu cũ đã chuyển sang `KẾT QUẢ/CƠ SỞ DỮ LIỆU/` (20/10/2026)

Toàn bộ dữ liệu của 5 file gốc đã được chép vào file kết quả, **giữ đúng số dòng như file gốc** (ví dụ hồ sơ Công ty S vẫn ở dòng 112 của ABTC). Đã đối chiếu từng ô: khớp 100%.

| File | Sheet | Dòng dữ liệu | Số dòng |
|---|---|---|---|
| ABTC 2026.xlsx | ABTC 2026 | 5–204 | 200 |
| BẢO HỘ CÔNG DÂN.xlsx | BHCD VN ở nước ngoài | 4–24 | 21 |
| | BHCD NN tại VIỆT NAM | 6–11 | 6 |
| HNHTQT _ TTHC _ 2026.xlsx | TTHC HNHTQT | 8–127 | 120 |
| NHẬP CẢNH 2026.xlsx | GIA HẠN VISA NĂM 2026 | 6–12 | 7 |
| | NHẬP CẢNH 2026 | — | 0 |
| ĐOÀN RA, VÀO, CỬ CÔNG TÁC, VIỆC RIÊNG 2026.xlsx | ĐOÀN RA CẤP SỞ NGÀNH 2026 | 5–30 | 26 |
| | ĐOÀN VÀO CẤP SỞ NGÀNH 2026 | 5–11 | 7 |
| | ĐOÀN VÀO CẤP THÀNH PHỐ 2026 | 5–8 | 4 |
| | CỬ CÁN BỘ ĐI CÔNG TÁC 2026 | 4–41 | 38 |
| | CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN | 5–36 | 32 |

- Dữ liệu được chép nguyên trạng, **chưa sửa lỗi cũ** (ngày lưu dạng số Excel, ghi "Tháng 9", đảo cột D/E ở dòng 139 ABTC…). Việc sửa làm theo cảnh báo của `ABTC_LOGIC.md` Mục 8, sau khi người dùng duyệt.
- Từ nay **mọi cập nhật ghi vào file trong `KẾT QUẢ/CƠ SỞ DỮ LIỆU/`**. Bản ghi mới thêm dưới dòng dữ liệu cuối cùng.

## 5. Công cụ

Tất cả công cụ nằm trong **một file `_HE_THONG/ABTC_Tools.ps1`** (danh sách lệnh và cách dùng đầy đủ: `README.md` ở thư mục gốc). Cấu hình file → sheet → dòng mẫu nằm ở biến `$Script:CAU_HINH` đầu file; thêm sheet/file mới thì sửa ở đó và cập nhật Mục 2.

```powershell
& "_HE_THONG\ABTC_Tools.ps1" tao-mau-excel                                  # tạo lại template Excel
& "_HE_THONG\ABTC_Tools.ps1" tao-mau-word -Source "<báo cáo tuần gốc.docx>"  # tạo lại template báo cáo tuần
& "_HE_THONG\ABTC_Tools.ps1" chuyen-du-lieu -Force                           # chép lại dữ liệu cũ
& "_HE_THONG\ABTC_Tools.ps1" ghi-thay-doi -File "<xlsx>" -Changes "<json>"   # ghi thay đổi đã duyệt
& "_HE_THONG\ABTC_Tools.ps1" thong-ke -Data "<json gộp>" -Out "<xlsx>"        # dựng file thống kê
```
- File phải lưu dạng **UTF-8 có BOM** (PowerShell 5.1 mới đọc đúng tiếng Việt).
- **File đang mở trong Excel có thể bị khóa**, kể cả không đọc được. Trước khi chạy, nhờ người dùng lưu và đóng file liên quan.
- Excel không mở được cùng lúc 2 file trùng tên. Bản sao tạm phải đặt tên khác (script chuyển dữ liệu dùng tiền tố `NGUON_`).

## 6. Lịch sử thay đổi
- **v1.2 – 20/10/2026:** Gom 5 file Excel gốc vào folder `DỮ LIỆU GỐC/`.
- **v1.1 – 20/10/2026:** Đổi tiêu ngữ báo cáo tuần thành Sở Ngoại vụ – Phòng Lãnh sự; chuyển toàn bộ dữ liệu cũ sang `KẾT QUẢ/CƠ SỞ DỮ LIỆU/`; gom cấu hình vào `template_config.json`.
- **v1.0 – 20/10/2026:** Tạo 5 template Excel (11 sheet) và 1 template báo cáo tuần.
