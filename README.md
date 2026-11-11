# HỆ THỐNG TỰ ĐỘNG HÓA CƠ SỞ DỮ LIỆU 2026 – PHÒNG LÃNH SỰ

Hệ thống đọc các folder hồ sơ (hiện đã làm cho **ABTC – thẻ đi lại doanh nhân APEC**), tự nhận diện văn bản, cập nhật Excel tổng hợp, lập thống kê và báo cáo tuần. Hệ thống chạy trong **VS Code với Claude Code**. Mọi công cụ nằm trong **một file: `_HE_THONG/ABTC_Tools.ps1`**.

---

## 1. Dùng trên máy khác (cài đặt một lần)

**Yêu cầu máy:** Windows 10/11, Microsoft Office có **Excel và Word**, VS Code có cài extension **Claude Code**. Không cần cài Python.

1. **Lấy thư mục dự án:** đồng bộ OneDrive (thư mục `Documents\CƠ SỞ DỮ LIỆU 2026`) hoặc chép nguyên thư mục sang máy mới. Phải giữ nguyên cấu trúc bên trong; đặt ở đâu cũng được, công cụ tự tìm đường dẫn.
2. **Mở trong VS Code:** `File → Open Folder…` → chọn thư mục `CƠ SỞ DỮ LIỆU 2026`. Claude tự đọc `CLAUDE.md` khi mở thư mục này.
3. **Kiểm tra máy:** mở Terminal trong VS Code (PowerShell) và chạy
   ```powershell
   powershell -ExecutionPolicy Bypass -File "_HE_THONG\ABTC_Tools.ps1" kiem-tra -SuaLoi
   ```
   Mọi dòng phải báo `OK`. Tham số `-SuaLoi` tự bật tùy chọn của Word để mở PDF không bị hỏi (cần cho việc đọc PDF).
4. Hoặc chỉ cần nói với Claude: *"Chạy kiểm tra máy theo README"*.

> Nếu PowerShell báo "running scripts is disabled", luôn gọi qua `powershell -ExecutionPolicy Bypass -File …` như trên. Cách này không thay đổi cài đặt máy.

---

## 1b. Ứng dụng (giao diện, không cần Claude)

**Bấm đúp phím tắt "Hồ sơ ABTC"** (có icon, ở Desktop và thư mục gốc) hoặc file `Mở ứng dụng.bat`. Cửa sổ máy chủ tự thu nhỏ xuống taskbar, cửa sổ ứng dụng (Edge) mở ra. Không đóng cửa sổ máy chủ trong lúc dùng.
> Trên máy mới: chạy `& "_HE_THONG\ABTC_Tools.ps1" tao-phim-tat` để tạo lại phím tắt (phím tắt lưu đường dẫn tuyệt đối nên phải tạo lại theo từng máy). Icon nằm ở `_HE_THONG/app/icon.ico`. Muốn đổi màu hay vẽ lại thì sửa và chạy `_HE_THONG/app/tao_icon.ps1`. Ứng dụng chạy hoàn toàn trên máy (`localhost`), không gửi dữ liệu ra ngoài, không tốn token.

| Tab | Việc |
|---|---|
| **Hồ sơ** | **Thêm hồ sơ:** kéo folder từ File Explorer thả vào cửa sổ ứng dụng, hoặc bấm *Chọn folder…*, hoặc dán đường dẫn rồi bấm *Mở* (xem bên dưới). Danh sách folder hồ sơ (trạng thái, dòng Excel, số văn bản cần nhập, số đề xuất). Chọn một hồ sơ để xem: thông tin DN, doanh nhân và ngày sinh theo từng văn bản (ô vàng là lệch), ý kiến 4 cơ quan liên ngành, **đề xuất ghi Excel** (đánh dấu, sửa giá trị, rồi bấm *Ghi*), cảnh báo, danh sách văn bản. |
| **Thống kê** | Chọn hồ sơ rồi bấm *Lập thống kê*. Hồ sơ đã có kết quả đọc bằng AI (`ket_qua.json`) thì dùng kết quả đó, còn lại dùng phân tích theo quy tắc. File ra ở `KẾT QUẢ/THỐNG KÊ/`. |
| **Công cụ** | Kiểm tra máy, mở README, mở thư mục kết quả và sao lưu. |

- **Thêm folder hồ sơ (kéo-thả / chọn / dán đường dẫn):**
  - Folder được thêm vào **thư mục đang xem** (ô "Thư mục chứa các folder hồ sơ" ở trên). Thêm xong, ứng dụng tự chuẩn bị (đọc PDF, OCR) và phân tích rồi mở hồ sơ đó.
  - **Kéo-thả:** kéo một hay nhiều folder từ File Explorer thả vào bất kỳ chỗ nào trong cửa sổ ứng dụng. Trình duyệt đọc tệp rồi gửi cho máy chủ cục bộ sao chép vào dự án; bản gốc giữ nguyên. Nếu folder cùng tên đã có trong thư mục đang xem thì chỉ **bổ sung tệp mới**, tệp đã có mà khác nội dung sẽ được hỏi trước khi ghi đè (dùng khi hồ sơ có thêm văn bản mới). Nếu folder cùng tên đã có ở chỗ khác trong dự án, ứng dụng hỏi có mở folder sẵn có thay vì sao chép không. Thả tệp rời (không trong folder) thì ứng dụng hỏi tên folder mới.
  - **Chọn folder…:** mở hộp thoại chọn thư mục của Windows (`_HE_THONG/app/chon_folder.ps1`). Folder trong dự án thì mở ngay (chọn thư mục cha thì đổi thư mục đang xem); folder ngoài dự án thì hỏi rồi sao chép vào.
  - **Dán đường dẫn:** chấp nhận đường dẫn có dấu nháy (dạng *Copy as path*) và dạng `file:///`. Enter hoặc bấm *Mở*.
  - Ứng dụng chỉ ghi vào trong thư mục dự án; các tệp `Thumbs.db`, `desktop.ini`, `~$…` được bỏ qua.
- **Phân tích theo quy tắc + OCR của Windows:**
  - Phần chắc chắn được làm tự động: loại văn bản, số/ngày lấy từ tên file, ký hiệu lấy từ chữ, luồng trạng thái, ghép dòng Excel, phát hiện file trùng, file lạc DN và ngày sinh lệch.
  - Số đóng dấu trên văn bản liên ngành được đọc bằng OCR. Đã thử trên 10 hồ sơ mẫu: đúng khoảng 92%, và các văn bản dùng OCR đều được đánh dấu để liếc lại.
- **Dòng tô vàng** là văn bản chưa đọc được số/ngày (thường là số viết tay). Bấm *Ảnh* để xem phần đầu văn bản, nhập số và ngày, rồi bấm *Lưu*. Giá trị nhập tay được nhớ lại cho các lần sau.
- **Sửa số / ký hiệu / ngày của bất kỳ văn bản nào:** bấm biểu tượng bút ở cuối dòng (dòng vạch vàng đã mở sẵn ô nhập). Để trống cả ba ô rồi *Lưu* là bỏ giá trị nhập tay.
- **Đổi tên và đặt tên chuẩn:** nút *Đổi tên folder* ở đầu hồ sơ; biểu tượng bút cạnh tên file để đổi tên từng văn bản (Enter lưu, Esc hủy). Với văn bản đã đọc được số/ngày (OCR hoặc nhập tay) nhưng tên file chưa có, ứng dụng gợi ý tên chuẩn dạng `LN, CA (6304, 11.8.26).pdf`; nút *Đặt tên chuẩn (N file)* đổi một lượt (xem và sửa tên mới trước khi đổi). Cache (chữ PDF, ảnh, OCR, giá trị nhập tay) đi theo tên mới nên không phải đọc lại. Đóng file đang mở trong Word/Excel trước khi đổi tên.
- **Sửa thông tin hồ sơ:** nút *Sửa thông tin* cạnh tên folder để sửa tên DN, tên viết tắt, cán bộ thụ lý khi hệ thống đoán sai. Giá trị được nhớ trong `_HE_THONG/cache/<folder>/ghi_tay.json` (khóa `_ho_so`); để trống ô nào thì dùng lại giá trị tự nhận diện.
- **Tìm kiếm / lọc:** ô tìm theo tên DN hoặc folder (gõ không dấu cũng được), bộ lọc *Đang xử lý / Chờ LN / Hoàn thành / Cần nhập*. Hồ sơ có file mới, đổi hoặc xóa kể từ lần phân tích hiện nhãn *Có thay đổi* và được phân tích lại khi mở.
- **Quét & cập nhật** (nút trên đầu danh sách): sau khi bỏ thêm văn bản vào các folder hồ sơ, bấm nút này để đọc file mới (chữ PDF, OCR) và phân tích lại; có thể chọn *chỉ hồ sơ có thay đổi* hoặc *quét lại tất cả*. Kết quả hiện bảng trạng thái trước → sau của từng hồ sơ. Lưu ý: ứng dụng chỉ đọc folder **trong thư mục dự án** (`CƠ SỞ DỮ LIỆU 2026\…`); văn bản bỏ vào bản sao ở nơi khác (ví dụ `Documents\Folder Hồ sơ ABTC Mẫu…`) sẽ không được thấy – hãy thêm vào bản trong dự án (kéo-thả vào ứng dụng cũng được).
- **Ghi Excel an toàn:** ghi vào `KẾT QUẢ/CƠ SỞ DỮ LIỆU/ABTC 2026.xlsx`, tự sao lưu trước khi ghi, từ chối ghi nếu ô đã bị sửa kể từ lúc phân tích, và tự thêm sự kiện vào nhật ký. Nhớ đóng file Excel tổng hợp trước khi ghi.
- **Khi cần đọc sâu hơn** (hồ sơ phức tạp, văn bản lạ), mở VS Code và nhờ Claude Code (Mục 4).
- Mở thẳng một hồ sơ: `http://localhost:8765/?ho_so=Công ty T`; mở thẳng một tab: `?tab=thongke` hoặc `?tab=congcu`.

---

## 2. Cấu trúc thư mục

```
CƠ SỞ DỮ LIỆU 2026/
├── README.md                  ← file này
├── Mở ứng dụng.bat            ← bấm đúp để mở ứng dụng giao diện
├── CLAUDE.md                  ← quy tắc Claude tự đọc
├── DỮ LIỆU GỐC/               ← 5 file Excel gốc (CHỈ ĐỌC)
├── Báo cáo tuần/              ← báo cáo tuần gốc (CHỈ ĐỌC)
├── Folder Hồ sơ ABTC Mẫu…/    ← các folder hồ sơ (mỗi folder = 1 hồ sơ)
├── KẾT QUẢ/
│   ├── MẪU/                   ← template rỗng (không điền trực tiếp)
│   ├── CƠ SỞ DỮ LIỆU/         ← Excel tổng hợp đang dùng (ghi vào đây)
│   ├── THỐNG KÊ/              ← file thống kê
│   └── BÁO CÁO TUẦN/          ← báo cáo tuần đã lập
└── _HE_THONG/
    ├── ABTC_Tools.ps1         ← TẤT CẢ công cụ
    ├── app/                   ← ứng dụng giao diện (server.ps1 + index.html)
    ├── ABTC_LOGIC.md          ← quy tắc nhận diện văn bản và luồng trạng thái hồ sơ ABTC
    ├── TEMPLATE_MAP.md        ← template nào, điền dòng/cột nào
    ├── ABTC_nhat_ky.csv       ← nhật ký sự kiện (dùng cho báo cáo tuần)
    ├── cache/                 ← chữ/ảnh đã trích từ folder hồ sơ (tự sinh, xóa được)
    ├── thay_doi/              ← các thay đổi đã duyệt (JSON)
    ├── thong_ke/              ← dữ liệu gộp để lập thống kê (JSON)
    ├── backup/                ← sao lưu tự động trước mỗi lần ghi
    └── _luu_tru/              ← script cũ (đã gộp, không dùng)
```

---

## 3. Danh sách lệnh

Gọi từ thư mục gốc: `& "_HE_THONG\ABTC_Tools.ps1" <lệnh> [tham số]` (hoặc `powershell -ExecutionPolicy Bypass -File "_HE_THONG\ABTC_Tools.ps1" <lệnh> …`).

| Lệnh | Việc | Ví dụ |
|---|---|---|
| `tro-giup` | Hiện danh sách lệnh | `… tro-giup` |
| `kiem-tra` | Kiểm tra Excel, Word, tùy chọn PDF, file mẫu (`-SuaLoi` để tự sửa) | `… kiem-tra -SuaLoi` |
| `app` | Mở ứng dụng giao diện (giống bấm `Mở ứng dụng.bat`) | `… app` |
| `tao-phim-tat` | Tạo phím tắt "Hồ sơ ABTC" có icon ở thư mục gốc và Desktop (`-KhongDesktop` để bỏ Desktop) | `… tao-phim-tat` |
| `phan-tich` | Phân tích 1 folder hồ sơ bằng quy tắc + OCR, đề xuất thay đổi Excel (không dùng AI). Kết quả ở `_HE_THONG/cache/<folder>/phan_tich.json`. | `… phan-tich -Path "Folder Hồ sơ ABTC Mẫu\2026, ABTC, CTY Công ty T"` |
| `doi-ten-folder` | Đổi tên folder hồ sơ, cache đi theo (không phải OCR lại) | `… doi-ten-folder -Path "Folder Hồ sơ ABTC Mẫu\2026, ABTC CTY Công ty P" -TenMoi "2026, ABTC, CTY Công ty S"` |
| `doi-ten-tep` | Đổi tên một file văn bản trong folder hồ sơ (giữ phần mở rộng), cache/nhập tay đi theo | `… doi-ten-tep -Path "<folder hồ sơ>" -File "LN, CA.pdf" -TenMoi "LN, CA (6304, 11.8.26).pdf"` |
| `chuan-bi` | **Chuẩn bị folder hồ sơ:** trích chữ PDF/DOCX, cắt ảnh đầu trang PDF chưa có số, phát hiện file trùng, lập `tom_tat.md`. Chỉ xử lý file mới hoặc đã đổi (`-LamLai` để làm lại hết). Nhận 1 folder hồ sơ hoặc folder cha. | `… chuan-bi -Path "Folder Hồ sơ ABTC Mẫu"` |
| `doc-xlsx` | Đọc Excel ra text | `… doc-xlsx -Path "KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx"` |
| `doc-docx` | Đọc Word ra text (giữ bảng) | `… doc-docx -Path "<file.docx>"` |
| `pdf-chu` | Trích chữ nhiều PDF qua Word | `… pdf-chu -Paths "a.pdf" "b.pdf" -OutDir "<thư mục>"` |
| `pdf-dau-trang` | Cắt 30% đầu trang 1 PDF thành ảnh nhỏ | `… pdf-dau-trang -Path "a.pdf" -Out "a.png"` |
| `tao-mau-excel` | Tạo lại template Excel từ `DỮ LIỆU GỐC` | `… tao-mau-excel` |
| `tao-mau-word` | Tạo lại template báo cáo tuần | `… tao-mau-word -Source "Báo cáo tuần\Tháng 1\BC tuần 19.01 - 23.01.2026.docx"` |
| `chuyen-du-lieu` | Chép dữ liệu gốc vào `KẾT QUẢ/CƠ SỞ DỮ LIỆU` (giữ số dòng) | `… chuyen-du-lieu -Force -Only "ABTC 2026.xlsx"` |
| `ghi-thay-doi` | Ghi thay đổi **đã duyệt** vào Excel: tự sao lưu, từ chối nếu ô đã bị sửa | `… ghi-thay-doi -File "KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx" -Changes "_HE_THONG\thay_doi\2026-10-20_Công ty S.json"` |
| `thong-ke` | Dựng file thống kê ABTC từ JSON gộp | `… thong-ke -Data "_HE_THONG\thong_ke\2026-10-21_folder_mau.json" -Out "KẾT QUẢ\THỐNG KÊ\ABTC - Thống kê.xlsx"` |

---

## 4. Các việc thường làm (nói với Claude bằng tiếng Việt)

### 4.1. Xử lý một folder hồ sơ ABTC mới
> *"Xử lý folder hồ sơ `<đường dẫn>` theo ABTC_LOGIC, cho tôi xem bảng thay đổi trước khi ghi."*

Claude sẽ làm các bước sau:
1. `chuan-bi` folder.
2. Đọc `tom_tat.md` và các file chữ/ảnh.
3. Nhận diện văn bản, ghép với dòng Excel, áp dụng luồng trạng thái.
4. Hiện bảng thay đổi để bạn duyệt.
5. Ghi bằng `ghi-thay-doi` và thêm vào nhật ký.

### 4.2. Thống kê nhiều folder hồ sơ
> *"Thống kê toàn bộ folder `<thư mục cha>` theo mẫu ABTC."*

Claude sẽ làm các bước sau:
1. `chuan-bi` thư mục cha.
2. Giao mỗi folder cho một agent theo mẫu giao việc ở Mục 5.
3. Gộp các file `ket_qua.json` thành một file trong `_HE_THONG/thong_ke/`.
4. Chạy `thong-ke`.

### 4.3. Tạo lại template khi file gốc đổi khung
> *"Tạo lại template Excel và báo cáo tuần."* → lệnh `tao-mau-excel`, `tao-mau-word`.

### 4.4. Xem file kết quả
> *"Mở file Excel ABTC tổng hợp."*

---

## 5. Mẫu giao việc cho agent đọc một folder (tiết kiệm token)

Claude dùng mẫu này khi chạy song song nhiều folder. Nên dùng model rẻ hơn (Sonnet hoặc Haiku) cho agent.

```
Đọc hồ sơ ABTC trong folder: <ĐƯỜNG DẪN FOLDER>.
Thư mục gốc dự án: <ĐƯỜNG DẪN CƠ SỞ DỮ LIỆU 2026>. KHÔNG sửa file nào trong dự án, chỉ ghi file kết quả nêu dưới.
1. Quy tắc: đọc _HE_THONG/ABTC_LOGIC.md các Mục 2.3, 3.2, 3.2b, 3.3, 5.2, 5.4, 7, 8 (không cần đọc cả file).
2. Đọc _HE_THONG/cache/<tên folder>/tom_tat.md, rồi các file .txt và *_dau.png liệt kê trong đó.
   KHÔNG đọc PDF bằng công cụ Read. Chỉ khi ảnh đầu trang không đọc được mới xem ảnh cả trang.
   Bỏ qua nội dung báo cáo thuế/BHXH của DN, bản .docx dự thảo đã có PDF số, file TRÙNG.
3. Đối chiếu Excel tổng hợp: chạy  & "_HE_THONG\ABTC_Tools.ps1" doc-xlsx -Path "KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx" -Out "<tạm>.txt"
   rồi tìm dòng khớp (số CV xin LN ở cột J, số VB DN ở cột B, tên DN + tên doanh nhân).
4. Ghi JSON (khuôn ở Mục 6, tối đa 5 cảnh báo quan trọng nhất) vào _HE_THONG/cache/<tên folder>/ket_qua.json
5. Chỉ trả lời 1–2 dòng: trạng thái cuối, dòng Excel khớp, số cảnh báo. KHÔNG gửi lại JSON.
```

---

## 6. Khuôn JSON kết quả một hồ sơ (đầu vào của lệnh `thong-ke`)

File gộp là **một mảng** các object dưới đây.

```json
{
  "folder": "2026, ABTC CTY Công ty P", "doanh_nghiep": "Công ty Cổ phần Công ty S", "ten_viet_tat": "Công ty S", "ma_so_dn": "0123456789",
  "dia_chi": "...", "so_vb_dn": "162/CV", "ngay_vb_dn": "29/05/2026", "ngay_nhan": "17/06/2026", "can_bo": "Thủy",
  "so_luong": 2,
  "van_ban": [ { "file": "", "loai": "XIN_Y_KIEN_LN", "co_quan": "Phòng", "so_ky_hieu": "12345/VP-LSLT", "ngay": "26/06/2026", "nguon_so": "tên file", "noi_dung": "" } ],
  "ln": { "CA": {"so": "", "ngay": "", "ket_luan": ""}, "THUE": {}, "BHXH": {}, "HQ": {} },
  "don_doc": [], "bo_sung": [],
  "trang_thai": "S7_UB_CHO_PHEP",
  "su_kien": [ { "ngay": "26/06/2026", "su_kien": "XIN_Y_KIEN_LN", "so_van_ban": "12345/VP-LSLT", "trang_thai_sau": "S3_XIN_Y_KIEN_LN" } ],
  "cot_excel": { "B": "", "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", "O": "", "P": "" },
  "dong_excel_hien_co": 112,
  "khac_biet_voi_excel": [], "canh_bao": []
}
```
- Mã trạng thái và mã loại văn bản: xem `ABTC_LOGIC.md` Mục 3.2 và 5.2.
- Ngày luôn ghi `dd/mm/yyyy`. Số văn bản trong `cot_excel` theo mẫu `{số}/{ký hiệu} ngày {dd/mm/yyyy}`.

---

## 7. Tiết kiệm token

- **Luôn chạy `chuan-bi` trước:** chữ PDF lấy qua Word, số/ngày lấy từ tên file hoặc ảnh cắt đầu trang. Cách này tiết kiệm khoảng 75% token so với đọc ảnh cả trang. Lệnh chỉ xử lý file mới.
- **Đặt tên file PDF kèm số và ngày** như `LN, CA (6304, 11.8.26).pdf` hoặc `Tờ trình (139, 21.9.26).pdf`. Khi đó hệ thống gần như không cần đọc ảnh.
- Agent ghi kết quả ra file và chỉ trả lời ngắn (Mục 5).

---

## 8. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| "running scripts is disabled" | Gọi `powershell -ExecutionPolicy Bypass -File "_HE_THONG\ABTC_Tools.ps1" …` |
| Chữ tiếng Việt bị lỗi font khi chạy script | File `.ps1` phải lưu **UTF-8 with BOM** (VS Code: góc dưới phải → *Save with Encoding* → *UTF-8 with BOM*). `kiem-tra` có kiểm tra mục này. |
| `pdf-chu` bị treo | Chưa bật tùy chọn Word → chạy `kiem-tra -SuaLoi`. Lệnh tự tắt Word sau 180 giây nếu bị treo. |
| "file đang mở / đang bị khóa" | Lưu và đóng file đó trong Excel rồi chạy lại. |
| Excel báo lỗi thoáng qua khi lập thống kê | Chạy lại lệnh. Công cụ đã tự thử lại và có phương án ghi từng ô. |
| Máy khác dùng Office bản khác 16.0 | Công cụ tự tìm bản Office cao nhất trong `HKCU\Software\Microsoft\Office\<phiên bản>\Word\Options` (`kiem-tra` in ra khóa đang dùng). |
| "Không đổi được tên…" khi đổi tên file/folder | File đang mở trong Word/Excel hoặc OneDrive đang đồng bộ. Đóng file, chờ vài giây rồi thử lại. |

---

## 9. Lịch sử
- **11/11/2026 (tên DN quá dài):** Tên doanh nghiệp được cắt trước các cụm mô tả dính cùng dòng ("hoạt động theo Giấy chứng nhận…", "mã số…", "do Sở Kế hoạch…", "có trụ sở…"); tên viết tắt bỏ đuôi "- CTCP" / gạch nối thừa; tên folder gợi ý bỏ ký tự Windows không cho dùng (`/ : * ? …`) và tối đa ~60 ký tự. Vd Công ty A: gợi ý "2026, ABTC, CTY THỰC PHẨM AN THỊNH PHÁT".
- **11/11/2026 (số/ngày văn bản liên ngành trả lời):** Đọc thêm số Thuế/Hải quan/Công an theo mốc ký hiệu (chịu ký hiệu bị OCR làm méo), ngày theo chữ ký số ("Thời gian ký: …") và ngày "SAO Y", ngày d.m trong tên tệp; **đối chiếu với tờ trình / phiếu trình kết quả** (nơi liệt kê lại số, ngày từng văn bản LN) và ưu tiên số gõ trong đó khi OCR đọc sai (có cảnh báo). "SĐ…" trong tên tệp là số đến, không lấy làm số văn bản. Giấy nộp tiền NSNN và văn bản "giải trình" của DN không còn bị tính là thư trả lời LN. Kết quả trên 36 hồ sơ: Thuế thiếu số 23→0, thiếu ngày 26→0; Hải quan 13→0 / 11→0; Công an 3→0 / 9→0; còn 5 văn bản cần nhập tay.
- **11/11/2026 (mất kết nối):** Khi trang không gọi được máy chủ (máy chủ đang khởi động lại, vừa đóng cửa sổ app…), app tự thử lại 3 lần cách nhau 2 giây rồi mới báo lỗi rõ ràng "Không kết nối được máy chủ ứng dụng – mở lại app bằng lối tắt", thay vì "Failed to fetch".
- **11/11/2026 (tên cá nhân, số/ngày xin ý kiến LN – mẫu mới):** Đọc tên doanh nhân dạng "Họ tên: …" (không có "và"), "Bà X, sinh ngày: …; Hộ chiếu …; Chức danh: …", bảng trong PDF mỗi ô một dòng. Số/ngày CV xin ý kiến LN không đọc được trên bản quét thì lấy từ văn bản khác trích dẫn lại ("Phúc đáp Công văn số …/VP-LSLT ngày …"); OCR nhận thêm số dính ký hiệu ("8270VP-LSLT"), nhãn "s6:", số 2 chữ số của NV-LS; bỏ số bắt đầu bằng 0. Kho "Kho B" (25 hồ sơ): hết hồ sơ thiếu tên / số / ngày xin LN; đối chiếu ODS 9/9 số khớp.
- **11/11/2026 (đọc cấu trúc folder):** Đưa vào **một folder tổng** (vd "FOLDER TỔNG (tháng 1-10.2026)") bằng *Chọn folder…*, kéo thả hoặc dán đường dẫn: hệ thống tự tìm mọi hồ sơ (Tháng → Doanh nghiệp → thư mục loại văn bản → tệp), sao chép vào kho giữ nhóm tháng, phân tích lần lượt. Loại văn bản lấy theo tên thư mục (Công văn xin ý kiến LN, Kết quả ý kiến LN, Kết quả Ủy ban, Tờ trình UB, Đề nghị bổ sung, Không đủ điều kiện…); đọc được ảnh JPG; số trong tên tệp tự do (`CAHN_9447_24122025`, `SĐ2484`, `CV37.BHXH`); ngày ở tên folder DN dùng đối chiếu ngày tờ trình/kết quả UB; trả hồ sơ một phần ghi vào cột Trả hồ sơ. Danh sách hồ sơ tìm đệ quy và hiện nhóm tháng. Chi tiết: ABTC_LOGIC.md mục 3.1b.
- **09/11/2026 (quét & thư mục gom hồ sơ):** *Quét & cập nhật* bỏ qua thư mục chỉ chứa hồ sơ con (vd "Kho C", "Kho con A") thay vì phân tích nó như một hồ sơ rồi báo lỗi ghi file; bảng lỗi hiện đúng tên hồ sơ. Trong danh sách, các thư mục này có nhãn *Thư mục · N hồ sơ*, bấm vào thì mở vào bên trong (không tính vào bộ lọc *Cần nhập*). Lưu ý: "Failed to fetch" nghĩa là trang không gọi được máy chủ – mở lại app bằng lối tắt.
- **09/11/2026 (cán bộ thụ lý):** Nhận cán bộ khi dòng nơi nhận ghi liền không ngoặc "Lưu: VT, LSAn." (Công ty N, Công ty L, Công ty R, Công ty X trước đây để trống cột cán bộ); kiểu bỏ dấu cũ "Thuỷ" được quy về "Thủy" cho khớp Excel.
- **09/11/2026 (tên doanh nhân – Công ty M):** Đọc được doanh nhân khi phiếu trình ghi danh sách đánh số "(1) Nguyễn Văn A, sinh ngày 01/01/1980; - Chức danh: …", và khi hồ sơ mới chỉ có CV đề nghị DN bổ sung (câu "đối với 01 nhân sự của Công ty: bà …, chức vụ"). Công ty R và Công ty N đọc thêm được ngày sinh.
- **09/11/2026 (thêm nhiều folder):** Thêm nhiều folder hồ sơ một lượt bằng 3 cách: (1) *Chọn folder…* → đánh dấu ô vuông ở từng folder (có *Đánh dấu tất cả*, đánh dấu được ở nhiều thư mục khác nhau) → *Thêm N folder đã đánh dấu*; (2) kéo thả nhiều folder cùng lúc; (3) dán nhiều đường dẫn vào ô (Explorer: chọn các folder → chuột phải → *Copy as path*). App kiểm tra hết, hỏi xác nhận một lần (có tùy chọn phân tích ngay), sao chép folder ngoài dự án, phân tích lần lượt, rồi hiện bảng tổng kết và đánh dấu chọn các hồ sơ vừa thêm. Chọn một thư mục chứa nhiều hồ sơ (vd "Tháng 11") thì tự tách thành từng hồ sơ con.
- **03/11/2026 (số CV xin ý kiến LN):** Đọc được số văn bản xin ý kiến liên ngành đóng dấu trên bản PDF đã ký khi OCR làm mất dấu "/" (ví dụ `1284 NV-LS` → `1234/NV-LS`), nên cột *VB XIN Ý KIẾN LIÊN NGÀNH* trong file thống kê (bên trái cột ngày) và cột J của Excel tổng hợp có đủ số. Công ty Q: 1234/NV-LS, Công ty H: 1234/NV-LS.
- **03/11/2026 (khung thống kê mới):** File thống kê (nút *Lập thống kê*) dùng khung mới 23 cột theo `ABTC 2026 NEW.ods` (mẫu ở `KẾT QUẢ/MẪU/MẪU - ABTC 2026 (khung mới).xlsx`): mỗi loại văn bản có cột SỐ và cột NGÀY riêng (văn bản đề nghị của DN, trả hồ sơ, bổ sung, xin ý kiến LN, đôn đốc, trình UB, kết quả UB); cột tên cá nhân chỉ ghi họ tên. Chi tiết: TEMPLATE_MAP.md mục 2.1b. Sửa ghi chú "Đang xin ý kiến LN (0/4): đã có ; chờ …" thừa chữ khi chưa cơ quan nào trả lời.
- **30/10/2026 (xóa thư mục chứa hồ sơ):** Xóa được cả thư mục chứa nhiều hồ sơ (ví dụ folder "Kho C" chứa 5 hồ sơ): app liệt kê các hồ sơ bên trong và yêu cầu gõ **XOA** để xác nhận; dữ liệu đã đọc của từng hồ sơ con được cất vào `_HE_THONG/_luu_tru/cache_da_xoa`. Kho và thư mục cấp cao nhất của dự án vẫn không xóa được từ app.
- **30/10/2026 (chọn nhiều):** Chọn nhiều hồ sơ trong danh sách: Ctrl + bấm (chọn thêm/bỏ), Shift + bấm (chọn liền một dải), giữ chuột kéo qua các hồ sơ (bôi đen), Ctrl+A (chọn tất cả đang hiện), Esc (bỏ chọn), phím Delete (xóa các hồ sơ đang chọn). Khi chọn từ 2 hồ sơ: thanh *Đã chọn N hồ sơ* (Phân tích lại · Chuyển kho · Xóa) và menu chuột phải cho cả nhóm. Xóa/chuyển nhiều hồ sơ làm lần lượt, hiện tiến độ, hồ sơ nào lỗi thì liệt kê riêng.
- **30/10/2026 (kho, xóa hồ sơ):** Ô *Kho hồ sơ* thay cho ô thư mục: chọn giữa nhiều kho (mỗi kho là một thư mục trong dự án, danh sách lưu ở `_HE_THONG/kho.json`), nút ⋯ để tạo kho mới, đổi tên hiển thị, thêm thư mục đang xem làm kho, gỡ kho khỏi danh sách (không xóa thư mục). Mọi kho dùng chung Excel tổng hợp. Bấm **chuột phải** vào hồ sơ: Mở, Phân tích lại, Mở folder, Sao chép đường dẫn, Đổi tên, Chuyển sang kho khác, **Xóa hồ sơ** (chuyển vào Thùng rác của Windows; dữ liệu đã đọc cất ở `_HE_THONG/_luu_tru/cache_da_xoa`; ghi lại ở `_HE_THONG/thay_doi/xoa_ho_so.log`; không sửa Excel). Không cho xóa kho, thư mục cấp cao nhất, thư mục hệ thống, folder đang có file mở. Không cho hai kho có hồ sơ trùng tên (dữ liệu đã đọc đặt theo tên folder).
- **30/10/2026 (chọn folder):** Nút *Chọn folder…* mở trình duyệt thư mục ngay trong ứng dụng (lối tắt Dự án / Tài liệu / Màn hình / Tải về / OneDrive / các ổ đĩa, đánh dấu thư mục có văn bản, gõ hoặc dán đường dẫn). Lý do: hộp thoại Windows do máy chủ chạy nền mở ra bị Windows đặt nằm sau trình duyệt. Hộp thoại Windows cũ vẫn dùng được qua liên kết *Dùng hộp thoại Windows*.
- **30/10/2026 (bổ sung):** Quy tắc mới cho mọi hồ sơ: quét kỹ toàn bộ phiếu trình để lấy đủ thông tin (địa chỉ trụ sở – kể cả khi ghi "(cũ theo đăng ký DN)", GCN đăng ký DN, lĩnh vực, người đại diện, số đến, CCCD và hộ chiếu từng doanh nhân), điền vào chỗ còn thiếu và hiện trong ứng dụng. 6 hồ sơ cũ (Công ty V, Công ty E, Công ty D, Công ty K, Công ty C, Công ty B) nay có địa chỉ. Chi tiết: ABTC_LOGIC.md mục 4.2.
- **30/10/2026:** Thêm case hồ sơ Công ty G (phiếu trình + dự thảo trả hồ sơ + công văn đề nghị DN bản quét): đọc được số/ngày văn bản DN dạng "Công văn số…" (kể cả ngày bằng chữ), ngày nhận "(nhận hồ sơ ngày …)", danh sách doanh nhân dạng gạch đầu dòng (hồ sơ Công ty U cũng đọc thêm được 2 doanh nhân), tên file "Công văn đề nghị" là văn bản của DN. Chi tiết: ABTC_LOGIC.md mục 4.1.
- **23/10/2026 (chiều):** Thêm nút *Quét & cập nhật* (phân tích lại hồ sơ có file mới, hiện trạng thái trước → sau). Phân loại văn bản: nhận thêm tên file "thông báo… không đủ điều kiện/đk" là trả hồ sơ; PDF không đoán được theo tên thì đọc nội dung để đoán (có cảnh báo). Đọc được ngày nhận và danh sách doanh nhân từ văn bản trả hồ sơ. So khớp tên DN bỏ khoảng trắng/gạch nối ("DH - Công ty K" = "DH-Công ty K") để không báo nhầm "văn bản của DN khác".
- **23/10/2026:** Ứng dụng bản 2: làm lại giao diện (danh sách có tìm kiếm/lọc, thanh tiến trình trạng thái, hộp thoại và thông báo riêng); đổi tên folder hồ sơ và file văn bản ngay trong ứng dụng, gợi ý/đặt tên chuẩn `<tên> (<số>, <d.m.yy>)`; sửa tay tên DN / viết tắt / cán bộ; sửa số/ký hiệu/ngày cho mọi văn bản; hiện "Khác biệt với Excel"; nhãn "Có thay đổi" khi folder có file mới. Sửa lỗi: ghi chú cột N bị lặp khi nối thêm, khóa Word cố định 16.0, nhật ký gộp nhầm các phiếu trình không số và thiếu trạng thái sau sự kiện, tên DN dính chữ ô bảng, file `.doc` cũ không được báo, `Mở ứng dụng.bat` không thu nhỏ cửa sổ. Bản cũ lưu ở `_HE_THONG/_luu_tru/v1_2026-10-23 (trước nâng cấp app)/`.
- **21/10/2026:** Thêm ứng dụng giao diện web cục bộ (`Mở ứng dụng.bat`), lệnh `phan-tich` (phân tích theo quy tắc) và OCR số/ngày bằng Windows. Thêm icon, phím tắt "Hồ sơ ABTC" (`tao-phim-tat`) và vùng thêm hồ sơ bằng kéo-thả / chọn folder / dán đường dẫn trong ứng dụng.
- **21/10/2026:** Gộp toàn bộ công cụ vào `ABTC_Tools.ps1`; thêm lệnh `chuan-bi` (xử lý file mới, đọc PDF tiết kiệm token) và `kiem-tra`; viết README.
- **20–21/10/2026:** Xây dựng logic ABTC, template, chuyển dữ liệu cũ, thống kê folder mẫu (10 folder, 11 đợt hồ sơ).

---

## Về kho mã nguồn này (GitHub)

Kho này chỉ chứa **phần mềm**: script PowerShell, giao diện web cục bộ, tài liệu logic và 2 file mẫu Excel rỗng (chỉ có dòng tiêu đề). **Không** có dữ liệu hồ sơ doanh nghiệp, file Excel tổng hợp, cache, sao lưu, nhật ký (xem `.gitignore`). Các ví dụ trong tài liệu đã được **ẩn danh** (họ tên, CCCD, hộ chiếu, ngày sinh, mã số thuế, địa chỉ là giá trị giả).

**Cấu trúc kho**

| Đường dẫn | Nội dung |
|---|---|
| `Mở ứng dụng.bat` | Mở ứng dụng (máy chủ cục bộ + giao diện trên Edge) |
| `README.md` | Hướng dẫn sử dụng và lịch sử thay đổi (file này) |
| `CLAUDE.md` | Hướng dẫn cho Claude Code khi mở thư mục dự án |
| `_HE_THONG/ABTC_Tools.ps1` | Toàn bộ logic: đọc PDF/Word/ảnh (Word + OCR Windows), nhận diện văn bản, phân tích hồ sơ, ghi Excel, lập thống kê |
| `_HE_THONG/ABTC_LOGIC.md` | Quy tắc nghiệp vụ: cấu trúc Excel, loại văn bản, máy trạng thái, các mẫu văn bản đã gặp (mục 3.1b, 4.1) |
| `_HE_THONG/TEMPLATE_MAP.md` | Khung các file mẫu Excel/Word, khung thống kê 23 cột (mục 2.1b) |
| `_HE_THONG/app/` | `server.ps1` (máy chủ `localhost`), `index.html` (giao diện), `chon_folder.ps1`, icon |
| `_HE_THONG/tools/` | Công cụ dựng mẫu, đọc DOCX/XLSX, chuyển dữ liệu cũ |
| `KẾT QUẢ/MẪU/MẪU - ABTC 2026*.xlsx` | Mẫu Excel rỗng: khung tổng hợp cũ và **khung thống kê mới 23 cột** |

**Yêu cầu:** Windows 10/11, PowerShell 5.1 (có sẵn), Microsoft Office (Excel, Word), Microsoft Edge. Không cần Python/Node. File `.ps1` phải giữ mã hóa **UTF-8 có BOM**.

**Cài đặt trên máy mới**
1. Tải kho về (nút *Code → Download ZIP* hoặc `git clone`), đổi tên thư mục thành `CƠ SỞ DỮ LIỆU 2026` (đặt ở đâu cũng được).
2. Tạo thêm thư mục `DỮ LIỆU GỐC` (5 file Excel gốc) và đặt file Excel tổng hợp vào `KẾT QUẢ/CƠ SỞ DỮ LIỆU/ABTC 2026.xlsx` (hoặc dựng từ mẫu – mục 2 và `TEMPLATE_MAP.md`).
3. Chạy kiểm tra máy: `powershell -ExecutionPolicy Bypass -File "_HE_THONG\ABTC_Tools.ps1" kiem-tra -SuaLoi`
4. Bấm đúp `Mở ứng dụng.bat`. Trong ứng dụng: tạo **kho hồ sơ** (nút ⋯), rồi **Chọn folder…** / kéo thả / dán đường dẫn một **folder tổng** – hệ thống tự tìm mọi hồ sơ (kể cả cấu trúc Tháng → Doanh nghiệp → thư mục loại văn bản), sao chép vào kho, đọc và phân tích.

**Tính năng chính (bản 11/2026)**
- Đọc hồ sơ dạng phẳng hoặc **có cấu trúc thư mục** (Công văn xin ý kiến LN, Kết quả ý kiến LN, Kết quả Ủy ban, Tờ trình UB, Đề nghị bổ sung, Không đủ điều kiện…); đọc PDF, Word, ảnh JPG.
- Nhận diện số/ngày văn bản: tên tệp, lớp chữ, OCR (chịu lỗi "VP"→"NP", số dính ký hiệu…), chữ ký số, ngày sao y, và **đối chiếu tờ trình / phiếu trình** (nơi liệt kê lại số, ngày văn bản liên ngành).
- Trích tên doanh nghiệp, địa chỉ, doanh nhân (nhiều mẫu văn bản), cán bộ thụ lý; máy trạng thái hồ sơ; đề xuất ghi Excel có sao lưu.
- Nhiều **kho hồ sơ**; chọn nhiều hồ sơ (Ctrl/Shift/kéo chuột); menu chuột phải (đổi tên, chuyển kho, **xóa vào Thùng rác**); thêm nhiều folder một lượt.
- **Lập thống kê** theo khung 23 cột (số và ngày văn bản tách riêng), kèm sheet thống kê, dữ liệu mốc thời gian, văn bản, cảnh báo.