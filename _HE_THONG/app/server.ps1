# Máy chủ web CỤC BỘ cho ứng dụng Hồ sơ ABTC. Chỉ nghe tại http://localhost (không mở ra mạng).
# Chạy: bấm đúp "Mở ứng dụng.bat" ở thư mục gốc, hoặc: powershell -ExecutionPolicy Bypass -File "_HE_THONG\app\server.ps1"
param([int]$Port = 8765, [switch]$KhongMoTrinhDuyet)

$APP = $PSScriptRoot
. (Join-Path (Split-Path $APP -Parent) 'ABTC_Tools.ps1')          # nạp toàn bộ công cụ
$GOC = $Script:GOC
$EXCEL = P 'KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx'
$MAC_DINH = 'Folder Hồ sơ ABTC Mẫu'
# Máy mới chưa có file Excel tổng hợp: tạo từ mẫu rỗng (chỉ khi file chưa tồn tại – không bao giờ ghi đè dữ liệu đã nhập)
try {
  if (-not (Test-Path -LiteralPath $EXCEL)) {
    $mau = P 'KẾT QUẢ\MẪU\MẪU - ABTC 2026.xlsx'
    if (Test-Path -LiteralPath $mau) { New-Item -ItemType Directory -Force (Split-Path $EXCEL -Parent) | Out-Null; Copy-Item -LiteralPath $mau -Destination $EXCEL }
  }
} catch { Write-Host "Không tạo được file Excel tổng hợp từ mẫu: $($_.Exception.Message)" }

function Send($ctx, [byte[]]$bytes, [string]$type, [int]$code = 200) {
  $r = $ctx.Response; $r.StatusCode = $code; $r.ContentType = $type; $r.Headers.Add('Cache-Control', 'no-store')
  $r.ContentLength64 = $bytes.Length; $r.OutputStream.Write($bytes, 0, $bytes.Length); $r.OutputStream.Close()
}
function SendJson($ctx, $obj, [int]$code = 200) {
  $json = ConvertTo-Json -InputObject $obj -Depth 10 -Compress
  if ($null -eq $json -or $json -eq '') { $json = if ($obj -is [array]) { '[]' } else { 'null' } }   # PS 5.1 trả rỗng cho mảng rỗng
  Send $ctx ([Text.Encoding]::UTF8.GetBytes($json)) 'application/json; charset=utf-8' $code
}
function Q($req, [string]$name) {
  # Tự giải mã tham số trên địa chỉ theo UTF-8 (QueryString mặc định của HttpListener giải mã sai tiếng Việt)
  foreach ($kv in ($req.Url.Query.TrimStart('?') -split '&')) { $p = $kv -split '=', 2; if ($p[0] -eq $name -and $p.Count -eq 2) { return [Uri]::UnescapeDataString(($p[1] -replace '\+', ' ')) } }
  ''
}
function Body($ctx) { $sr = New-Object IO.StreamReader($ctx.Request.InputStream, [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close(); if ($t) { $t | ConvertFrom-Json } else { $null } }
function Safe([string]$rel) {
  if (-not $rel.Trim('\')) { return $GOC }
  $full = [IO.Path]::GetFullPath((Join-Path $GOC $rel)).TrimEnd('\')
  if (-not $full.StartsWith($GOC, [StringComparison]::OrdinalIgnoreCase)) { throw 'Đường dẫn nằm ngoài thư mục dự án' }
  $full
}
function Rel([string]$full) { $full.Substring($GOC.Length).TrimStart('\') }
function RelJoin([string]$cha, [string]$ten) { if ($cha.Trim('\')) { $cha.Trim('\') + '\' + $ten } else { $ten } }
function LaHoSo([string]$full) { Test-LaHoSo $full }   # hồ sơ phẳng (PDF/Word ngay trong folder) hoặc hồ sơ có cấu trúc (thư mục con theo loại văn bản) – xem ABTC_Tools.ps1

function Select-Folder([string]$batDau) {
  # Mở hộp thoại chọn thư mục trong một tiến trình PowerShell -STA riêng (xem chon_folder.ps1); trả về đường dẫn hoặc ''
  $tmp = Join-Path $env:TEMP ('abtc_chon_' + [guid]::NewGuid().ToString('n') + '.txt')
  $arg = '-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -KetQua "{1}" -BatDau "{2}"' -f (Join-Path $APP 'chon_folder.ps1'), $tmp, $batDau.TrimEnd('\')
  Start-Process powershell.exe -ArgumentList $arg -Wait -WindowStyle Hidden
  $p = ''
  if (Test-Path -LiteralPath $tmp) { $p = ([IO.File]::ReadAllText($tmp, [Text.Encoding]::UTF8)).Trim(); Remove-Item -LiteralPath $tmp -Force }
  $p
}

function Import-Folder($b) {
  # Nhận một đường dẫn folder do người dùng dán/chọn. Trong dự án: trả về rel. Ngoài dự án: báo để hỏi, hoặc sao chép vào thư mục đang xem nếu sao_chep=true.
  $p = ([string]$b.duong_dan).Trim().Trim('"').Trim()
  if ($p -match '^file:///') { $p = [Uri]::UnescapeDataString($p.Substring(8)) -replace '/', '\' }
  if (-not $p -or -not (Test-Path -LiteralPath $p -PathType Container)) { return @{ loi = "Không tìm thấy thư mục: $p"; ma = 400 } }
  $full = [IO.Path]::GetFullPath($p).TrimEnd('\'); $ten = Split-Path $full -Leaf
  $laHoSo = LaHoSo $full; $soFile = @(Get-ChildItem -LiteralPath $full -File -Recurse).Count
  if ($full.StartsWith($GOC, [StringComparison]::OrdinalIgnoreCase)) { return @{ rel = (Rel $full); ten = $ten; la_ho_so = $laHoSo; so_file = $soFile; trong_du_an = $true } }
  if (-not $b.sao_chep) { return @{ ngoai = $true; ten = $ten; la_ho_so = $laHoSo; so_file = $soFile; duong_dan = $full } }
  $dich = Safe (RelJoin ([string]$b.cha) $ten)
  $trung = Find-HoSoTrungTen $ten $dich
  if ($trung -and (LaHoSo $full)) { return @{ loi = "Kho '$trung' đã có hồ sơ tên '$ten'. Hai hồ sơ cùng tên sẽ dùng chung dữ liệu đã đọc – đổi tên folder trước khi sao chép, hoặc mở kho '$trung'."; ma = 409 } }
  New-Item -ItemType Directory -Force $dich | Out-Null
  $moi = 0; $trung = 0; $khac = 0
  foreach ($f in @(Get-ChildItem -LiteralPath $full -File -Recurse)) {
    if ($f.Name -match '^(~\$|Thumbs\.db$|desktop\.ini$)') { continue }
    $d = Join-Path $dich ($f.FullName.Substring($full.Length).TrimStart('\'))
    if (Test-Path -LiteralPath $d) { if ((Get-Item -LiteralPath $d).Length -eq $f.Length) { $trung++ } else { $khac++ }; continue }
    New-Item -ItemType Directory -Force (Split-Path $d -Parent) | Out-Null; Copy-Item -LiteralPath $f.FullName -Destination $d; $moi++
  }
  @{ rel = (Rel $dich); ten = $ten; la_ho_so = (LaHoSo $dich); sao_chep = $true; moi = $moi; trung = $trung; khac = $khac }
}

function Receive-Tep($ctx) {
  # Nhận MỘT tệp (thân request là nội dung tệp) do trình duyệt gửi khi kéo-thả folder; ghi vào <cha>\<folder>\<file>
  $req = $ctx.Request; $cha = Q $req 'cha'; $folder = Q $req 'folder'; $file = (Q $req 'file') -replace '/', '\'; $ghiDe = (Q $req 'ghi_de') -eq '1'
  if (-not $folder -or $folder -match '[\\/:*?"<>|]' -or $folder -match '^\.') { throw 'Tên folder không hợp lệ' }
  if (-not $file -or $file -match '(^|\\)\.\.(\\|$)|[:*?"<>|]') { throw 'Tên tệp không hợp lệ' }
  $thuMuc = Safe (RelJoin $cha $folder); $dest = Safe (RelJoin (Rel $thuMuc) $file)
  $tt = 'moi'
  if (Test-Path -LiteralPath $dest) { if ((Get-Item -LiteralPath $dest).Length -eq $req.ContentLength64) { $tt = 'trung' } elseif ($ghiDe) { $tt = 'ghi_de' } else { $tt = 'khac' } }
  if ($tt -eq 'moi' -or $tt -eq 'ghi_de') {
    New-Item -ItemType Directory -Force (Split-Path $dest -Parent) | Out-Null
    $fs = [IO.File]::Create($dest); try { $req.InputStream.CopyTo($fs) } finally { $fs.Close() }
  } else { $req.InputStream.CopyTo([IO.Stream]::Null) }
  @{ trang_thai = $tt; rel = (Rel $thuMuc) }
}
function CacheDir([string]$full) { Join-Path $Script:HE_THONG ('cache\' + (Split-Path $full -Leaf)) }
function ReadJson([string]$p) { if (Test-Path -LiteralPath $p) { Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null } }

function Test-FolderCu([string]$full) {
  # Folder có file mới/đổi/xóa so với lần chuẩn bị trước (so chữ ký Length|LastWrite trong manifest.json)? Chưa có manifest -> $false (chưa phân tích, không phải "cũ")
  $c = CacheDir $full; $manP = Join-Path $c 'manifest.json'; $ptP = Join-Path $c 'phan_tich.json'
  $man = ReadJson $manP; if (-not $man) { return $false }
  if ((Test-Path -LiteralPath $ptP) -and (Get-Item -LiteralPath $ptP).LastWriteTimeUtc -lt (Get-Item -LiteralPath $manP).LastWriteTimeUtc) { return $true }   # đã chuẩn bị lại nhưng chưa phân tích lại
  $fs = @(Get-TepHoSo $full | Where-Object { $_.File.Extension -ne '.doc' })   # cả tệp trong thư mục loại văn bản (khóa = đường dẫn tương đối)
  if ($fs.Count -ne @($man.PSObject.Properties).Count) { return $true }
  foreach ($x in $fs) { $f = $x.File; $m = $man.PSObject.Properties[$x.Key]; if (-not $m -or $m.Value.sig -ne "$($f.Length)|$($f.LastWriteTimeUtc.Ticks)") { return $true } }
  $false
}
function Get-DsDir([string]$relCha) {
  # Mọi hồ sơ trong thư mục đang xem, tìm đệ quy (Tháng → Doanh nghiệp…); thư mục đang xem là 1 hồ sơ thì trả chính nó
  $cha = Safe $relCha
  if (-not (Test-Path -LiteralPath $cha -PathType Container)) { return @() }
  @(Find-HoSo $cha)
}
function Invoke-QuetTatCa([string]$relCha, [bool]$tatCa) {
  # Quét thư mục: phân tích lại hồ sơ có file mới/đổi hoặc chưa phân tích (tatCa = $true: mọi hồ sơ). Trả bảng trạng thái trước/sau.
  @(Get-DsDir $relCha | ForEach-Object {
    $dir = $_   # giữ lại: trong catch, $_ là lỗi chứ không còn là folder
    if (-not (LaHoSo $dir.FullName)) { return }   # thư mục gom nhiều hồ sơ (vd "Kho C") hoặc folder rỗng: không phải hồ sơ, bỏ qua
    $ptCu = ReadJson (Join-Path (CacheDir $dir.FullName) 'phan_tich.json'); $cu = Test-FolderCu $dir.FullName
    if (-not $tatCa -and $ptCu -and -not $cu) { return }
    try { $pt = Invoke-PhanTich -Path $dir.FullName -Excel $EXCEL -Im
      [pscustomobject]@{ ten = $dir.Name; dn = $pt.ten_viet_tat; truoc = $(if ($ptCu) { $ptCu.trang_thai } else { '' }); sau = $pt.trang_thai; de_xuat = @($pt.de_xuat | Where-Object chon).Count; can_nhap = @($pt.van_ban | Where-Object can_nhap).Count; loi = '' } }
    catch { [pscustomobject]@{ ten = $dir.Name; dn = ''; truoc = ''; sau = ''; de_xuat = 0; can_nhap = 0; loi = $_.Exception.Message } }
  })
}
function Get-DsFolder([string]$relCha) {
  $dirs = Get-DsDir $relCha
  @($dirs | ForEach-Object {
    $c = CacheDir $_.FullName; $pt = ReadJson (Join-Path $c 'phan_tich.json')
    $laThuMuc = $false   # từ 11/2026 danh sách tìm hồ sơ đệ quy nên không còn mục "thư mục gom hồ sơ"
    $chaFull = Safe $relCha; $nhom = if ($_.FullName.TrimEnd('\') -ne $chaFull.TrimEnd('\')) { (Split-Path $_.FullName -Parent).Substring($chaFull.TrimEnd('\').Length).Trim('\') } else { '' }
    if ($laThuMuc) { $pt = $null }
    [pscustomobject]@{
      la_thu_muc = $laThuMuc; so_con = $(if ($laThuMuc) { @(Get-ChildItem -LiteralPath $_.FullName -Directory).Count } else { 0 })
      ten = $_.Name; rel = (Rel $_.FullName); nhom = $nhom; so_file = @(Get-TepHoSo $_.FullName).Count; co_cau_truc = [bool](Test-HoSoCoCauTruc $_.FullName)
      da_chuan_bi = (Test-Path (Join-Path $c 'manifest.json')); co_ket_qua_ai = (Test-Path (Join-Path $c 'ket_qua.json')); cu = ((-not $laThuMuc) -and (Test-FolderCu $_.FullName))
      trang_thai = $(if ($pt) { $pt.trang_thai } else { '' }); dn = $(if ($pt) { $pt.ten_viet_tat } else { '' })
      dong = $(if ($pt) { $pt.dong_ghi } else { $null }); dong_moi = [bool]($pt -and -not $pt.dong_excel_hien_co)
      canh_bao = $(if ($pt) { @($pt.canh_bao).Count } else { 0 }); can_nhap = $(if ($pt) { @($pt.van_ban | Where-Object can_nhap).Count } else { 0 })
      de_xuat = $(if ($pt) { @($pt.de_xuat | Where-Object chon).Count } else { 0 }); ten_chuan = $(if ($pt) { @($pt.van_ban | Where-Object ten_goi_y).Count } else { 0 })
    }
  })
}

# ================= Kho hồ sơ: nhiều thư mục chứa hồ sơ trong dự án (danh sách lưu ở _HE_THONG\kho.json) =================
$KHO_FILE = Join-Path $Script:HE_THONG 'kho.json'
function Get-KhoDs {
  $ds = @()
  # Lưu ý PS 5.1: ConvertFrom-Json trả cả mảng như MỘT đối tượng → phải trải ra bằng ForEach-Object, nếu không các kho bị gộp làm một
  if (Test-Path -LiteralPath $KHO_FILE) { try { $j = Get-Content -LiteralPath $KHO_FILE -Raw -Encoding UTF8 | ConvertFrom-Json; $ds = @($j | ForEach-Object { $_ }) } catch { $ds = @() } }
  elseif (Test-Path -LiteralPath (P $MAC_DINH)) { $ds = @([pscustomobject]@{ ten = $MAC_DINH; rel = $MAC_DINH }) }   # lần đầu: kho mặc định
  @($ds | Where-Object { $_ -and $_.rel } | ForEach-Object {
    $full = P $_.rel; $co = Test-Path -LiteralPath $full -PathType Container
    [pscustomobject]@{ ten = [string]$_.ten; rel = [string]$_.rel; ton_tai = $co; so_ho_so = $(if ($co) { @(Get-ChildItem -LiteralPath $full -Directory).Count } else { 0 }) }
  })
}
function Save-KhoDs($ds) {
  $j = ConvertTo-Json -InputObject @(@($ds) | ForEach-Object { [ordered]@{ ten = [string]$_.ten; rel = [string]$_.rel } }) -Depth 3
  if (-not $j) { $j = '[]' }
  [IO.File]::WriteAllText($KHO_FILE, $j, (New-Object Text.UTF8Encoding($false)))
}
function Set-Kho($b) {
  # hanh_dong: tao (tạo thư mục mới + đăng ký) | them (đăng ký thư mục có sẵn) | doi-ten (đổi tên hiển thị) | go (gỡ khỏi danh sách, không xóa thư mục)
  $ds = @(Get-KhoDs); $rel = ''
  switch ([string]$b.hanh_dong) {
    'tao' {
      $ten = ([string]$b.ten).Trim()
      if (-not (Test-TenHopLe $ten)) { throw "Tên kho không hợp lệ: '$ten' (không dùng \ / : * ? `" < > |, không bắt đầu bằng dấu chấm)" }
      if ($ten -match '^_HE_THONG$|^KẾT QUẢ$|^DỮ LIỆU GỐC$|^Báo cáo tuần$') { throw "Không dùng tên '$ten' – trùng thư mục hệ thống của dự án" }
      $full = Safe $ten
      if (Test-Path -LiteralPath $full) { throw "Đã có thư mục '$ten' trong dự án. Mở thư mục đó rồi chọn 'Thêm thư mục đang xem làm kho'." }
      New-Item -ItemType Directory -Path $full | Out-Null
      $rel = $ten; $ds += [pscustomobject]@{ ten = $ten; rel = $ten }
    }
    'them' {
      $full = Safe ([string]$b.rel); $rel = Rel $full
      if (-not $rel -or $full.StartsWith($Script:HE_THONG, [StringComparison]::OrdinalIgnoreCase)) { throw 'Không dùng thư mục này làm kho' }
      if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw "Không tìm thấy thư mục: $rel" }
      if (-not ($ds | Where-Object { $_.rel -eq $rel })) { $ten = ([string]$b.ten).Trim(); if (-not $ten) { $ten = Split-Path $full -Leaf }; $ds += [pscustomobject]@{ ten = $ten; rel = $rel } }
    }
    'doi-ten' {
      $rel = [string]$b.rel; $ten = ([string]$b.ten).Trim(); if (-not $ten) { throw 'Chưa nhập tên kho' }
      foreach ($k in $ds) { if ($k.rel -eq $rel) { $k.ten = $ten } }
    }
    'go' { $rel = [string]$b.rel; $ds = @($ds | Where-Object { $_.rel -ne $rel }) }
    default { throw 'Thao tác kho không hợp lệ' }
  }
  Save-KhoDs $ds
  @{ ds = @(Get-KhoDs); rel = $rel }
}
function Find-HoSoTrungTen([string]$ten, [string]$truFull) {
  # Cache đặt theo TÊN folder hồ sơ → hai hồ sơ cùng tên ở hai kho sẽ dùng chung cache (sai dữ liệu). Trả tên kho đang có hồ sơ trùng tên, '' nếu không.
  foreach ($k in Get-KhoDs) {
    if (-not $k.ton_tai) { continue }
    $d = Join-Path (P $k.rel) $ten
    if ((Test-Path -LiteralPath $d -PathType Container) -and ([IO.Path]::GetFullPath($d).TrimEnd('\') -ne $truFull.TrimEnd('\')) -and (Test-LaHoSo $d)) { return $k.ten }   # folder rỗng cùng tên thì không tính
  }
  ''
}
$CS_THUNG_RAC = @'
using System;
using System.Runtime.InteropServices;
public static class AbtcThungRac {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
  struct SHFILEOPSTRUCT { public IntPtr hwnd; public uint wFunc; public string pFrom; public string pTo; public ushort fFlags; public bool fAnyOperationsAborted; public IntPtr hNameMappings; public string lpszProgressTitle; }
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)] static extern int SHFileOperation(ref SHFILEOPSTRUCT op);
  // Chuyển vào Thùng rác, không hỏi, không hiện hộp thoại (tránh hộp thoại bị nằm sau trình duyệt làm treo máy chủ)
  public static int Chuyen(string path) {
    var op = new SHFILEOPSTRUCT(); op.wFunc = 3; op.pFrom = path + "\0\0";
    op.fFlags = 0x0040 | 0x0010 | 0x0004 | 0x0400;   // ALLOWUNDO | NOCONFIRMATION | SILENT | NOERRORUI
    int r = SHFileOperation(ref op); if (r == 0 && op.fAnyOperationsAborted) r = -1; return r;
  }
}
'@
function Get-HoSoCon([string]$full) {
  # Folder không có PDF/Word trực tiếp nhưng có thư mục con = thư mục chứa nhiều hồ sơ → trả tên các thư mục con; hồ sơ thường → @()
  if (LaHoSo $full) { return @() }
  @(Find-HoSo $full | ForEach-Object { $_.Name })   # mọi hồ sơ bên trong (kể cả lồng Tháng → Doanh nghiệp)
}
function Remove-HoSo([string]$rel, [switch]$CaThuMuc) {
  $full = Safe $rel
  if ($full -eq $GOC -or $full.StartsWith($Script:HE_THONG, [StringComparison]::OrdinalIgnoreCase)) { throw 'Không được xóa thư mục này' }
  if (Get-KhoDs | Where-Object { (P $_.rel).TrimEnd('\') -eq $full }) { throw 'Đây là một kho, không phải hồ sơ. Muốn bỏ kho thì dùng "Gỡ kho khỏi danh sách".' }
  if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw "Không tìm thấy folder: $rel" }
  # Chặn độc lập với danh sách kho: thư mục cấp cao nhất của dự án (kho, KẾT QUẢ, DỮ LIỆU GỐC…) không phải hồ sơ
  if ((Split-Path $full -Parent).TrimEnd('\') -eq $GOC.TrimEnd('\')) { throw 'Đây là thư mục cấp cao nhất của dự án (kho hoặc thư mục hệ thống), không phải một hồ sơ – không xóa từ ứng dụng.' }
  # Hồ sơ = folder có PDF/Word trực tiếp bên trong, hoặc folder rỗng. Folder chỉ chứa thư mục con là thư mục chứa nhiều hồ sơ.
  # Thư mục chứa nhiều hồ sơ chỉ được xóa khi người dùng đã xác nhận riêng (ca_thu_muc = true, giao diện yêu cầu gõ XOA)
  $con = @(Get-HoSoCon $full)
  if ($con.Count -and -not $CaThuMuc) { throw 'Folder này chứa nhiều hồ sơ – cần xác nhận riêng.' }
  $gioiHan = if ($CaThuMuc) { 300 } else { 20 }
  if (@(Get-ChildItem -LiteralPath $full -Directory -Recurse).Count -gt $gioiHan) { throw "Folder này chứa hơn $gioiHan thư mục con – không xóa từ ứng dụng. Xóa trong Explorer nếu chắc chắn." }
  $khoa = @(Get-ChildItem -LiteralPath $full -File -Recurse | Where-Object { try { $s = [IO.File]::Open($_.FullName, 'Open', 'ReadWrite', 'None'); $s.Close(); $false } catch { $true } } | ForEach-Object { $_.Name })
  if ($khoa.Count) { throw ('Đang có file mở hoặc OneDrive đang đồng bộ: ' + (($khoa | Select-Object -First 3) -join ', ') + '. Đóng file rồi thử lại.') }
  if (-not ('AbtcThungRac' -as [type])) { Add-Type -TypeDefinition $CS_THUNG_RAC }
  $r = [AbtcThungRac]::Chuyen($full)
  if ($r -ne 0 -or (Test-Path -LiteralPath $full)) { throw "Không chuyển được vào Thùng rác (mã $r). Đóng các file trong folder rồi thử lại." }
  # Cất cache (chữ PDF, OCR, giá trị nhập tay) vào lưu trữ – trừ khi kho khác còn hồ sơ cùng tên đang dùng chung cache
  # Thư mục chứa nhiều hồ sơ: cất cache của từng hồ sơ con (cache đặt theo tên hồ sơ con)
  $ten = Split-Path $full -Leaf; $luuDs = @(); $tg = Get-Date -Format 'yyyyMMdd-HHmmss'
  foreach ($t in @($ten) + $con) {
    $c = Join-Path $Script:HE_THONG ('cache\' + $t)
    if ((Test-Path -LiteralPath $c) -and -not (Find-HoSoTrungTen $t '')) {
      $dst = Join-Path $Script:HE_THONG ('_luu_tru\cache_da_xoa\' + $t + ' (' + $tg + ')')
      New-Item -ItemType Directory -Force (Split-Path $dst -Parent) | Out-Null; Move-Item -LiteralPath $c -Destination $dst; $luuDs += (Rel $dst)
    }
  }
  $luu = $luuDs -join '; '
  $log = Join-Path $Script:HE_THONG 'thay_doi\xoa_ho_so.log'; New-Item -ItemType Directory -Force (Split-Path $log -Parent) | Out-Null
  Add-Content -LiteralPath $log -Value ("{0}`tĐã chuyển vào Thùng rác: {1}`tcache: {2}" -f (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'), $rel, $luu) -Encoding UTF8
  @{ ok = $true; ten = $ten; cache = $luu }
}
function Move-HoSoKho([string]$rel, [string]$khoRel) {
  $full = Safe $rel; $dichCha = Safe $khoRel
  if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw "Không tìm thấy folder: $rel" }
  if (-not (Test-Path -LiteralPath $dichCha -PathType Container)) { throw "Không tìm thấy kho: $khoRel" }
  $dich = Join-Path $dichCha (Split-Path $full -Leaf)
  if ((Split-Path $full -Parent) -eq $dichCha) { throw 'Hồ sơ đã ở kho này' }
  if (Test-Path -LiteralPath $dich) { throw "Kho đích đã có folder '$(Split-Path $full -Leaf)'" }
  try { Move-Item -LiteralPath $full -Destination $dich -ErrorAction Stop }
  catch { throw "Không chuyển được – có file đang mở trong Word/Excel hoặc OneDrive đang đồng bộ. Đóng file rồi thử lại. ($($_.Exception.Message))" }
  @{ rel = (Rel $dich) }
}

function Add-NhatKy($o) {
  $csv = Join-Path $Script:HE_THONG 'ABTC_nhat_ky.csv'
  $hdr = 'ngay_su_kien,ngay_xu_ly,dong_excel,doanh_nghiep,ten_viet_tat,can_bo,su_kien,so_van_ban,trang_thai_truoc,trang_thai_sau,file_nguon'
  $lines = if (Test-Path $csv) { @([IO.File]::ReadAllLines($csv, [Text.Encoding]::UTF8)) } else { @($hdr) }
  # Khóa chống trùng: dòng Excel | ngày | sự kiện | số văn bản (có ngày để các phiếu trình không số không bị gộp làm một)
  $keys = @{}; foreach ($l in $lines | Select-Object -Skip 1) { $f = ($l -replace '^"', '' -replace '"$', '') -split '","'; if ($f.Count -ge 8) { $keys["$($f[2])|$($f[0])|$($f[6])|$($f[7])"] = 1; $keys["$($f[2])|$($f[6])|$($f[7])"] = 1 } }
  $q = { param($s) '"' + ([string]$s -replace '"', '""') + '"' }; $them = 0; $truoc = ''
  foreach ($e in @($o.su_kien)) {
    $k = "$($o.dong_ghi)|$($e.ngay)|$($e.su_kien)|$($e.so_van_ban)"
    if ($keys.ContainsKey($k) -or ($e.so_van_ban -and $keys.ContainsKey("$($o.dong_ghi)|$($e.su_kien)|$($e.so_van_ban)"))) { $truoc = $e.trang_thai_sau; continue }   # khóa cũ (không ngày) chỉ áp cho sự kiện có số văn bản
    $lines += (@($e.ngay, (Get-Date -Format 'dd/MM/yyyy'), $o.dong_ghi, $o.doanh_nghiep, $o.ten_viet_tat, $o.can_bo, $e.su_kien, $e.so_van_ban, $truoc, $e.trang_thai_sau, $o.folder) | ForEach-Object { & $q $_ }) -join ','
    $keys[$k] = 1; $them++; $truoc = $e.trang_thai_sau
  }
  [IO.File]::WriteAllLines($csv, [string[]]$lines, (New-Object Text.UTF8Encoding($true))); $them
}

function Handle($ctx) {
  $req = $ctx.Request; $path = $req.Url.AbsolutePath; $m = $req.HttpMethod
  if ($path -eq '/' -or $path -eq '/index.html') { Send $ctx ([IO.File]::ReadAllBytes((Join-Path $APP 'index.html'))) 'text/html; charset=utf-8'; return }
  $icons = @{ '/icon.svg' = @('icon.svg', 'image/svg+xml'); '/favicon.ico' = @('icon.ico', 'image/x-icon'); '/icon.png' = @('icon.png', 'image/png') }
  if ($icons.ContainsKey($path)) { $f = Join-Path $APP $icons[$path][0]; if (Test-Path $f) { Send $ctx ([IO.File]::ReadAllBytes($f)) $icons[$path][1] } else { Send $ctx ([byte[]]@()) 'text/plain' 404 }; return }
  switch ("$m $path") {
    'GET /api/tong-quan' {
      SendJson $ctx @{ goc = $GOC; excel = (Rel $EXCEL); co_excel = (Test-Path -LiteralPath $EXCEL); thu_muc_mac_dinh = $MAC_DINH; word_pdf = (Test-WordPdf) }; return }
    'GET /api/ds-folder' { SendJson $ctx (Get-DsFolder ((Q $req 'cha'))); return }
    'POST /api/chuan-bi' { $b = Body $ctx; $out = Invoke-ChuanBiFolder -Path (Safe $b.rel) | Out-String; SendJson $ctx @{ log = $out }; return }
    'POST /api/quet' { $b = Body $ctx; SendJson $ctx @{ ds = (Invoke-QuetTatCa ([string]$b.cha) ([bool]$b.tat_ca)) }; return }
    'POST /api/phan-tich' { $b = Body $ctx; SendJson $ctx (Invoke-PhanTich -Path (Safe $b.rel) -Excel $EXCEL -Im); return }
    'GET /api/phan-tich' {
      # Trả kết quả đã có; phân tích lại khi chưa có hoặc folder đã có file mới/đổi kể từ lần trước
      $full = Safe ((Q $req 'rel')); $pt = ReadJson (Join-Path (CacheDir $full) 'phan_tich.json')
      if (-not $pt -or (Test-FolderCu $full)) { $pt = Invoke-PhanTich -Path $full -Excel $EXCEL -Im }; SendJson $ctx $pt; return }
    'POST /api/ghi-tay' {
      # file = tên văn bản: lưu số / ký hiệu / ngày nhập tay (rỗng cả 3 -> xóa). file = "_ho_so": lưu doanh_nghiep / ten_viet_tat / can_bo cho cả hồ sơ.
      $b = Body $ctx; $full = Safe $b.rel; $c = CacheDir $full; $p = Join-Path $c 'ghi_tay.json'
      $h = @{}; $old = ReadJson $p; if ($old) { $old.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value } }
      if ($b.file -eq '_ho_so') {
        if (-not $b.doanh_nghiep -and -not $b.ten_viet_tat -and -not $b.can_bo) { $h.Remove('_ho_so') } else { $h['_ho_so'] = @{ doanh_nghiep = [string]$b.doanh_nghiep; ten_viet_tat = [string]$b.ten_viet_tat; can_bo = [string]$b.can_bo } }
      } elseif (-not $b.so -and -not $b.ngay -and -not $b.ky_hieu) { $h.Remove($b.file) } else { $h[$b.file] = @{ so = [string]$b.so; ngay = [string]$b.ngay; ky_hieu = [string]$b.ky_hieu } }
      New-Item -ItemType Directory -Force $c | Out-Null
      [IO.File]::WriteAllText($p, ($h | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))
      SendJson $ctx (Invoke-PhanTich -Path $full -Excel $EXCEL -Im); return }
    'POST /api/doi-ten-folder' {
      $b = Body $ctx; $full = Safe $b.rel
      $trung = Find-HoSoTrungTen ([string]$b.ten_moi) $full
      if ($trung) { SendJson $ctx @{ loi = "Kho '$trung' đã có hồ sơ tên '$($b.ten_moi)'. Hai hồ sơ cùng tên sẽ dùng chung dữ liệu đã đọc – hãy đặt tên khác." } 409; return }
      try { $moi = Rename-HoSoFolder -Path $full -TenMoi ([string]$b.ten_moi) } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409; return }
      SendJson $ctx @{ rel = (Rel $moi); phan_tich = (Invoke-PhanTich -Path $moi -Excel $EXCEL -Im) }; return }
    'POST /api/doi-ten-tep' {
      $b = Body $ctx; $full = Safe $b.rel
      try { $moi = Rename-HoSoTep -Path $full -File ([string]$b.file) -TenMoi ([string]$b.ten_moi) } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409; return }
      SendJson $ctx @{ file = $moi; phan_tich = (Invoke-PhanTich -Path $full -Excel $EXCEL -Im) }; return }
    'POST /api/doi-ten-nhieu' {
      # Đổi tên nhiều file một lượt (đặt tên chuẩn); phân tích lại một lần cuối; trả danh sách thành công / lỗi
      $b = Body $ctx; $full = Safe $b.rel; $ok = @(); $loi = @()
      foreach ($d in @($b.ds)) { try { [void](Rename-HoSoTep -Path $full -File ([string]$d.file) -TenMoi ([string]$d.ten_moi)); $ok += [string]$d.ten_moi } catch { $loi += "$($d.file): $($_.Exception.Message)" } }
      SendJson $ctx @{ ok = $ok; loi = $loi; phan_tich = (Invoke-PhanTich -Path $full -Excel $EXCEL -Im) }; return }
    'POST /api/ghi' {
      $b = Body $ctx; $full = Safe $b.rel
      $list = @($b.thay_doi | ForEach-Object { [pscustomobject]@{ sheet = 'ABTC 2026'; o = $_.o; cu = $_.cu; moi = $_.moi } })
      if (-not $list.Count) { SendJson $ctx @{ loi = 'Chưa chọn ô nào' } 400; return }
      $pt0 = ReadJson (Join-Path (CacheDir $full) 'phan_tich.json'); $ten = if ($pt0) { ($pt0.ten_viet_tat -replace '[\\/:*?"<>|]', '') } else { 'ho_so' }
      $dirTd = Join-Path $Script:HE_THONG 'thay_doi'; New-Item -ItemType Directory -Force $dirTd | Out-Null
      $jf = Join-Path $dirTd ("{0}_{1}.json" -f (Get-Date -Format 'yyyy-MM-dd_HHmm'), $ten)
      [IO.File]::WriteAllText($jf, (ConvertTo-Json -InputObject $list -Depth 4), (New-Object Text.UTF8Encoding($false)))
      try { $log = Set-ThayDoi -File $EXCEL -Changes $jf | Out-String } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409; return }
      $pt = Invoke-PhanTich -Path $full -Excel $EXCEL -Im; $nk = Add-NhatKy $pt
      SendJson $ctx @{ log = $log; nhat_ky = $nk; file_thay_doi = (Rel $jf); phan_tich = $pt }; return }
    'POST /api/thong-ke' {
      $b = Body $ctx; $recs = @(); $nguon = @()
      foreach ($r in @($b.rels)) {
        $full = Safe $r; $c = CacheDir $full; $ai = ReadJson (Join-Path $c 'ket_qua.json')
        if ($ai) { $recs += @($ai | ForEach-Object { $_ }); $nguon += 'AI' } else { $recs += (Invoke-PhanTich -Path $full -Excel $EXCEL -Im); $nguon += 'quy tắc' }
      }
      if (-not $recs.Count) { SendJson $ctx @{ loi = 'Chưa chọn hồ sơ nào' } 400; return }
      $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'; $dirTk = Join-Path $Script:HE_THONG 'thong_ke'; New-Item -ItemType Directory -Force $dirTk | Out-Null
      $jd = Join-Path $dirTk "${stamp}_app.json"; [IO.File]::WriteAllText($jd, (ConvertTo-Json -InputObject $recs -Depth 10), (New-Object Text.UTF8Encoding($false)))
      $out = P ("KẾT QUẢ\THỐNG KÊ\ABTC - Thống kê " + (Get-Date -Format 'dd.MM.yyyy HHmm') + '.xlsx')
      $log = New-ThongKe -Data $jd -Out $out -Nguon ("{0} hồ sơ ({1} theo AI, {2} theo quy tắc)" -f $recs.Count, @($nguon | Where-Object { $_ -eq 'AI' }).Count, @($nguon | Where-Object { $_ -ne 'AI' }).Count) | Out-String
      SendJson $ctx @{ log = $log; file = (Rel $out) }; return }
    'POST /api/mo' { $b = Body $ctx; Start-Process -FilePath (Safe $b.rel); SendJson $ctx @{ ok = $true }; return }
    'POST /api/chon-folder' {
      $b = Body $ctx; $bd = $GOC; try { $bd = Safe ([string]$b.cha); if (-not (Test-Path -LiteralPath $bd)) { $bd = $GOC } } catch { }
      SendJson $ctx @{ duong_dan = (Select-Folder $bd) }; return }
    'POST /api/nhap-folder' { $r = Import-Folder (Body $ctx); if ($r.loi) { SendJson $ctx @{ loi = $r.loi } $r.ma } else { SendJson $ctx $r }; return }
    'POST /api/tai-len' { SendJson $ctx (Receive-Tep $ctx); return }
    'GET /api/kho' { SendJson $ctx @{ ds = @(Get-KhoDs) }; return }
    'POST /api/kho' { try { SendJson $ctx (Set-Kho (Body $ctx)) } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409 }; return }
    'POST /api/xoa-ho-so' {
      $b = Body $ctx
      try {
        # Thư mục chứa nhiều hồ sơ: trả danh sách hồ sơ con để giao diện hỏi xác nhận riêng
        $full = Safe ([string]$b.rel)
        if (-not $b.ca_thu_muc -and (Test-Path -LiteralPath $full -PathType Container)) {
          $con = @(Get-HoSoCon $full)
          if ($con.Count) { SendJson $ctx @{ loi = "Folder '$(Split-Path $full -Leaf)' chứa $($con.Count) hồ sơ – cần xác nhận riêng"; thu_muc_chua = $true; con = $con } 409; return }
        }
        SendJson $ctx (Remove-HoSo ([string]$b.rel) -CaThuMuc:([bool]$b.ca_thu_muc))
      } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409 }; return }
    'POST /api/chuyen-kho' {
      $b = Body $ctx
      try {
        $trung = Find-HoSoTrungTen (Split-Path (Safe ([string]$b.rel)) -Leaf) (Safe ([string]$b.rel))
        if ($trung) { throw "Kho '$trung' đã có hồ sơ cùng tên" }
        SendJson $ctx (Move-HoSoKho ([string]$b.rel) ([string]$b.kho))
      } catch { SendJson $ctx @{ loi = $_.Exception.Message } 409 }; return }
    'GET /api/tim-ho-so' {
      # Mọi hồ sơ trong một folder tổng (trong hoặc ngoài dự án): ?p=<đường dẫn đầy đủ> hoặc ?rel=<trong dự án>. nhom = đường dẫn tương đối của thư mục cha tính từ folder tổng (vd "THÁNG 1")
      $p = (Q $req 'p').Trim().Trim('"'); if (-not $p) { $p = Safe (Q $req 'rel') }
      if (-not (Test-Path -LiteralPath $p -PathType Container)) { SendJson $ctx @{ loi = "Không tìm thấy thư mục: $p" } 400; return }
      # Lưu ý: PowerShell không phân biệt hoa/thường – không đặt tên $goc (sẽ đè $GOC = thư mục dự án)
      $gocTim = [IO.Path]::GetFullPath($p).TrimEnd('\')
      $ds = @(Find-HoSo $gocTim | ForEach-Object {
        $cha = (Split-Path $_.FullName -Parent).TrimEnd('\')
        [pscustomobject]@{ ten = $_.Name; duong_dan = $_.FullName; nhom = $(if ($_.FullName.TrimEnd('\') -eq $gocTim) { '' } elseif ($cha.Length -gt $gocTim.Length) { $cha.Substring($gocTim.Length + 1) } else { '' })
          so_tep = @(Get-TepHoSo $_.FullName).Count; co_cau_truc = [bool](Test-HoSoCoCauTruc $_.FullName); trong_du_an = $_.FullName.StartsWith($GOC, [StringComparison]::OrdinalIgnoreCase) }
      })
      SendJson $ctx @{ goc = $gocTim; ten_goc = (Split-Path $gocTim -Leaf); ds = $ds }; return }
    'GET /api/duyet' {
      # Trình duyệt thư mục trong ứng dụng (thay hộp thoại Windows hay bị nằm sau cửa sổ trình duyệt). Chỉ đọc danh sách thư mục, không đọc nội dung tệp.
      $noi = @()
      foreach ($x in @(@('Dự án', $GOC), @('Tài liệu', [Environment]::GetFolderPath('MyDocuments')), @('Màn hình', [Environment]::GetFolderPath('Desktop')),
                       @('Tải về', (Join-Path $env:USERPROFILE 'Downloads')), @('OneDrive', $env:OneDrive))) {
        if ($x[1] -and (Test-Path -LiteralPath $x[1] -PathType Container)) { $noi += [pscustomobject]@{ ten = $x[0]; duong_dan = $x[1] } }
      }
      foreach ($dr in [IO.DriveInfo]::GetDrives()) { if ($dr.IsReady -and $dr.DriveType -in 'Fixed', 'Removable', 'Network') { $noi += [pscustomobject]@{ ten = "Ổ $($dr.Name.TrimEnd('\'))"; duong_dan = $dr.Name } } }
      $p = (Q $req 'p').Trim().Trim('"').Trim()
      if ($p -match '^file:///') { $p = [Uri]::UnescapeDataString($p.Substring(8)) -replace '/', '\' }
      if (-not $p) { SendJson $ctx @{ duong_dan = ''; cha = ''; noi = $noi; ds = @(); so_tep = 0 }; return }
      if (-not (Test-Path -LiteralPath $p -PathType Container)) { SendJson $ctx @{ loi = "Không tìm thấy thư mục: $p" } 400; return }
      $full = [IO.Path]::GetFullPath($p); if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
      $chaDir = [IO.Directory]::GetParent($full); $chaP = if ($chaDir) { $chaDir.FullName } else { '' }
      $an = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
      $con = @(Get-ChildItem -LiteralPath $full -Directory -ErrorAction SilentlyContinue | Where-Object { -not ($_.Attributes -band $an) } | Sort-Object Name)
      $reHS = '\.(pdf|docx?)$'
      $ds = @(foreach ($d in $con) {
        $hs = $false
        if ($con.Count -le 400) { try { foreach ($f in [IO.Directory]::EnumerateFiles($d.FullName)) { if ($f -match $reHS) { $hs = $true; break } } } catch { }; if (-not $hs) { try { $hs = [bool](Test-HoSoCoCauTruc $d.FullName) } catch { } } }
        [pscustomobject]@{ ten = $d.Name; duong_dan = $d.FullName; ho_so = $hs }
      })
      $soTep = 0; try { $soTep = @([IO.Directory]::EnumerateFiles($full) | Where-Object { $_ -match $reHS }).Count } catch { }
      SendJson $ctx @{ duong_dan = $full; cha = $chaP; noi = $noi; ds = $ds; so_tep = $soTep; trong_du_an = $full.StartsWith($GOC, [StringComparison]::OrdinalIgnoreCase) }; return }
    'GET /api/tim-folder' {
      $ten = Q $req 'ten'
      $ds = @(Get-ChildItem -LiteralPath $GOC -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $ten -and -not $_.FullName.StartsWith($Script:HE_THONG, [StringComparison]::OrdinalIgnoreCase) } | ForEach-Object { Rel $_.FullName })
      SendJson $ctx $ds; return }
    'GET /api/anh' {
      # Ảnh đầu trang trong cache của một văn bản: ?rel=<folder hồ sơ>&file=<tên ảnh>
      $full = Safe ((Q $req 'rel')); $f = Join-Path (CacheDir $full) ((Q $req 'file') -replace '[\\/]', '')
      if (Test-Path -LiteralPath $f) { Send $ctx ([IO.File]::ReadAllBytes($f)) $(if ($f -match '\.jpe?g$') { 'image/jpeg' } else { 'image/png' }) } else { Send $ctx ([byte[]]@()) 'text/plain' 404 }; return }
    'GET /api/tep' {
      $full = Safe ((Q $req 'rel')); $ext = [IO.Path]::GetExtension($full).ToLower()
      $type = switch ($ext) { '.png' { 'image/png' } '.txt' { 'text/plain; charset=utf-8' } '.md' { 'text/plain; charset=utf-8' } default { 'application/octet-stream' } }
      Send $ctx ([IO.File]::ReadAllBytes($full)) $type; return }
    'POST /api/kiem-tra' { SendJson $ctx @{ log = (Test-MoiTruong | Out-String) }; return }
    'POST /api/tat' { SendJson $ctx @{ ok = $true }; $Script:TAT = $true; return }
    default { SendJson $ctx @{ loi = "Không có đường dẫn $m $path" } 404 }
  }
}

# ---- Khởi động: tìm cổng trống, mở trình duyệt, vòng lặp phục vụ ----
$listener = $null
foreach ($pt in $Port..($Port + 10)) {
  try { $l = New-Object Net.HttpListener; $l.Prefixes.Add("http://localhost:$pt/"); $l.Start(); $listener = $l; $Port = $pt; break } catch { }
}
if (-not $listener) { Write-Host 'Không mở được cổng nào từ 8765 đến 8775.'; exit 1 }
$url = "http://localhost:$Port/"
Write-Host "Ứng dụng Hồ sơ ABTC đang chạy tại $url"
Write-Host 'Giữ cửa sổ này mở trong lúc dùng. Đóng cửa sổ (hoặc bấm "Tắt ứng dụng") để tắt.'
if (-not $KhongMoTrinhDuyet) { try { Start-Process 'msedge.exe' "--app=$url" -ErrorAction Stop } catch { Start-Process $url } }
$Script:TAT = $false
try {
  while ($listener.IsListening -and -not $Script:TAT) {
    $ctx = $listener.GetContext()
    $t0 = Get-Date
    try { Handle $ctx } catch { try { SendJson $ctx @{ loi = $_.Exception.Message } 500 } catch {} ; Write-Host "LỖI: $($_.Exception.Message)" }
    Write-Host ("{0:HH:mm:ss} {1} {2} ({3:N1}s)" -f $t0, $ctx.Request.HttpMethod, $ctx.Request.Url.PathAndQuery, ((Get-Date) - $t0).TotalSeconds)
  }
} finally { $listener.Stop(); $listener.Close(); Write-Host 'Đã tắt ứng dụng.' }
