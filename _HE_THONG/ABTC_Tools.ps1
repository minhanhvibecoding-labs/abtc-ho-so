# =====================================================================================
#  ABTC_Tools.ps1 – BỘ CÔNG CỤ HỢP NHẤT CHO CƠ SỞ DỮ LIỆU 2026 (Phòng Lãnh sự – Sở Ngoại vụ)
#  Yêu cầu: Windows 10/11, Windows PowerShell 5.1, Microsoft Excel + Word (dùng COM). Không cần Python.
#  File phải lưu dạng UTF-8 CÓ BOM (PowerShell 5.1 mới đọc đúng tiếng Việt).
#
#  CÁCH GỌI (từ thư mục gốc "CƠ SỞ DỮ LIỆU 2026"):
#     powershell -ExecutionPolicy Bypass -File "_HE_THONG\ABTC_Tools.ps1" <lệnh> [-Tham_so giá_trị ...]
#     hoặc trong PowerShell:  & "_HE_THONG\ABTC_Tools.ps1" <lệnh> ...
#     hoặc nạp hàm:           . "_HE_THONG\ABTC_Tools.ps1"   rồi gọi trực tiếp tên hàm
#  Xem danh sách lệnh:        & "_HE_THONG\ABTC_Tools.ps1" tro-giup
# =====================================================================================

$Script:HE_THONG = Split-Path -Parent $PSCommandPath          # ...\_HE_THONG
$Script:GOC      = Split-Path -Parent $Script:HE_THONG         # ...\CƠ SỞ DỮ LIỆU 2026
$Script:W_NS     = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Cấu hình: folder chứa file gốc, file -> sheet -> dòng mẫu (dòng dữ liệu đầu tiên)
$Script:CAU_HINH = @{
  ThuMucGoc = 'DỮ LIỆU GỐC'
  Files = @(
    @{ File = 'ABTC 2026.xlsx'; Sheets = [ordered]@{ 'ABTC 2026' = 5 } },
    @{ File = 'BẢO HỘ CÔNG DÂN.xlsx'; Sheets = [ordered]@{ 'BHCD VN ở nước ngoài' = 4; 'BHCD NN tại VIỆT NAM' = 6 } },
    @{ File = 'HNHTQT _ TTHC _ 2026.xlsx'; Sheets = [ordered]@{ 'TTHC HNHTQT' = 8 } },
    @{ File = 'NHẬP CẢNH 2026.xlsx'; Sheets = [ordered]@{ 'GIA HẠN VISA NĂM 2026' = 6; 'NHẬP CẢNH 2026' = 6 } },
    @{ File = 'ĐOÀN RA, VÀO, CỬ CÔNG TÁC, VIỆC RIÊNG 2026.xlsx'; Sheets = [ordered]@{
        'ĐOÀN RA CẤP SỞ NGÀNH 2026' = 5; 'ĐOÀN VÀO CẤP SỞ NGÀNH 2026' = 5; 'ĐOÀN VÀO CẤP THÀNH PHỐ 2026' = 5
        'CỬ CÁN BỘ ĐI CÔNG TÁC 2026' = 4; 'CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN' = 5 } }
  )
}

function P([string]$rel) { Join-Path $Script:GOC $rel }     # đường dẫn tính từ thư mục gốc
function New-TempFile([string]$ext) { Join-Path $env:TEMP ('abtc_' + [guid]::NewGuid().ToString('N') + $ext) }
function Open-Zip([string]$path) { Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::OpenRead($path) }
function Read-ZipXml($zip, [string]$name) {
  $e = $zip.Entries | Where-Object { $_.FullName -eq $name }; if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close(); [xml]$t
}
function New-Excel { $x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x }
function Close-Com($app) { try { $app.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
function Get-WordOptionsKey {
  # Khóa tùy chọn Word của bản Office đang cài (16.0, 15.0…): lấy bản cao nhất có mục Word trong HKCU, mặc định 16.0
  $vs = @(Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d+\.\d$' -and (Test-Path (Join-Path $_.PSPath 'Word')) } | Sort-Object { [double]$_.PSChildName } -Descending)
  $v = if ($vs.Count) { $vs[0].PSChildName } else { '16.0' }
  "HKCU:\Software\Microsoft\Office\$v\Word\Options"
}
function Test-WordPdf { (Get-ItemProperty (Get-WordOptionsKey) -Name DisableConvertPdfWarning -ErrorAction SilentlyContinue).DisableConvertPdfWarning -eq 1 }

# =====================================================================================
#  1. KIỂM TRA MÔI TRƯỜNG (chạy đầu tiên trên máy mới)
# =====================================================================================
function Test-MoiTruong {
  param([switch]$SuaLoi)   # -SuaLoi: tự bật tùy chọn Word cần thiết
  function Dong($ten, $dat, $ghichu) { "{0,-4} {1,-45} {2}" -f $(if ($dat) { 'OK' } else { 'LỖI' }), $ten, $ghichu }
  Dong 'Windows PowerShell 5.1' ($PSVersionTable.PSVersion.Major -eq 5) "Đang dùng $($PSVersionTable.PSVersion)"
  foreach ($app in 'Excel.Application', 'Word.Application') {
    try { $o = New-Object -ComObject $app; $v = $o.Version; Close-Com $o; Dong $app $true "phiên bản $v" } catch { Dong $app $false 'Chưa cài Office hoặc COM bị chặn' }
  }
  $k = Get-WordOptionsKey
  $val = (Get-ItemProperty $k -Name DisableConvertPdfWarning -ErrorAction SilentlyContinue).DisableConvertPdfWarning
  if ($val -ne 1 -and $SuaLoi) {
    if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
    New-ItemProperty -Path $k -Name DisableConvertPdfWarning -Value 1 -PropertyType DWord -Force | Out-Null; $val = 1
  }
  Dong 'Word mở PDF không hỏi (DisableConvertPdfWarning)' ($val -eq 1) $(if ($val -eq 1) { "($k)" } else { 'Chạy lại với -SuaLoi để bật (cần cho lệnh pdf-chu)' })
  try { Add-Type -AssemblyName System.Runtime.WindowsRuntime; $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]; Dong 'Windows.Data.Pdf (cắt đầu trang PDF)' $true '' } catch { Dong 'Windows.Data.Pdf (cắt đầu trang PDF)' $false $_.Exception.Message }
  $bom = [IO.File]::ReadAllBytes($PSCommandPath)[0..2] -join ','
  Dong 'ABTC_Tools.ps1 lưu UTF-8 có BOM' ($bom -eq '239,187,191') $(if ($bom -ne '239,187,191') { 'Lưu lại file dạng UTF-8 with BOM' } else { '' })
  foreach ($f in $Script:CAU_HINH.Files) { $p = P (Join-Path $Script:CAU_HINH.ThuMucGoc $f.File); Dong "Gốc: $($f.File)" (Test-Path -LiteralPath $p) '' }
  foreach ($f in $Script:CAU_HINH.Files) { $p = P ("KẾT QUẢ\MẪU\MẪU - " + $f.File); Dong "Mẫu: $($f.File)" (Test-Path -LiteralPath $p) '' }
  Dong 'Mẫu báo cáo tuần' (Test-Path -LiteralPath (P 'KẾT QUẢ\MẪU\MẪU - Báo cáo tuần.docx')) ''
  foreach ($f in 'CLAUDE.md', '_HE_THONG\ABTC_LOGIC.md', '_HE_THONG\TEMPLATE_MAP.md') { Dong $f (Test-Path -LiteralPath (P $f)) '' }
  "Thư mục gốc: $Script:GOC"
}

# =====================================================================================
#  2. ĐỌC FILE
# =====================================================================================
function Read-Xlsx {
  # Đọc toàn bộ .xlsx ra text (mỗi ô "B5=giá trị"). Đọc được cả khi file đang mở trong Excel (nếu Excel cho phép đọc chung).
  param([Parameter(Mandatory)][string]$Path, [string]$Out)
  if (-not $Out) { $Out = New-TempFile '.txt' }
  $tmp = New-TempFile '.xlsx'; Copy-Item -LiteralPath $Path $tmp -Force
  $zip = Open-Zip $tmp
  try {
    $ss = @(); $sst = Read-ZipXml $zip 'xl/sharedStrings.xml'
    if ($sst) { foreach ($si in $sst.sst.si) {
      if ($si.t) { if ($si.t -is [string]) { $ss += $si.t } else { $ss += $si.t.'#text' } }
      else { $ss += (($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '') } } }
    $wb = Read-ZipXml $zip 'xl/workbook.xml'; $rels = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
    $sb = New-Object System.Text.StringBuilder
    foreach ($s in $wb.workbook.sheets.sheet) {
      $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      $target = (($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target) -replace '^/xl/', '' -replace '^/', ''
      $sx = Read-ZipXml $zip ("xl/" + $target)
      [void]$sb.AppendLine("===== SHEET: " + $s.name + " =====")
      foreach ($row in $sx.worksheet.sheetData.row) {
        $cells = @()
        foreach ($c in $row.c) {
          if ($c.t -eq 's') { $v = $ss[[int]$c.v] } elseif ($c.t -eq 'inlineStr') { $v = $c.is.t } else { $v = $c.v }
          if ($null -ne $v -and "$v" -ne '') { $cells += ($c.r + "=" + (("$v") -replace "`r?`n", ' / ')) }
        }
        if ($cells.Count) { [void]$sb.AppendLine(($cells -join ' | ')) }
      }
    }
    [IO.File]::WriteAllText($Out, $sb.ToString(), [Text.Encoding]::UTF8)
  } finally { $zip.Dispose(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
  $Out
}

function Read-Docx {
  # Đọc .docx ra text, giữ bảng (dòng bảng "| ô | ô |", xuống dòng trong ô thành " / ").
  param([Parameter(Mandatory)][string]$Path, [string]$Out)
  if (-not $Out) { $Out = New-TempFile '.txt' }
  $tmp = New-TempFile '.docx'; Copy-Item -LiteralPath $Path $tmp -Force
  $zip = Open-Zip $tmp
  try { $xml = Read-ZipXml $zip 'word/document.xml' } finally { $zip.Dispose(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace('w', $Script:W_NS)
  $sb = New-Object System.Text.StringBuilder
  foreach ($node in $xml.SelectSingleNode('//w:body', $ns).ChildNodes) {
    if ($node.LocalName -eq 'p') {
      $t = ($node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
      if ($t.Trim()) { [void]$sb.AppendLine($t) }
    } elseif ($node.LocalName -eq 'tbl') {
      [void]$sb.AppendLine('[TABLE]')
      foreach ($tr in $node.SelectNodes('./w:tr', $ns)) {
        $cells = foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) { (($tc.SelectNodes('.//w:p', $ns) | ForEach-Object { ($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '' }) -join ' / ').Trim() }
        [void]$sb.AppendLine('| ' + ($cells -join ' | ') + ' |')
      }
      [void]$sb.AppendLine('[/TABLE]')
    }
  }
  [IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
  $Out
}

function Get-PdfText {
  # Trích CHỮ của nhiều PDF bằng Word (không tốn token ảnh). Lớp chữ bản scan thường là dự thảo -> số/ngày có thể trống.
  # -Names: tên tệp chữ đầu ra (không đuôi) cho từng PDF, cùng thứ tự với -Paths (hồ sơ có thư mục con: tránh trùng tên). Bỏ trống = tên PDF.
  param([Parameter(Mandatory)][string[]]$Paths, [Parameter(Mandatory)][string]$OutDir, [int]$TimeoutSec = 180, [string[]]$Names)
  New-Item -ItemType Directory -Force $OutDir | Out-Null
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
  $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
  $mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
  # Canh gác: Word treo (hộp thoại ẩn) quá thời gian -> tắt đúng phiên Word này
  $guard = Start-Job -ScriptBlock { param($ids, $s) Start-Sleep -Seconds $s; foreach ($i in $ids) { Stop-Process -Id $i -Force -ErrorAction SilentlyContinue } } -ArgumentList (, $mine), $TimeoutSec
  try {
    for ($pi = 0; $pi -lt $Paths.Count; $pi++) {
      $p = $Paths[$pi]
      $name = if ($Names -and $Names.Count -gt $pi -and $Names[$pi]) { $Names[$pi] } else { [IO.Path]::GetFileNameWithoutExtension($p) }; $tmp = New-TempFile '.pdf'
      Copy-Item -LiteralPath $p $tmp -Force
      try {
        $d = $w.Documents.Open($tmp, $false, $true, $false); $t = $d.Content.Text; $d.Close($false)
        $t = ($t -replace "`r", "`n" -replace "[ `t]+", ' ' -replace "(\n\s*){2,}", "`n").Trim()
        [IO.File]::WriteAllText((Join-Path $OutDir "$name.txt"), $t, (New-Object Text.UTF8Encoding($false)))
        "{0,6} ký tự  {1}" -f $t.Length, $name
      } catch { "LỖI  $name : $($_.Exception.Message)" }
      finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
  } finally { Close-Com $w; Stop-Job $guard -ErrorAction SilentlyContinue; Remove-Job $guard -Force -ErrorAction SilentlyContinue }
}

function Get-PdfHeader {
  # Cắt PHẦN ĐẦU trang 1 PDF (vùng "Số: .../ký hiệu", "ngày ... tháng ...") thành PNG nhỏ (~400–500 token khi đọc).
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Out, [double]$TopRatio = 0.30, [int]$Width = 900)
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
  $ext = [System.WindowsRuntimeSystemExtensions].GetMethods()
  $opM = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $acM = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
  $wait = { param($op, $t) $task = $opM.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
  $file = & $wait ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path -LiteralPath $Path).Path)) ([Windows.Storage.StorageFile])
  $doc = & $wait ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
  $page = $doc.GetPage(0)
  $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $opt.SourceRect = New-Object Windows.Foundation.Rect(0, 0, $page.Size.Width, ($page.Size.Height * $TopRatio))
  $opt.DestinationWidth = $Width
  $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  $task = $acM.Invoke($null, @($page.RenderToStreamAsync($ms, $opt))); $task.Wait(-1) | Out-Null
  $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0))
  $fs = [IO.File]::Create($Out); $net.CopyTo($fs); $fs.Close()
  "{0}  ({1} trang)" -f (Split-Path $Out -Leaf), $doc.PageCount
}

function Get-OcrText {
  # Nhận dạng chữ trên ảnh bằng OCR có sẵn của Windows (đọc tốt chữ số/ký hiệu; dấu tiếng Việt có thể mất).
  param([Parameter(Mandatory)][string]$Path)
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
  $opM = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $wait = { param($op, $t) $task = $opM.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
  $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages(); if (-not $eng) { return '' }
  $f = & $wait ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path -LiteralPath $Path).Path)) ([Windows.Storage.StorageFile])
  $s = & $wait ($f.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  try {
    $dec = & $wait ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($s)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = & $wait ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $r = & $wait ($eng.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    (($r.Lines | ForEach-Object { $_.Text }) -join "`n")
  } finally { $s.Dispose() }
}

# =====================================================================================
#  3. CHUẨN BỊ FOLDER HỒ SƠ (tiết kiệm token: chỉ xử lý file mới/đổi; chữ PDF qua Word; ảnh cắt đầu trang)
# =====================================================================================
function Invoke-ChuanBiFolder {
  # -Path: 1 folder hồ sơ, hoặc folder cha chứa nhiều folder hồ sơ. Kết quả ở _HE_THONG\cache\<tên folder>\ :
  #   *.txt (chữ của PDF/DOCX), *_dau.png (đầu trang 1 khi tên file không có số), manifest.json, tom_tat.md
  param([Parameter(Mandatory)][string]$Path, [string]$CacheDir = (Join-Path $Script:HE_THONG 'cache'), [switch]$LamLai)
  $root = (Resolve-Path -LiteralPath $Path).Path
  # 1 hồ sơ (phẳng hoặc có cấu trúc thư mục loại văn bản) → chính nó; folder tổng (Tháng/Doanh nghiệp…) → mọi hồ sơ tìm được bên trong
  $dirs = if (Test-LaHoSo $root) { @(Get-Item -LiteralPath $root) } else { @(Find-HoSo $root) }
  foreach ($d in $dirs) {
    $cache = Join-Path $CacheDir $d.Name; New-Item -ItemType Directory -Force $cache | Out-Null
    $manPath = Join-Path $cache 'manifest.json'
    $man = @{}; if ((Test-Path $manPath) -and -not $LamLai) { (Get-Content $manPath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $man[$_.Name] = $_.Value } }
    $tep = @(Get-TepHoSo $d.FullName | Where-Object { $_.File.Extension -ne '.doc' })   # .doc cũ: không đọc được chữ
    $moi = @(); $giu = 0
    foreach ($t in $tep) {
      $f = $t.File; $key = $t.Key; $sig = "$($f.Length)|$($f.LastWriteTimeUtc.Ticks)"
      if ($man.ContainsKey($key) -and $man[$key].sig -eq $sig) { $giu++; continue }
      $moi += $t
      $man[$key] = [pscustomobject]@{ sig = $sig; hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash; txt = ''; anh = ''; so = ''; ngay = '' }
    }
    # Chữ: DOCX đọc trực tiếp, PDF gom lại cho 1 phiên Word (thời gian chờ theo số PDF), ảnh JPG/PNG (văn bản chụp) đọc bằng OCR
    foreach ($t in $moi | Where-Object { $_.File.Extension -eq '.docx' }) { $o = Join-Path $cache ($t.CacheBase + '.txt'); [void](Read-Docx -Path $t.File.FullName -Out $o); $man[$t.Key].txt = Split-Path $o -Leaf }
    $pdfMoi = @($moi | Where-Object { $_.File.Extension -eq '.pdf' })
    if ($pdfMoi.Count) {
      Get-PdfText -Paths @($pdfMoi | ForEach-Object { $_.File.FullName }) -Names @($pdfMoi | ForEach-Object { $_.CacheBase }) -OutDir $cache -TimeoutSec ([Math]::Max(180, 45 * $pdfMoi.Count)) | Out-Null
      foreach ($t in $pdfMoi) { $man[$t.Key].txt = $t.CacheBase + '.txt' }
    }
    foreach ($t in $moi | Where-Object { $_.File.Extension -match '^\.(jpe?g|png)$' }) {
      $anh = Join-Path $cache ($t.CacheBase + '_dau' + $t.File.Extension.ToLower())
      try {
        Copy-Item -LiteralPath $t.File.FullName -Destination $anh -Force; $man[$t.Key].anh = Split-Path $anh -Leaf
        $ocr = Get-OcrText -Path $anh; $u8 = New-Object Text.UTF8Encoding($false)
        [IO.File]::WriteAllText((Join-Path $cache ($t.CacheBase + '_dau.ocr.txt')), $ocr, $u8)
        [IO.File]::WriteAllText((Join-Path $cache ($t.CacheBase + '.txt')), $ocr, $u8); $man[$t.Key].txt = $t.CacheBase + '.txt'
      } catch { $man[$t.Key].anh = 'LỖI: ' + $_.Exception.Message }
    }
    # Số/ngày: từ tên file; nếu không có thì cắt đầu trang 1 (PDF)
    foreach ($t in $moi) {
      $f = $t.File; $sn = Get-SoNgayTuTen $f.Name
      if ($sn) { $man[$t.Key].so = $sn.so; $man[$t.Key].ngay = $sn.ngay }
      elseif ($f.Extension -eq '.pdf') {
        $png = Join-Path $cache ($t.CacheBase + '_dau.png')
        try {
          [void](Get-PdfHeader -Path $f.FullName -Out $png); $man[$t.Key].anh = Split-Path $png -Leaf
          try { $ocr = Get-OcrText -Path $png; [IO.File]::WriteAllText(($png -replace '\.png$', '.ocr.txt'), $ocr, (New-Object Text.UTF8Encoding($false))) } catch {}   # OCR số/ngày
        } catch { $man[$t.Key].anh = 'LỖI: ' + $_.Exception.Message }
      }
    }
    # Bỏ file đã bị xóa khỏi folder
    $conKey = @{}; foreach ($t in $tep) { $conKey[$t.Key] = 1 }
    foreach ($k in @($man.Keys)) { if (-not $conKey.ContainsKey($k)) { $man.Remove($k) } }
    [IO.File]::WriteAllText($manPath, ($man | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
    # Tóm tắt cho agent / người đọc
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# $($d.Name)"); [void]$sb.AppendLine("Nguồn: $($d.FullName)"); [void]$sb.AppendLine('')
    [void]$sb.AppendLine('| File | Loại (đoán theo tên) | Số (tên file) | Ngày (tên file) | File chữ | Ảnh đầu trang | Ghi chú |'); [void]$sb.AppendLine('|---|---|---|---|---|---|---|')
    foreach ($t in $tep) {
      $f = $t.File; $m = $man[$t.Key]; $n = $t.Key
      $loai = Get-LoaiTheoTen $f.Name $f.Extension; if ($t.LoaiThuMuc -and $loai -eq 'KHAC') { $loai = $t.LoaiThuMuc + ' (theo thư mục)' }
      if ($loai -eq 'DN_HO_SO_KEM') { $loai += ' (chỉ cần số/ngày)' }
      $trung = @($man.GetEnumerator() | Where-Object { $_.Key -ne $n -and $_.Value.hash -eq $m.hash } | ForEach-Object { $_.Key })
      $ghi = @(); if ($trung.Count) { $ghi += 'TRÙNG với ' + ($trung -join ', ') }; if ($n -match 'bản ký') { $ghi += 'bản ký chưa vào số' }
      [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} |' -f $n, $loai, $m.so, $m.ngay, $m.txt, $m.anh, ($ghi -join '; ')))
    }
    [IO.File]::WriteAllText((Join-Path $cache 'tom_tat.md'), $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
    "{0}: {1} file mới/đổi, {2} file giữ nguyên  -> {3}" -f $d.Name, $moi.Count, $giu, $cache
  }
}

# =====================================================================================
#  4. TEMPLATE VÀ DỮ LIỆU
# =====================================================================================
function New-MauExcel {
  # Tạo template Excel rỗng từ file gốc: giữ tiêu đề/định dạng, xóa dữ liệu, chừa 1 dòng mẫu trống.
  param([string]$OutDir = (P 'KẾT QUẢ\MẪU'))
  New-Item -ItemType Directory -Force $OutDir | Out-Null
  $xl = New-Excel
  try {
    foreach ($c in $Script:CAU_HINH.Files) {
      $dst = Join-Path $OutDir ('MẪU - ' + $c.File)
      Copy-Item -LiteralPath (P (Join-Path $Script:CAU_HINH.ThuMucGoc $c.File)) $dst -Force -ErrorAction Stop
      $wb = $xl.Workbooks.Open($dst)
      foreach ($ws in $wb.Worksheets) {
        if (-not $c.Sheets.Contains($ws.Name)) { "$($c.File) | $($ws.Name) | giữ nguyên"; continue }
        $start = [int]$c.Sheets[$ws.Name]; $ur = $ws.UsedRange; $last = $ur.Row + $ur.Rows.Count - 1
        if ($last -gt $start) { [void]$ws.Range("$($start + 1):$last").EntireRow.Delete() }
        $r = $ws.Rows($start); [void]$r.ClearContents(); [void]$r.ClearComments(); $r.Interior.ColorIndex = -4142
        if ($ws.Name -eq 'CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN' -and $ws.Range('A1').Text -eq 'Y') { [void]$ws.Range('A1').ClearContents() }
        "$($c.File) | $($ws.Name) | dòng mẫu $start"
      }
      $wb.Worksheets.Item(1).Activate(); $wb.Save(); $wb.Close($false)
    }
  } finally { Close-Com $xl }
}

function New-MauWord {
  # Tạo template báo cáo tuần từ 1 báo cáo tuần gốc (.docx): giữ khung + dòng mục, 1 dòng mẫu trống dưới mỗi mục,
  # tiêu đề thành "{{TU_NGAY}} - {{DEN_NGAY}}", thay tiêu ngữ cơ quan bằng Org1/Org2.
  param([Parameter(Mandatory)][string]$Source, [string]$Out = (P 'KẾT QUẢ\MẪU\MẪU - Báo cáo tuần.docx'),
        [string]$Org1 = 'SỞ NGOẠI VỤ', [string]$Org2 = 'PHÒNG LÃNH SỰ')
  Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem
  $W = $Script:W_NS
  Copy-Item -LiteralPath $Source $Out -Force
  $zip = [System.IO.Compression.ZipFile]::Open($Out, 'Update')
  try {
    $entry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xml = New-Object System.Xml.XmlDocument; $xml.PreserveWhitespace = $true; $xml.LoadXml($sr.ReadToEnd()); $sr.Close()
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace('w', $W)
    function CellText($tc) { (($tc.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim() }
    function ClearCell($tc) {
      $ps = @($tc.SelectNodes('./w:p', $ns))
      for ($i = 1; $i -lt $ps.Count; $i++) { [void]$tc.RemoveChild($ps[$i]) }
      if ($ps.Count -gt 0) { foreach ($r in @($ps[0].SelectNodes('./w:r | ./w:hyperlink', $ns))) { [void]$ps[0].RemoveChild($r) } }
    }
    # Tiêu đề
    foreach ($p in $xml.SelectNodes('//w:body/w:p', $ns)) {
      $txt = ($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
      if ($txt -match 'tuần từ ngày') {
        $ts = @($p.SelectNodes('.//w:t', $ns)); $ts[0].InnerText = ($txt -replace 'tuần từ ngày.*$', 'tuần từ ngày {{TU_NGAY}} - {{DEN_NGAY}}')
        [void]$ts[0].SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve'); for ($i = 1; $i -lt $ts.Count; $i++) { $ts[$i].InnerText = '' }
      }
    }
    # Tiêu ngữ cơ quan
    $orgCell = $xml.SelectSingleNode('//w:body/w:tbl[1]/w:tr[1]/w:tc[1]', $ns)
    $orgParas = @($orgCell.SelectNodes('./w:p', $ns) | Where-Object { (($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim() })
    $newOrg = @($Org1, $Org2)
    for ($k = 0; $k -lt [Math]::Min(2, $orgParas.Count); $k++) { $ts = @($orgParas[$k].SelectNodes('.//w:t', $ns)); $ts[0].InnerText = $newOrg[$k]; for ($j = 1; $j -lt $ts.Count; $j++) { $ts[$j].InnerText = '' } }
    # Bảng chính: giữ dòng mục, 1 dòng mẫu trống, xóa phần còn lại
    $tbl = @($xml.SelectNodes('//w:body/w:tbl', $ns))[1]; $rows = @($tbl.SelectNodes('./w:tr', $ns)); $kept = $false
    for ($i = 2; $i -lt $rows.Count; $i++) {
      $cells = @($rows[$i].SelectNodes('./w:tc', $ns)); $c1 = CellText $cells[0]; $c2 = if ($cells.Count -gt 1) { CellText $cells[1] } else { '' }
      if ($c1 -or $c2) { for ($k = 2; $k -lt $cells.Count; $k++) { ClearCell $cells[$k] }; $kept = $false }
      elseif (-not $kept) { foreach ($tc in $cells) { ClearCell $tc }; $kept = $true }
      else { [void]$tbl.RemoveChild($rows[$i]) }
    }
    # Dòng mẫu chuẩn 8 ô; thay dòng lỗi số ô
    $rows = @($tbl.SelectNodes('./w:tr', $ns)); $std = $null
    for ($i = 2; $i -lt $rows.Count; $i++) { $cs = @($rows[$i].SelectNodes('./w:tc', $ns)); if ($cs.Count -eq 8 -and -not (CellText $cs[0]) -and -not (CellText $cs[1])) { $std = $rows[$i]; break } }
    for ($i = 2; $i -lt $rows.Count; $i++) { $cs = @($rows[$i].SelectNodes('./w:tc', $ns)); $head = (CellText $cs[0]) -or ($cs.Count -gt 1 -and (CellText $cs[1])); if (-not $head -and $cs.Count -ne 8) { [void]$tbl.ReplaceChild($std.CloneNode($true), $rows[$i]) } }
    # Mục không có mục con, chưa có dòng mẫu -> thêm
    $rows = @($tbl.SelectNodes('./w:tr', $ns))
    for ($i = 2; $i -lt $rows.Count; $i++) {
      $cs = @($rows[$i].SelectNodes('./w:tc', $ns)); $c1 = CellText $cs[0]; $c2 = if ($cs.Count -gt 1) { CellText $cs[1] } else { '' }
      if (-not ($c1 -or $c2)) { continue }; if ($c1 -match '^[IVX]+$' -and $c1 -ne 'II') { continue }
      if ($i + 1 -lt $rows.Count) {
        $ncs = @($rows[$i + 1].SelectNodes('./w:tc', $ns)); $n1 = CellText $ncs[0]; $n2 = if ($ncs.Count -gt 1) { CellText $ncs[1] } else { '' }
        if (-not ($n1 -or $n2)) { continue }; if ($c1 -and ($n1 -like "$c1.*" -or (-not $n1 -and $n2))) { continue }
      }
      [void]$tbl.InsertAfter($std.CloneNode($true), $rows[$i])
    }
    # Cắt cột "ma" thừa trong khung bảng
    $maxCols = 0
    foreach ($tr in $tbl.SelectNodes('./w:tr', $ns)) {
      $n = 0; foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) { $gs = $tc.SelectSingleNode('./w:tcPr/w:gridSpan', $ns); $n += $(if ($gs) { [int]$gs.GetAttribute('val', $W) } else { 1 }) }
      if ($n -gt $maxCols) { $maxCols = $n }
      foreach ($tag in 'gridAfter', 'wAfter', 'gridBefore', 'wBefore') { $el = $tr.SelectSingleNode("./w:trPr/w:$tag", $ns); if ($el) { [void]$el.ParentNode.RemoveChild($el) } }
    }
    $grid = $tbl.SelectSingleNode('./w:tblGrid', $ns); $cols = @($grid.SelectNodes('./w:gridCol', $ns))
    for ($k = $maxCols; $k -lt $cols.Count; $k++) { [void]$grid.RemoveChild($cols[$k]) }
    $sum = 0; foreach ($gc in $grid.SelectNodes('./w:gridCol', $ns)) { $sum += [int]$gc.GetAttribute('w', $W) }
    $tblW = $tbl.SelectSingleNode('./w:tblPr/w:tblW', $ns); if ($tblW) { [void]$tblW.SetAttribute('w', $W, "$sum"); [void]$tblW.SetAttribute('type', $W, 'dxa') }
    $entry.Delete(); $ne = $zip.CreateEntry('word/document.xml')
    $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false))); $xml.Save($sw); $sw.Close()
    "Đã tạo mẫu báo cáo tuần: $Out ($maxCols cột)"
  } finally { $zip.Dispose() }
}

function Move-DuLieuCu {
  # Chép dữ liệu từ file gốc (DỮ LIỆU GỐC) vào file kết quả dựng từ template, giữ đúng số dòng như file gốc.
  param([switch]$Force, [string[]]$Only, [string]$FallbackDir, [string]$OutDir = (P 'KẾT QUẢ\CƠ SỞ DỮ LIỆU'))
  $tplDir = P 'KẾT QUẢ\MẪU'; New-Item -ItemType Directory -Force $OutDir | Out-Null
  $xl = New-Excel
  try {
    foreach ($f in $Script:CAU_HINH.Files) {
      if ($Only -and ($Only -notcontains $f.File)) { continue }
      $dst = Join-Path $OutDir $f.File
      if ((Test-Path -LiteralPath $dst) -and -not $Force) { "BỎ QUA $($f.File): đã có (dùng -Force để làm lại)"; continue }
      $src = Join-Path $env:TEMP ('NGUON_' + [guid]::NewGuid().ToString('N') + '_' + $f.File)   # Excel không mở được 2 file trùng tên
      try { Copy-Item -LiteralPath (P (Join-Path $Script:CAU_HINH.ThuMucGoc $f.File)) $src -Force -ErrorAction Stop }
      catch { $fb = if ($FallbackDir) { Join-Path $FallbackDir $f.File }; if ($fb -and (Test-Path -LiteralPath $fb)) { Copy-Item -LiteralPath $fb $src -Force } else { "LỖI $($f.File): file gốc đang bị khóa – đóng Excel rồi chạy lại với -Only"; continue } }
      $wbS = $null; $wbD = $null
      try {
        Copy-Item -LiteralPath (Join-Path $tplDir ('MẪU - ' + $f.File)) $dst -Force -ErrorAction Stop
        $wbS = $xl.Workbooks.Open($src, 0, $true); $wbD = $xl.Workbooks.Open($dst)
        foreach ($sheet in $f.Sheets.Keys) {
          $wsS = $wbS.Worksheets.Item($sheet); $wsD = $wbD.Worksheets.Item($sheet); $start = [int]$f.Sheets[$sheet]
          $lc = $wsS.Cells.Find('*', [Type]::Missing, -4123, 2, 1, 2); $last = if ($lc) { $lc.Row } else { 0 }
          if ($last -lt $start) { "$($f.File) | $sheet | không có dữ liệu"; continue }
          [void]$wsS.Range("$($start):$last").EntireRow.Copy($wsD.Range("A$start"))
          "$($f.File) | $sheet | dòng $start-$last"
        }
        $wbD.Worksheets.Item(1).Activate(); $wbD.Save()
      } catch { "LỖI $($f.File): $($_.Exception.Message)" }
      finally { if ($wbD) { $wbD.Close($false) }; if ($wbS) { $wbS.Close($false) }; Remove-Item $src -Force -ErrorAction SilentlyContinue }
    }
  } finally { Close-Com $xl }
}

function Set-ThayDoi {
  # Ghi thay đổi ĐÃ DUYỆT (file JSON) vào Excel. Tự sao lưu; chỉ ghi khi giá trị hiện tại khớp "cu" lúc duyệt.
  # JSON: [ { "sheet": "ABTC 2026", "o": "L112", "cu": "", "moi": "123/TTr-NV ngày 21/09/2026" } ]  ("\n" = xuống dòng trong ô)
  param([Parameter(Mandatory)][string]$File, [Parameter(Mandatory)][string]$Changes, [string]$BackupDir = (Join-Path $Script:HE_THONG 'backup'))
  $list = @((Get-Content -LiteralPath $Changes -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ })
  New-Item -ItemType Directory -Force $BackupDir | Out-Null
  $bk = Join-Path $BackupDir ('{0}_{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($File), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($File))
  Copy-Item -LiteralPath $File $bk -ErrorAction Stop; "Đã sao lưu: $bk"
  $norm = { param($s) if ($null -eq $s) { '' } else { ([string]$s -replace "`r", '' -replace "`n", ' / ' -replace '\s+', ' ').Trim() } }   # ô nhiều dòng so khớp như bản đọc text
  $xl = New-Excel; $wb = $null
  try {
    $wb = $xl.Workbooks.Open($File); if ($wb.ReadOnly) { throw 'File đang mở ở nơi khác. Đóng file trong Excel rồi chạy lại.' }
    $bad = foreach ($c in $list) { $cur = & $norm ($wb.Worksheets.Item($c.sheet).Range($c.o).Value2); if ($cur -ne (& $norm $c.cu)) { "$($c.sheet)!$($c.o): hiện là [$cur], lúc duyệt là [$($c.cu)]" } }
    if ($bad) { throw ("Giá trị đã thay đổi so với lúc duyệt, KHÔNG ghi:`n  " + ($bad -join "`n  ")) }
    foreach ($c in $list) {
      $r = $wb.Worksheets.Item($c.sheet).Range($c.o); $r.NumberFormat = '@'
      $v = ([string]$c.moi) -replace '\\n', "`n"; $r.Value2 = $v; if ($v.Contains("`n") -or $v.Length -gt 40) { $r.WrapText = $true }
      "Đã ghi $($c.sheet)!$($c.o)"
    }
    $wb.Save()
  } finally { if ($wb) { $wb.Close($false) }; Close-Com $xl }
}

# =====================================================================================
#  5. THỐNG KÊ ABTC
# =====================================================================================
function New-ThongKe {
  # Dựng file thống kê từ JSON gộp kết quả đọc folder (mảng hồ sơ, khuôn ở README).
  # Sheet: ABTC 2026 | Thống kê (công thức) | Dữ liệu (mốc ngày) | Văn bản | Cảnh báo
  param([Parameter(Mandatory)][string]$Data, [Parameter(Mandatory)][string]$Out, [string]$NgayLap = (Get-Date -Format 'dd/MM/yyyy'),
        [string]$Nguon = '', [string]$Template = (P 'KẾT QUẢ\MẪU\MẪU - ABTC 2026 (khung mới).xlsx'))
  $recs = @((Get-Content -LiteralPath $Data -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ })   # PS 5.1: tách mảng
  $n = $recs.Count
  $TEN_TT = [ordered]@{ 'S0_MOI' = 'Mới tiếp nhận'; 'S1_CHO_BO_SUNG' = 'Chờ DN bổ sung'; 'S2_DA_BO_SUNG' = 'DN đã bổ sung'; 'S3_XIN_Y_KIEN_LN' = 'Đang xin ý kiến LN'
    'S4_DON_DOC_LN' = 'Đã đôn đốc LN'; 'S5_DU_Y_KIEN_LN' = 'Đủ ý kiến LN, chờ trình'; 'S5b_DANG_TRINH_SO' = 'Đang trình BGĐ Sở'; 'S6_DA_TRINH_UB' = 'Đã trình UB, chờ kết quả'
    'S7_UB_CHO_PHEP' = 'UB cho phép'; 'S8_UB_KHONG_CHO_PHEP' = 'UB không cho phép'; 'S9_TRA_HO_SO' = 'Trả hồ sơ'; 'S10_DN_RUT' = 'DN rút hồ sơ' }
  function Txt($v) { if ($null -eq $v) { '' } else { ([string]$v) -replace '\\n', "`n" } }
  function ToDate([string]$s) { if ($s -match '(\d{1,2})/(\d{1,2})/(\d{4})') { try { return ([datetime]::new([int]$Matches[3], [int]$Matches[2], [int]$Matches[1])).ToOADate() } catch {} }; $null }
  function FirstEvent($r, [string]$pat) { foreach ($e in @($r.su_kien)) { if ($e.su_kien -match $pat) { return $e.ngay } }; '' }
  function MaTT($r) { $m = [string]$r.trang_thai; foreach ($k in $TEN_TT.Keys) { if ($m -like "$k*") { return $k } }; $m }
  function Put($ws, [int]$row, [int]$col, $rows) {
    $h = $rows.Count; if ($h -eq 0) { return $null }
    $wd = ($rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $arr = New-Object 'object[,]' $h, $wd
    for ($i = 0; $i -lt $h; $i++) { for ($j = 0; $j -lt $rows[$i].Count; $j++) {
      $v = $rows[$i][$j]
      if ($v -is [array]) { $v = ($v | ForEach-Object { [string]$_ }) -join '; ' } elseif ($null -ne $v -and -not ($v -is [string] -or $v -is [ValueType])) { $v = [string]$v }
      $arr[$i, $j] = $v } }
    $rng = $ws.Range($ws.Cells.Item($row, $col), $ws.Cells.Item($row + $h - 1, $col + $wd - 1))
    # Excel đôi khi từ chối ghi cả khối (lỗi thoáng qua) -> thử lại 1 lần, rồi ghi từng ô (ô lỗi thì ghi dạng chữ)
    $done = $false
    foreach ($try in 1..2) { try { $rng.Value2 = $arr; $done = $true; break } catch { Start-Sleep -Milliseconds 700 } }
    if (-not $done) { for ($i = 0; $i -lt $h; $i++) { for ($j = 0; $j -lt $wd; $j++) { $v = $arr[$i, $j]; if ($null -eq $v) { continue }
      $cell = $ws.Cells.Item($row + $i, $col + $j); try { $cell.Value2 = $v } catch { try { $cell.Value2 = [string]$v } catch { Start-Sleep -Milliseconds 300; $cell.Value2 = [string]$v } } } } }
    return , $rng
  }
  function Head($rg) { $rg.Font.Bold = $true; $rg.Interior.Color = 0xEBD9C6; $rg.HorizontalAlignment = -4108; $rg.VerticalAlignment = -4108; $rg.WrapText = $true }
  function Grid($rg) { $rg.Borders.LineStyle = 1; $rg.Borders.Weight = 2; $rg.VerticalAlignment = -4160 }
  function Title($ws, [string]$t, [string]$sub, [int]$cols) {
    $ws.Cells.Item(1, 1).Value2 = $t; $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cols)).Merge(); $ws.Cells.Item(1, 1).Font.Bold = $true; $ws.Cells.Item(1, 1).Font.Size = 14; $ws.Cells.Item(1, 1).HorizontalAlignment = -4108
    $ws.Cells.Item(2, 1).Value2 = $sub; $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item(2, $cols)).Merge(); $ws.Cells.Item(2, 1).Font.Italic = $true; $ws.Cells.Item(2, 1).HorizontalAlignment = -4108
  }
  function NewList { , (New-Object 'System.Collections.Generic.List[object[]]') }
  function TachSoNgay([string]$s) {
    # "số 12345/VP-LSLT ngày 10/8/2026; 12345/VP-LSLT ngày 2/9/2026 (đôn đốc)" → so = "12345/VP-LSLT⏎12345/VP-LSLT", ngay = "10/08/2026⏎02/09/2026"
    # Bỏ "(nhận ngày …)", ghi chú trong ngoặc, "Lần 1:", "1.", "Văn bản số"; ô chỉ có ngày thì số để trống (không nhét ngày vào cột số)
    $so = @(); $ng = @()
    foreach ($p in ($s -split "`r?`n|;")) {
      $p = $p -replace '\((?:đã\s+)?nhận[^)]*\)', ''
      $p = $p -replace '^\s*(?:lần\s*\d+\s*[:\-–]|\d+\s*[\.\)])\s*', '' -replace '^\s*(?:(?:văn bản|công văn|cv|tờ trình)\s+)?số\s*:?\s*', ''
      $p = $p.Trim(' ', ',', '.', ';', '/')
      if (-not $p) { continue }
      $d = ''
      if ($p -match 'ngày\s+(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})') { $d = '{0:00}/{1:00}/{2}' -f [int]$Matches[1], [int]$Matches[2], $Matches[3] }
      elseif ($p -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') { $d = '{0:00}/{1:00}/{2}' -f [int]$Matches[1], [int]$Matches[2], $Matches[3]; $p = '' }
      $sv = (($p -replace '\s*ngày\s+\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4}[\s\S]*$', '') -replace '\s*\([^)]*\)\s*$', '').Trim(' ', ',', ';', '.')
      if (-not $sv -and -not $d) { continue }
      $so += $sv; $ng += $d
    }
    # Giữ nguyên thứ tự dòng để số và ngày khớp nhau; vế nào trống hết thì để ô trống
    @{ so = $(if (($so -join '').Trim()) { $so -join "`n" } else { '' }); ngay = $(if (($ng -join '').Trim()) { $ng -join "`n" } else { '' }) }
  }
  function TenCaNhan($x) {
    # Cột TÊN CÁ NHÂN: chỉ họ tên, mỗi người một dòng (bỏ "Ông/Bà" và chức vụ)
    $ds = @(@($x.ds_ca_nhan) | Where-Object { $_ -and $_.ho_ten } | ForEach-Object { [string]$_.ho_ten })
    if (-not $ds.Count) { $ds = @((Txt $x.cot_excel.G) -split "`r?`n|;" | ForEach-Object { ($_ -split ',')[0] }) }
    (@($ds | ForEach-Object { ($_ -replace '^\s*(?:ông|bà|Ông|Bà)\s+', '' -replace '\s+', ' ').Trim() } | Where-Object { $_ }) -join "`n")
  }

  New-Item -ItemType Directory -Force (Split-Path $Out -Parent) | Out-Null
  Copy-Item -LiteralPath $Template $Out -Force
  $xl = New-Excel; $wb = $null
  try {
    $wb = $xl.Workbooks.Open($Out); $font = $wb.Worksheets.Item(1).Range('B2').Font.Name
    # --- ABTC 2026 (khung mới 23 cột, tiêu đề dòng 2, dòng mẫu 3) ---
    # Mỗi loại văn bản có 2 cột: SỐ văn bản (chỉ số/ký hiệu) và NGÀY (dd/mm/yyyy). Nhiều văn bản → mỗi văn bản một dòng trong ô, số và ngày cùng thứ tự.
    $ws = $wb.Worksheets.Item('ABTC 2026'); $D0 = 3
    if ($n -gt 1) { [void]$ws.Rows($D0).Copy(); [void]$ws.Rows("$($D0 + 1):$($D0 + $n - 1)").Insert() }
    $ws.Range("B$($D0):W$($D0 + $n - 1)").NumberFormat = '@'
    $RA = NewList
    for ($i = 0; $i -lt $n; $i++) {
      $x = $recs[$i]; $c = $x.cot_excel
      $vbDn = TachSoNgay (Txt $c.B); if (-not $vbDn.so -and $x.so_vb_dn) { $vbDn = @{ so = [string]$x.so_vb_dn; ngay = (Format-Ngay ([string]$x.ngay_vb_dn)) } }
      $tra = TachSoNgay (Txt $c.H); $bs = TachSoNgay (Txt $c.I); $xin = TachSoNgay (Txt $c.J); $dd = TachSoNgay (Txt $c.K); $tr = TachSoNgay (Txt $c.L); $kq = TachSoNgay (Txt $c.M)
      $RA.Add([object[]]@(($i + 1), $vbDn.so, $vbDn.ngay, (Txt $c.C), (Txt $c.D), (Txt $c.E), (Txt $c.F), (TenCaNhan $x),
        $tra.so, $tra.ngay, $bs.so, $bs.ngay, $xin.so, $xin.ngay, $dd.so, $dd.ngay, $tr.so, $tr.ngay, $kq.so, $kq.ngay, (Txt $c.N), (Txt $c.O), (Txt $c.P)))
    }
    $rg = Put $ws $D0 1 $RA; $rg.WrapText = $true; $rg.VerticalAlignment = -4160
    # --- Dữ liệu ---
    $wd = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $wd.Name = 'Dữ liệu'
    Title $wd 'DỮ LIỆU MỐC THỜI GIAN HỒ SƠ ABTC' "Ngày lập: $NgayLap – $Nguon" 24
    $RD = NewList
    $RD.Add([object[]]@('STT', 'Doanh nghiệp', 'Viết tắt', 'Cán bộ', 'Số doanh nhân', 'Mã trạng thái', 'Trạng thái', 'Ngày nhận HS', 'Ngày xin LN', 'Ngày đủ LN', 'Ngày trình UB', 'Ngày KQ UB', 'Số LN đã trả lời (/4)', 'Số lần đôn đốc', 'Số cảnh báo', 'Xin LN → đủ LN (ngày)', 'Đủ LN → KQ UB (ngày)', 'Nhận HS → KQ UB (ngày)', 'Tình trạng', 'LN Công an', 'LN Thuế', 'LN BHXH', 'LN Hải quan', 'Dòng trong Excel tổng hợp'))
    for ($i = 0; $i -lt $n; $i++) {
      $x = $recs[$i]; $ma = MaTT $x
      $lnD = @(@('CA', 'THUE', 'BHXH', 'HQ') | ForEach-Object { ToDate ([string]$x.ln.$_.ngay) })
      $du = ($lnD | Where-Object { $_ } | Measure-Object -Maximum).Maximum
      $kq = ToDate (FirstEvent $x 'KQ_UB')
      if ($kq -and $du -and $du -gt $kq) { $pt = @($x.su_kien | Where-Object { $_.trang_thai_sau -like 'S5b*' } | ForEach-Object { ToDate $_.ngay } | Where-Object { $_ }); $du = if ($pt.Count) { $pt[-1] } else { $kq } }
      $RD.Add([object[]]@(($i + 1), $x.doanh_nghiep, $x.ten_viet_tat, $x.can_bo, [int]$x.so_luong, $ma, $TEN_TT[$ma], (ToDate $x.ngay_nhan), (ToDate (FirstEvent $x 'XIN_Y_KIEN_LN')), $du, (ToDate (FirstEvent $x 'TRINH_UB')), $kq,
        @($lnD | Where-Object { $_ }).Count, @($x.don_doc).Count, @($x.canh_bao).Count, $null, $null, $null, $null, $lnD[0], $lnD[1], $lnD[2], $lnD[3], $x.dong_excel_hien_co))
    }
    $rg = Put $wd 3 1 $RD; Grid $rg; Head $wd.Range('A3:X3'); $last = 3 + $n
    foreach ($cl in 'H', 'I', 'J', 'K', 'L', 'T', 'U', 'V', 'W') { $wd.Range("$($cl)4:$cl$last").NumberFormat = 'dd/mm/yyyy' }
    $wd.Range("P4:P$last").FormulaR1C1 = '=IF(AND(RC9<>"",RC10<>""),RC10-RC9,"")'
    $wd.Range("Q4:Q$last").FormulaR1C1 = '=IF(AND(RC10<>"",RC12<>""),RC12-RC10,"")'
    $wd.Range("R4:R$last").FormulaR1C1 = '=IF(AND(RC8<>"",RC12<>""),RC12-RC8,"")'
    $wd.Range("S4:S$last").FormulaR1C1 = '=IF(OR(LEFT(RC6,3)="S7_",LEFT(RC6,3)="S8_",LEFT(RC6,3)="S9_",LEFT(RC6,4)="S10_"),"Kết thúc","Đang xử lý")'
    $wd.Range("P4:R$last").NumberFormat = '0'; $wd.Columns('A:X').ColumnWidth = 12; $wd.Columns('B').ColumnWidth = 40; $wd.Columns('G').ColumnWidth = 24
    # --- Thống kê ---
    $wt = $wb.Worksheets.Add($wb.Worksheets.Item('Dữ liệu')); $wt.Name = 'Thống kê'
    Title $wt 'THỐNG KÊ HỒ SƠ CHO PHÉP SỬ DỤNG THẺ ABTC' "Ngày lập: $NgayLap – $Nguon" 6
    $rgOf = { param($c) "'Dữ liệu'!`$$c`$4:`$$c`$$last" }
    $wt.Cells.Item(4, 1).Value2 = 'Tổng số hồ sơ'; $wt.Cells.Item(4, 2).Formula = "=COUNTA($(& $rgOf 'B'))"
    $wt.Cells.Item(5, 1).Value2 = 'Tổng số doanh nhân đề nghị'; $wt.Cells.Item(5, 2).Formula = "=SUM($(& $rgOf 'E'))"
    $wt.Cells.Item(6, 1).Value2 = 'Đã kết thúc / Đang xử lý'; $wt.Cells.Item(6, 2).Formula = "=COUNTIF($(& $rgOf 'S'),""Kết thúc"")"; $wt.Cells.Item(6, 3).Formula = "=COUNTIF($(& $rgOf 'S'),""Đang xử lý"")"
    Grid $wt.Range('A4:C6'); $wt.Range('A4:A6').Font.Bold = $true
    # A. Theo trạng thái
    $row = 9; $wt.Cells.Item($row, 1).Value2 = 'A. THEO TRẠNG THÁI'; $wt.Cells.Item($row, 1).Font.Bold = $true; $row++
    $present = @($TEN_TT.Keys | Where-Object { $k = $_; @($recs | Where-Object { (MaTT $_) -eq $k }).Count -gt 0 })
    $RT = NewList; $RT.Add([object[]]@('Trạng thái', 'Số hồ sơ', 'Số doanh nhân', 'Doanh nghiệp'))
    foreach ($k in $present) { $RT.Add([object[]]@($TEN_TT[$k], $null, $null, ((@($recs | Where-Object { (MaTT $_) -eq $k } | ForEach-Object { $_.ten_viet_tat })) -join ', '))) }
    $RT.Add([object[]]@('Cộng', $null, $null, '')); $rg = Put $wt $row 1 $RT; Grid $rg; Head $wt.Range("A$($row):D$row")
    for ($i = 1; $i -le $present.Count; $i++) { $rr = $row + $i; $wt.Cells.Item($rr, 2).Formula = "=COUNTIF($(& $rgOf 'G'),A$rr)"; $wt.Cells.Item($rr, 3).Formula = "=SUMIF($(& $rgOf 'G'),A$rr,$(& $rgOf 'E'))" }
    $tot = $row + $present.Count + 1; $wt.Cells.Item($tot, 2).Formula = "=SUM(B$($row + 1):B$($tot - 1))"; $wt.Cells.Item($tot, 3).Formula = "=SUM(C$($row + 1):C$($tot - 1))"; $wt.Range("A$($tot):D$tot").Font.Bold = $true
    # B. Thời gian xử lý
    $row = $tot + 2; $wt.Cells.Item($row, 1).Value2 = 'B. THỜI GIAN XỬ LÝ (ngày lịch)'; $wt.Cells.Item($row, 1).Font.Bold = $true; $row++
    $stages = @(@('Xin ý kiến LN → đủ ý kiến LN', 'P'), @('Đủ ý kiến LN → có KQ UB', 'Q'), @('Nhận hồ sơ → có KQ UB', 'R'))
    $RB = NewList; $RB.Add([object[]]@('Giai đoạn', 'Trung bình', 'Nhanh nhất', 'Chậm nhất', 'Số hồ sơ tính')); foreach ($s in $stages) { $RB.Add([object[]]@($s[0], $null, $null, $null, $null)) }
    $rg = Put $wt $row 1 $RB; Grid $rg; Head $wt.Range("A$($row):E$row")
    for ($i = 0; $i -lt $stages.Count; $i++) { $rr = $row + 1 + $i; $cr = & $rgOf $stages[$i][1]
      $wt.Cells.Item($rr, 2).Formula = "=IFERROR(ROUND(AVERAGE($cr),0),""-"")"; $wt.Cells.Item($rr, 3).Formula = "=IF(COUNT($cr)=0,""-"",MIN($cr))"; $wt.Cells.Item($rr, 4).Formula = "=IF(COUNT($cr)=0,""-"",MAX($cr))"; $wt.Cells.Item($rr, 5).Formula = "=COUNT($cr)" }
    # C. Liên ngành
    $row = $row + $stages.Count + 2; $wt.Cells.Item($row, 1).Value2 = 'C. Ý KIẾN LIÊN NGÀNH'; $wt.Cells.Item($row, 1).Font.Bold = $true; $row++
    $ags = @(@('Công an TP', 'T'), @('Thuế', 'U'), @('BHXH', 'V'), @('Hải quan', 'W'))
    $RC = NewList; $RC.Add([object[]]@('Cơ quan', 'Số hồ sơ đã trả lời', 'Thời gian trả lời TB (ngày, tính từ CV xin LN)')); foreach ($a in $ags) { $RC.Add([object[]]@($a[0], $null, $null)) }
    $rg = Put $wt $row 1 $RC; Grid $rg; Head $wt.Range("A$($row):C$row"); $colXin = & $rgOf 'I'
    for ($i = 0; $i -lt $ags.Count; $i++) { $rr = $row + 1 + $i; $cr = & $rgOf $ags[$i][1]
      $wt.Cells.Item($rr, 2).Formula = "=COUNT($cr)"; $wt.Cells.Item($rr, 3).Formula = "=IFERROR(ROUND(SUMPRODUCT(($cr<>"""")*($colXin<>"""")*($cr-$colXin))/SUMPRODUCT(($cr<>"""")*($colXin<>"""")),0),""-"")" }
    # D. Theo cán bộ
    $row = $row + $ags.Count + 2; $wt.Cells.Item($row, 1).Value2 = 'D. THEO CÁN BỘ THỤ LÝ'; $wt.Cells.Item($row, 1).Font.Bold = $true; $row++
    $cbs = @($recs | ForEach-Object { $_.can_bo } | Where-Object { $_ } | Sort-Object -Unique)
    $RCB = NewList; $RCB.Add([object[]]@('Cán bộ', 'Số hồ sơ', 'Đã kết thúc', 'Đang xử lý', 'Số doanh nhân')); foreach ($cb in $cbs) { $RCB.Add([object[]]@($cb, $null, $null, $null, $null)) }
    $rg = Put $wt $row 1 $RCB; Grid $rg; Head $wt.Range("A$($row):E$row")
    for ($i = 0; $i -lt $cbs.Count; $i++) { $rr = $row + 1 + $i
      $wt.Cells.Item($rr, 2).Formula = "=COUNTIF($(& $rgOf 'D'),A$rr)"; $wt.Cells.Item($rr, 3).Formula = "=COUNTIFS($(& $rgOf 'D'),A$rr,$(& $rgOf 'S'),""Kết thúc"")"
      $wt.Cells.Item($rr, 4).Formula = "=COUNTIFS($(& $rgOf 'D'),A$rr,$(& $rgOf 'S'),""Đang xử lý"")"; $wt.Cells.Item($rr, 5).Formula = "=SUMIF($(& $rgOf 'D'),A$rr,$(& $rgOf 'E'))" }
    $wt.Columns('A').ColumnWidth = 34; $wt.Columns('B:C').ColumnWidth = 16; $wt.Columns('D').ColumnWidth = 60; $wt.Columns('E').ColumnWidth = 14
    $wt.Range('B4:E200').HorizontalAlignment = -4108; $wt.Range("D10:D$tot").HorizontalAlignment = -4131; $wt.Range("D10:D$tot").WrapText = $true
    # --- Văn bản ---
    $wv = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $wv.Name = 'Văn bản'
    Title $wv 'DANH MỤC VĂN BẢN ĐÃ NHẬN DIỆN' "Ngày lập: $NgayLap – $Nguon" 8
    $RV = NewList; $RV.Add([object[]]@('Doanh nghiệp', 'File', 'Loại văn bản', 'Cơ quan', 'Số / ký hiệu', 'Ngày', 'Nguồn số', 'Nội dung chính'))
    foreach ($x in $recs) { foreach ($vb in @($x.van_ban)) { $RV.Add([object[]]@($x.ten_viet_tat, $vb.file, $vb.loai, $vb.co_quan, $vb.so_ky_hieu, $vb.ngay, $vb.nguon_so, $vb.noi_dung)) } }
    $wv.Range("A4:H$(2 + $RV.Count)").NumberFormat = '@'; $rg = Put $wv 3 1 $RV; Grid $rg; Head $wv.Range('A3:H3'); $rg.WrapText = $true
    $wv.Columns('A').ColumnWidth = 14; $wv.Columns('B').ColumnWidth = 38; $wv.Columns('C').ColumnWidth = 18; $wv.Columns('D').ColumnWidth = 10; $wv.Columns('E').ColumnWidth = 22; $wv.Columns('F').ColumnWidth = 12; $wv.Columns('G').ColumnWidth = 16; $wv.Columns('H').ColumnWidth = 60
    # --- Cảnh báo ---
    $wc = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $wc.Name = 'Cảnh báo'
    Title $wc 'CẢNH BÁO VÀ KHÁC BIỆT CẦN KIỂM TRA' "Ngày lập: $NgayLap – $Nguon" 4
    $RW = NewList; $RW.Add([object[]]@('Doanh nghiệp', 'Loại', 'Dòng Excel tổng hợp', 'Nội dung'))
    foreach ($x in $recs) {
      foreach ($cbw in @($x.canh_bao)) { if ($cbw) { $RW.Add([object[]]@($x.ten_viet_tat, 'Cảnh báo', $x.dong_excel_hien_co, $cbw)) } }
      foreach ($kb in @($x.khac_biet_voi_excel)) { if ($kb) { $RW.Add([object[]]@($x.ten_viet_tat, 'Khác với Excel tổng hợp', $x.dong_excel_hien_co, $kb)) } }
    }
    $rg = Put $wc 3 1 $RW; Grid $rg; Head $wc.Range('A3:D3'); $rg.WrapText = $true
    $wc.Columns('A').ColumnWidth = 16; $wc.Columns('B').ColumnWidth = 22; $wc.Columns('C').ColumnWidth = 12; $wc.Columns('D').ColumnWidth = 100
    foreach ($s in @($wt, $wd, $wv, $wc)) { $s.Cells.Font.Name = $font; [void]$s.Rows.AutoFit() }
    $wt.Activate(); [void]$wt.Range('A1').Select(); $wb.Save()
    "Đã tạo: $Out ($n hồ sơ, $($RV.Count - 1) văn bản, $($RW.Count - 1) cảnh báo/khác biệt)"
  } finally { if ($wb) { $wb.Close($false) }; Close-Com $xl }
}

# =====================================================================================
#  7. PHÂN TÍCH HỒ SƠ ABTC BẰNG QUY TẮC (chạy trên máy, không dùng AI, không tốn token)
#     Chắc chắn: số/ngày từ tên file, ký hiệu từ chữ, loại văn bản, luồng trạng thái, ghép dòng Excel.
#     Không chắc: đánh dấu "can_nhap" (cần nhập số/ngày tay – xem ảnh đầu trang) hoặc đưa vào cảnh báo.
# =====================================================================================
function Remove-Dau([string]$s) {
  if (-not $s) { return '' }
  $n = $s.Normalize([Text.NormalizationForm]::FormD); $sb = New-Object Text.StringBuilder
  foreach ($ch in $n.ToCharArray()) { if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) } }
  (($sb.ToString() -replace 'đ', 'd' -replace 'Đ', 'D').ToLower() -replace '\s+', ' ').Trim()
}
function Get-LoaiTheoTen([string]$n, [string]$ext) {
  # So khớp trên tên đã BỎ DẤU (tên file tiếng Việt có thể khác cách mã hóa dấu)
  $pdf = ($ext -match '^\.(pdf|jpe?g|png)$'); $a = Remove-Dau $n; $ky = ($a -match 'ban ky')   # ảnh chụp văn bản đã ký coi như PDF
  switch -Regex ($a) {
    'du thao|\(nhay\)' { return 'DU_THAO' }
    'to trinh|ttr' { if ($pdf) { return 'TRINH_UB' } else { return 'DU_THAO' } }
    '^kq|ket qua cty' { if ($pdf) { return 'KQ_UB' } else { return 'DU_THAO' } }
    'phieu trinh' { return 'PHIEU_TRINH' }
    'don doc' { if ($pdf) { return 'DON_DOC_LN' } else { return 'DU_THAO' } }
    'bo sung' { if ($pdf) { return 'YC_BO_SUNG' } else { return 'DU_THAO' } }
    'xin ln|xin (yk )?lien nganh|xin y kien|xin yk|gui thue|gui ca' { if ($pdf -and -not $ky) { return 'XIN_Y_KIEN_LN' } else { return 'DU_THAO' } }
    'tra ho so|thong tin|thong bao|khong du (dieu kien|dk)' { if ($pdf -and -not $ky) { return 'TRA_HO_SO' } else { return 'DU_THAO' } }
    '^ln|xac nhan thue' { return 'LN_TRA_LOI' }
    '^bc|bao cao' { return 'DN_HO_SO_KEM' }
    '^(cong van|cv|van ban|don|to trinh dn|tt) ?(cua )?(dn|doanh nghiep|cong ty|cty)|cong van de nghi|van ban de nghi|cv de nghi|^de nghi cap' { if ($pdf) { return 'DN_HO_SO_KEM' } else { return 'DU_THAO' } }
    default { if ($pdf) { return 'KHAC' } else { return 'DU_THAO' } }
  }
}
function Get-LoaiTheoNoiDung([string]$head, [string]$kh) {
  # Dự phòng khi tên file không nói lên loại văn bản (PDF loại KHAC): đoán theo trích yếu / ký hiệu ở phần đầu văn bản
  if (-not $head) { return 'KHAC' }
  if ($head -match 'TỜ TRÌNH' -or $kh -match 'TTr') { return 'TRINH_UB' }
  if ($kh -match 'UBND' -and $head -match 'cho phép') { return 'KQ_UB' }
  if ($head -match 'xin ý kiến về doanh nghiệp|xin ý kiến.{0,40}thẻ ABTC') { return 'XIN_Y_KIEN_LN' }
  if ($head -match 'đôn đốc|khẩn trương có ý kiến|chưa nhận được ý kiến') { return 'DON_DOC_LN' }
  if ($head -match 'thông tin kết quả hồ sơ|không đủ điều kiện|chưa đáp ứng các quy định|trả (lại )?hồ sơ') { return 'TRA_HO_SO' }
  if ($head -match 'đề nghị.{0,40}bổ sung|hoàn thiện hồ sơ') { return 'YC_BO_SUNG' }
  if ((Get-CoQuan '' $kh $head) -in 'CA', 'THUE', 'BHXH', 'HQ') { return 'LN_TRA_LOI' }
  # Công văn của doanh nghiệp gửi UBND thành phố đề nghị cho phép sử dụng thẻ ABTC (bản quét, ký hiệu dạng "…/CV-ECO")
  if ($head -match 'K[íi]nh\s*g\S{1,3}i\s*:?\s*\S{0,3}\s*[ÚUỦ]y\s*ban\s*nh|\d{4}/CV-|đề nghị cấp văn bản cho phép sử dụng thẻ ABTC') { return 'DN_HO_SO_KEM' }
  'KHAC'
}
function Get-SoNgayTuTen([string]$n) {
  $n = Remove-Dau $n
  if ($n -match '\((?:so\s*)?(\d+)\D*?(\d{1,2})\.(\d{1,2})\.(\d{2,4})\)') {
    $y = $Matches[4]; if ($y.Length -eq 2) { $y = '20' + $y }
    return @{ so = $Matches[1]; ngay = ('{0:00}/{1:00}/{2}' -f [int]$Matches[2], [int]$Matches[3], $y) }
  }
  $null
}
function Get-SoTuTenPhu([string]$n) {
  # Số (và ngày nếu đủ năm) ghi tự do trong tên tệp, dùng khi tên không theo chuẩn "(số, d.m.yy)":
  #   "yk CAHN_9447_24122025" → 9447, 24/12/2025 · "CAHN_1857_1932026" → 1857, 19/03/2026 · "THUẾ 11.9 Công ty S2_SĐ2484" → 2484
  #   "bhxh Thạch Thất - CV37.BHXH" → 37 (ký hiệu BHXH) · "cahn 6304_11.8" → 6304
  $a = Remove-Dau ([IO.Path]::GetFileNameWithoutExtension($n))
  # "yk thuế 697-HAN" → 697 (ký hiệu Thuế HAN); "… 8270LSLT", "… 8270 VP-LSLT", "… 82NV-LS" → số + ký hiệu của Văn phòng UBND / Sở Ngoại vụ
  if ($a -match '(?<![\d])(\d{2,6})\s*[-_]\s*(han|hqkv\d?|cahn|bhxh)(?![a-z])') { return @{ so = $Matches[1]; ngay = ''; ky_hieu = '' } }
  if ($a -match '(?<![\d])(\d{2,6})\s*[-_]?\s*(vp-?lslt|lslt|nv-?ls)(?![a-z])') { return @{ so = $Matches[1]; ngay = ''; ky_hieu = $(if ($Matches[2] -match '^nv') { 'NV-LS' } else { 'VP-LSLT' }) } }
  if ($a -match '(?<![\d])(\d{2,6})_(\d{6,8})(?!\d)') {
    $so = $Matches[1]; $s = $Matches[2]; $y = $s.Substring($s.Length - 4); $dm = $s.Substring(0, $s.Length - 4)
    foreach ($cut in @(@(2, 2), @(2, 1), @(1, 2), @(1, 1))) {
      if ($cut[0] + $cut[1] -ne $dm.Length) { continue }
      $d = [int]$dm.Substring(0, $cut[0]); $m = [int]$dm.Substring($cut[0])
      if ($d -ge 1 -and $d -le 31 -and $m -ge 1 -and $m -le 12 -and $y -match '^20\d\d$') { return @{ so = $so; ngay = ('{0:00}/{1:00}/{2}' -f $d, $m, $y); ky_hieu = '' } }
    }
    return @{ so = $so; ngay = ''; ky_hieu = '' }
  }
  # "SĐ2484" = SỐ ĐẾN (đăng ký văn bản đến) – không phải số văn bản → bỏ khỏi tên trước khi tìm số
  $a = ($a -replace '(?<![a-z])sd\s*[:\.]?\s*\d{1,6}', ' ')
  if ($a -match '(?<![a-z])so\s*[:\.]?\s*(\d{2,6})(?!\d)') { return @{ so = $Matches[1]; ngay = ''; ky_hieu = '' } }
  if ($a -match '(?<![a-z])cv\s*\.?\s*(\d{2,6})(?:\.([a-z]{2,6}))?(?![\d])') { return @{ so = $Matches[1]; ngay = ''; ky_hieu = $(if ($Matches[2]) { $Matches[2].ToUpper() } else { '' }) } }
  if ($a -match '(?<![\d\.])(\d{3,6})_(\d{1,2})\.(\d{1,2})(?![\d])') { return @{ so = $Matches[1]; ngay = ''; ky_hieu = '' } }
  # Một số 3–6 chữ số đứng riêng (không phải ngày d.m, năm, số thứ tự đầu tên): "3. thuế 20470 (SĐ3373)" → 20470
  $ung = @([regex]::Matches(($a -replace '^\s*\d{1,2}\s*[\.\)]\s*', ''), '(?<![\d\.\/a-z])(\d{3,6})(?![\d\.\/a-z])') | Where-Object { $_.Value -notmatch '^(19|20)\d\d$' -and $_.Value -notmatch '^0' })
  if ($ung.Count -eq 1) { return @{ so = $ung[0].Value; ngay = ''; ky_hieu = ''; yeu = $true } }   # yếu: chỉ dùng cho thư trả lời LN, sau chữ và OCR
  $null
}
function Get-NgayTuTenFolder([string]$ten) {
  # Ngày cuối tên folder doanh nghiệp: "1. Công ty N2 16.1", "AN THỊNH PHÁT_SĐ 60 (29.10)", "XNK Công ty H1 VN_12.8" → @{ d; m } (chưa có năm)
  if ($ten -match '[\s_\(](\d{1,2})\.(\d{1,2})\)?\s*$') { $d = [int]$Matches[1]; $m = [int]$Matches[2]; if ($d -ge 1 -and $d -le 31 -and $m -ge 1 -and $m -le 12) { return @{ d = $d; m = $m } } }
  $null
}

# ---- Hồ sơ CÓ CẤU TRÚC (11/2026): Tháng → Doanh nghiệp → thư mục loại văn bản → tệp ----
# Ví dụ: "THÁNG 1\1. Công ty N2 16.1\Kết quả ý kiến LN\1. bhxh hn.pdf". Folder doanh nghiệp = 1 hồ sơ; tên thư mục con cho biết loại văn bản.
$Script:RE_TEP_HS = '^\.(pdf|docx?|jpe?g|png)$'
function Test-TepRac($f) { $f.Name -like '~$*' -or $f.Name -like '~WRL*' -or $f.Name -match '^(Thumbs\.db|desktop\.ini)$' -or $f.Extension -match '^\.(tmp|lnk|bak)$' }
function Get-LoaiTheoThuMuc([string]$ten) {
  # Tên thư mục loại văn bản → mã loại; '' nếu không phải thư mục loại văn bản (vd "THÁNG 1", "1. Công ty N2 16.1"). Bỏ dấu, không phân biệt hoa/thường, chấp nhận "ý kiên", "UB/ub/Uỷ ban".
  $a = Remove-Dau $ten
  switch -Regex ($a) {
    'phieu trinh' { return 'PHIEU_TRINH' }
    'xin (y kien|yk|ln|lien nganh)' { return 'XIN_Y_KIEN_LN' }
    'ket qua (y kien|yk|ln|lien nganh)|y kien (cua )?(ln|lien nganh)|tra loi (cua )?(ln|lien nganh)' { return 'LN_TRA_LOI' }
    'ket qua (cua )?(uy ban|ub|ubnd)|^kq\b|van ban (cua )?(ub|ubnd)' { return 'KQ_UB' }
    'to trinh|\bttr\b' { return 'TRINH_UB' }
    'don doc' { return 'DON_DOC_LN' }
    'bo sung|giai trinh' { return 'YC_BO_SUNG' }
    'khong du|tra ho so|tu choi|khong dong y' { return 'TRA_HO_SO' }
    'ho so (cua )?(dn|doanh nghiep|goc)|cong van (cua )?(dn|de nghi)|de nghi cua (dn|doanh nghiep)' { return 'DN_HO_SO_KEM' }
  }
  ''
}
function Test-CongVanDN([string]$head) {
  # Công văn của doanh nghiệp (bản quét): tiêu đề bắt đầu bằng tên công ty, thường kèm "đề nghị … ABTC"
  if (-not $head) { return $false }
  $a = Remove-Dau ($head.Substring(0, [Math]::Min(160, $head.Length)))
  [bool]($a -match '^[^a-z0-9]*(tong )?c[o0]ng ty' -and (Remove-Dau $head) -match 'abtc|de nghi|apec')
}
function Test-HoSoCoCauTruc([string]$full) {
  # Không có văn bản trực tiếp nhưng có ít nhất một thư mục con mang tên loại văn bản và có văn bản bên trong
  foreach ($s in @(Get-ChildItem -LiteralPath $full -Directory -ErrorAction SilentlyContinue)) {
    if (-not (Get-LoaiTheoThuMuc $s.Name)) { continue }
    if (@(Get-ChildItem -LiteralPath $s.FullName -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match $Script:RE_TEP_HS -and -not (Test-TepRac $_) }).Count) { return $true }
  }
  $false
}
function Test-LaHoSo([string]$full) {
  # Hồ sơ = folder có PDF/Word ngay bên trong (kiểu cũ) HOẶC hồ sơ có cấu trúc (thư mục con theo loại văn bản)
  if (@(Get-ChildItem -LiteralPath $full -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(pdf|docx?)$' -and -not (Test-TepRac $_) }).Count) { return $true }
  Test-HoSoCoCauTruc $full
}
function Get-ThuTuTuNhien([string]$s) { [regex]::Replace($s, '\d+', { param($m) $m.Value.PadLeft(6, '0') }) }   # "THÁNG 2" trước "THÁNG 10"
function Find-HoSo([string]$full, [int]$sau = 0) {
  # Tìm mọi hồ sơ trong một folder tổng (đệ quy tối đa 5 cấp): Tổng → Tháng → Doanh nghiệp; không đi vào thư mục loại văn bản và thư mục hệ thống
  $kq = @()
  if (Test-LaHoSo $full) {
    $kq += Get-Item -LiteralPath $full
    # Hồ sơ phẳng (PDF/Word ngay trong folder): thư mục con là tài liệu đính kèm → dừng. Hồ sơ có cấu trúc: vẫn tìm tiếp ở các thư mục con
    # KHÔNG phải loại văn bản (trường hợp folder vừa có thư mục loại VB vừa có folder doanh nghiệp khác bên cạnh).
    if (@(Get-ChildItem -LiteralPath $full -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(pdf|docx?)$' -and -not (Test-TepRac $_) }).Count) { return $kq }
  }
  if ($sau -ge 5) { return $kq }
  $con = @(Get-ChildItem -LiteralPath $full -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -ne $Script:HE_THONG -and $_.Name -notmatch '^(_HE_THONG|KẾT QUẢ|DỮ LIỆU GỐC|Báo cáo tuần|\..*)$' -and -not (Get-LoaiTheoThuMuc $_.Name) } |
    Sort-Object { Get-ThuTuTuNhien $_.Name } | ForEach-Object { Find-HoSo $_.FullName ($sau + 1) })
  # Bên trong một hồ sơ có cấu trúc chỉ nhận hồ sơ có cấu trúc khác (folder doanh nghiệp), không nhận thư mục tài liệu đính kèm có PDF
  if ($kq.Count) { $con = @($con | Where-Object { Test-HoSoCoCauTruc $_.FullName }) }
  $kq + $con
}
function Get-CacheBase([string]$key) {
  # Tên gốc của các tệp cache cho một văn bản. Tệp trong thư mục con: "<thư mục> - <tên>" để hai tệp cùng tên ở hai thư mục không đè nhau
  $b = [IO.Path]::GetFileNameWithoutExtension($key); $p = Split-Path $key -Parent
  if ($p) { (($p -replace '[\\/]+', ' - ') + ' - ' + $b) } else { $b }
}
function Get-TepHoSo([string]$full) {
  # Văn bản của 1 hồ sơ: tệp ngay trong folder + tệp trong các thư mục con mang tên loại văn bản (không lấy thư mục con khác, vd tài liệu đính kèm của DN).
  # Key = đường dẫn tương đối (vd "Kết quả ý kiến LN\1. bhxh hn.pdf"); LoaiThuMuc = loại suy từ tên thư mục con.
  $root = (Get-Item -LiteralPath $full).FullName.TrimEnd('\')
  $ds = @(Get-ChildItem -LiteralPath $root -File -ErrorAction SilentlyContinue)
  foreach ($s in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) { if (Get-LoaiTheoThuMuc $s.Name) { $ds += @(Get-ChildItem -LiteralPath $s.FullName -File -Recurse -ErrorAction SilentlyContinue) } }
  @($ds | Where-Object { $_.Extension -match $Script:RE_TEP_HS -and -not (Test-TepRac $_) } | ForEach-Object {
    $key = $_.FullName.Substring($root.Length + 1); $tm = Split-Path $key -Parent
    [pscustomobject]@{ Key = $key; File = $_; ThuMuc = $tm; LoaiThuMuc = $(if ($tm) { Get-LoaiTheoThuMuc (($tm -split '\\')[0]) } else { '' }); CacheBase = (Get-CacheBase $key) }
  } | Sort-Object Key)
}
function Format-Ngay([string]$s) { if ($s -match '(\d{1,2})/(\d{1,2})/(\d{4})') { '{0:00}/{1:00}/{2}' -f [int]$Matches[1], [int]$Matches[2], $Matches[3] } else { $s } }
function ConvertTo-Date([string]$s) { if ($s -match '(\d{1,2})/(\d{1,2})/(\d{4})') { try { return [datetime]::new([int]$Matches[3], [int]$Matches[2], [int]$Matches[1]) } catch {} }; $null }
function Get-CoQuan([string]$name, [string]$kh, [string]$text) {
  if ($kh) {
    if ($kh -match 'CAHN|^CA') { return 'CA' }; if ($kh -match 'HQ') { return 'HQ' }; if ($kh -match 'BHXH') { return 'BHXH' }
    if ($kh -match 'TMĐT|TCS|HAN-|QLDN|CNHKD|QLT3') { return 'THUE' }; if ($kh -match 'TTr') { return 'Sở NV' }
    if ($kh -match 'UBND') { return 'UBND' }; if ($kh -match 'VP-LSLT|NV-LS') { return 'Phòng' }
  }
  $a = Remove-Dau $name
  # (?<![a-z]) thay cho \b: tên tệp hay nối bằng "_" (vd "CV_CAHN_825") mà "_" được tính là chữ trong \b
  if ($a -match 'thue') { return 'THUE' }; if ($a -match '(?<![a-z])hq|hai quan') { return 'HQ' }; if ($a -match '(?<![a-z])bh') { return 'BHXH' }; if ($a -match '(?<![a-z])ca(?![a-z])|(?<![a-z])cahn|cong an|catp|(?<![a-z])ca ?tp') { return 'CA' }
  $h = if ($text.Length -gt 300) { $text.Substring(0, 300) } else { $text }
  if ($h -match 'CÔNG AN') { return 'CA' }; if ($h -match 'HẢI QUAN') { return 'HQ' }; if ($h -match 'BẢO HIỂM XÃ HỘI') { return 'BHXH' }; if ($h -match 'THUẾ') { return 'THUE' }
  ''
}
function Get-VietTat([string]$dn) {
  # Bỏ loại hình + cụm từ chung (so khớp theo từ đã bỏ dấu), giữ phần tên riêng. Vd: "Công ty cổ phần đầu tư nhà và thương mại Công ty D" -> "Công ty D"
  if (-not $dn) { return '' }
  $words = @($dn -split '\s+' | Where-Object { $_ }); $asc = @($words | ForEach-Object { (Remove-Dau $_) -replace '[^a-z0-9&]', '' })
  $chung = @('tong cong ty', 'cong ty', 'co phan', 'cp', 'tnhh', 'mtv', 'mot thanh vien', 'tap doan', 'dau tu nha', 'dau tu', 'thuong mai', 'xuat nhap khau', 'quoc te', 'viet nam', 'vn',
    'cong nghe', 'dich vu truong hoc', 'dich vu', 'san xuat', 'phat trien', 'va', '&', 'ctcp', 'jsc') | ForEach-Object { , @($_ -split ' ') }
  $keep = @(); $i = 0
  while ($i -lt $words.Count) {
    $bo = 0; foreach ($c in $chung) { if ($i + $c.Count -le $words.Count -and (($asc[$i..($i + $c.Count - 1)]) -join ' ') -eq ($c -join ' ')) { $bo = $c.Count; break } }
    if ($bo) { $i += $bo } else { $keep += $words[$i]; $i++ }
  }
  $x = ($keep -join ' ').Trim().TrimEnd(' ', '-', '–', ',').Trim(); if (-not $x) { $x = $dn }; $x.Substring(0, 1).ToUpper() + $x.Substring(1)   # bỏ gạch nối thừa: "X - Công ty cổ phần" → "X"
}

# ---- Đổi tên folder hồ sơ / file văn bản (cache, nhập tay đi theo) và gợi ý tên chuẩn ----
function Get-TenChuan([string]$file, [string]$so, [string]$ngay, [string]$loai) {
  # Gợi ý tên file chuẩn "<tên> (<số>, <d>.<m>.<yy>)<ext>" để lần sau đọc số/ngày từ tên file, không cần OCR. Trả '' nếu không cần hoặc không đủ dữ liệu.
  if ($loai -notin 'XIN_Y_KIEN_LN', 'LN_TRA_LOI', 'TRINH_UB', 'KQ_UB', 'YC_BO_SUNG', 'DON_DOC_LN', 'TRA_HO_SO') { return '' }
  if (-not $so -or $so -notmatch '^\d+$' -or -not $ngay -or (Get-SoNgayTuTen $file)) { return '' }
  if ($ngay -notmatch '^(\d{1,2})/(\d{1,2})/(\d{4})$') { return '' }
  $d = [int]$Matches[1]; $mo = [int]$Matches[2]; $yy = $Matches[3].Substring(2)
  $ext = [IO.Path]::GetExtension($file); $base = (([IO.Path]::GetFileNameWithoutExtension($file)) -replace '\s+', ' ').Trim().TrimEnd(',', '.', '-', '_', ' ')
  if (-not $base) { return '' }
  '{0} ({1}, {2}.{3}.{4}){5}' -f $base, $so, $d, $mo, $yy, $ext
}
function Test-TenHopLe([string]$ten) {
  # Tên file/folder hợp lệ trên Windows: không ký tự cấm, không bắt đầu bằng ".", không kết thúc bằng dấu cách/chấm
  [bool]($ten -and $ten.Trim() -ceq $ten -and $ten -notmatch '[\\/:*?"<>|]' -and $ten -notmatch '^\.' -and $ten -notmatch '[. ]$' -and $ten.Length -le 200)
}
function Rename-HoSoFolder {
  # Đổi tên folder hồ sơ; chuyển cache (_HE_THONG\cache\<tên>) đi theo để không phải đọc PDF/OCR lại. Trả về đường dẫn mới.
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$TenMoi)
  if (-not (Test-TenHopLe $TenMoi)) { throw "Tên folder không hợp lệ: '$TenMoi' (không dùng \ / : * ? `" < > |, không bắt đầu bằng dấu chấm)" }
  $dir = Get-Item -LiteralPath $Path
  if ($dir.Name -ceq $TenMoi) { return $dir.FullName }
  $dich = Join-Path $dir.Parent.FullName $TenMoi
  if ((Test-Path -LiteralPath $dich) -and $dir.Name -ne $TenMoi) { throw "Đã có folder '$TenMoi' trong cùng thư mục" }
  try { Rename-Item -LiteralPath $dir.FullName -NewName $TenMoi -ErrorAction Stop }
  catch { throw "Không đổi được tên folder – có file đang mở trong Word/Excel hoặc OneDrive đang đồng bộ. Đóng file rồi thử lại. ($($_.Exception.Message))" }
  $cCu = Join-Path $Script:HE_THONG ('cache\' + $dir.Name); $cMoi = Join-Path $Script:HE_THONG ('cache\' + $TenMoi)
  if (Test-Path -LiteralPath $cCu) {
    if ((Test-Path -LiteralPath $cMoi) -and $dir.Name -ne $TenMoi) { Remove-Item -LiteralPath $cMoi -Recurse -Force }
    Rename-Item -LiteralPath $cCu -NewName $TenMoi -Force
  }
  $dich
}
function Rename-HoSoTep {
  # Đổi tên một file trong folder hồ sơ (giữ phần mở rộng); đổi tên các file cache (.txt, _dau.png/.jpg, _dau.ocr.txt) và khóa trong manifest.json / ghi_tay.json đi theo.
  # -File có thể là đường dẫn tương đối trong hồ sơ có cấu trúc ("Kết quả ý kiến LN\1. bhxh hn.pdf"); -TenMoi chỉ là TÊN tệp mới (đổi tại chỗ, không chuyển thư mục). Trả về khóa mới.
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$File, [Parameter(Mandatory)][string]$TenMoi)
  $dir = Get-Item -LiteralPath $Path; $src = Join-Path $dir.FullName $File
  if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { throw "Không thấy file '$File' trong folder" }
  if (-not (Test-TenHopLe $TenMoi)) { throw "Tên file không hợp lệ: '$TenMoi' (không dùng \ / : * ? `" < > |)" }
  $extCu = [IO.Path]::GetExtension($File)
  if ([IO.Path]::GetExtension($TenMoi) -ne $extCu) { throw "Phải giữ phần mở rộng '$extCu'" }
  $sub = Split-Path $File -Parent; $tenCu = Split-Path $File -Leaf
  $keyMoi = if ($sub) { Join-Path $sub $TenMoi } else { $TenMoi }
  if ($tenCu -ceq $TenMoi) { return $File }
  if ((Test-Path -LiteralPath (Join-Path $dir.FullName $keyMoi)) -and $tenCu -ne $TenMoi) { throw "Đã có file '$TenMoi' trong cùng thư mục" }
  try { Rename-Item -LiteralPath $src -NewName $TenMoi -ErrorAction Stop }
  catch { throw "Không đổi được tên '$tenCu' – file đang mở trong Word/Excel? Đóng file rồi thử lại. ($($_.Exception.Message))" }
  $cache = Join-Path $Script:HE_THONG ('cache\' + $dir.Name)
  if (-not (Test-Path -LiteralPath $cache)) { return $keyMoi }
  $bCu = Get-CacheBase $File; $bMoi = Get-CacheBase $keyMoi
  if ($bCu -cne $bMoi) {
    foreach ($suf in '.txt', '_dau.png', '_dau.jpg', '_dau.jpeg', '_dau.ocr.txt') {
      $a = Join-Path $cache ($bCu + $suf); if (-not (Test-Path -LiteralPath $a)) { continue }
      $b = Join-Path $cache ($bMoi + $suf); if ((Test-Path -LiteralPath $b) -and $bCu -ne $bMoi) { Remove-Item -LiteralPath $b -Force }
      Rename-Item -LiteralPath $a -NewName ($bMoi + $suf) -Force
    }
  }
  foreach ($jn in 'manifest.json', 'ghi_tay.json') {
    $jp = Join-Path $cache $jn; if (-not (Test-Path -LiteralPath $jp)) { continue }
    $h = [ordered]@{}; (Get-Content -LiteralPath $jp -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
    if (-not $h.Contains($File)) { continue }
    $v = $h[$File]; $h.Remove($File)
    if ($jn -eq 'manifest.json') {
      if ($v.txt) { $v.txt = $bMoi + '.txt' }
      if ($v.anh -and $v.anh -notmatch '^LỖI') { $v.anh = $bMoi + '_dau' + [IO.Path]::GetExtension([string]$v.anh) }
      $sn = Get-SoNgayTuTen $TenMoi; if ($sn) { $v.so = $sn.so; $v.ngay = $sn.ngay }
    }
    $h[$keyMoi] = $v
    [IO.File]::WriteAllText($jp, ($h | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
  }
  $keyMoi
}
function Read-ExcelRows([string]$Excel, [string]$Sheet = 'ABTC 2026') {
  # Máy mới chưa có file Excel tổng hợp: coi như bảng trống (mọi hồ sơ là dòng mới, bắt đầu từ dòng 5) thay vì báo lỗi
  if (-not $Excel -or -not (Test-Path -LiteralPath $Excel)) { return @{ rows = @{}; maxRow = 4; maxStt = 0; thieu_file = $true } }
  $txt = Read-Xlsx -Path $Excel; $rows = @{}; $in = $false; $maxRow = 0; $maxStt = 0
  foreach ($line in [IO.File]::ReadAllLines($txt, [Text.Encoding]::UTF8)) {
    if ($line -like '===== SHEET:*') { $in = ($line -eq "===== SHEET: $Sheet =====") ; continue }
    if (-not $in) { continue }
    foreach ($tok in ($line -split ' \| ')) {
      if ($tok -match '^([A-Z]+)(\d+)=(.*)$') {
        $r = [int]$Matches[2]; if (-not $rows.ContainsKey($r)) { $rows[$r] = @{} }; $rows[$r][$Matches[1]] = $Matches[3]
        if ($r -gt $maxRow) { $maxRow = $r }
        if ($Matches[1] -eq 'A' -and $Matches[3] -match '^\d+(\.0)?$') { $v = [int][double]$Matches[3]; if ($v -gt $maxStt) { $maxStt = $v } }
      }
    }
  }
  Remove-Item $txt -Force -ErrorAction SilentlyContinue
  @{ rows = $rows; maxRow = $maxRow; maxStt = $maxStt }
}

function Invoke-PhanTich {
  # Phân tích 1 folder hồ sơ ABTC bằng quy tắc. Kết quả (khuôn giống ket_qua.json + "de_xuat") ghi ở cache\<folder>\phan_tich.json
  param([Parameter(Mandatory)][string]$Path, [string]$Excel = (P 'KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx'), [string]$Out, [switch]$Im)
  Invoke-ChuanBiFolder -Path $Path | Out-Null
  $dir = Get-Item -LiteralPath $Path
  $cache = Join-Path $Script:HE_THONG ('cache\' + $dir.Name)
  $man = @{}; (Get-Content (Join-Path $cache 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $man[$_.Name] = $_.Value }
  $tay = @{}; $tayPath = Join-Path $cache 'ghi_tay.json'
  if (Test-Path $tayPath) { (Get-Content $tayPath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $tay[$_.Name] = $_.Value } }
  $canhBao = New-Object System.Collections.ArrayList

  # ---- 1. Văn bản ----
  $docs = @(); $daThay = @{}
  foreach ($tp0 in @(Get-TepHoSo $dir.FullName)) {
    # $k = khóa văn bản (đường dẫn tương đối, vd "Kết quả ý kiến LN\1. bhxh hn.pdf"); $f.Name = tên tệp để đoán loại/số theo tên
    $f = $tp0.File; $k = $tp0.Key; $goiY = $tp0.LoaiThuMuc
    $m = $man[$k]; $t = ''
    if ($m -and $m.txt) { $tp = Join-Path $cache $m.txt; if (Test-Path -LiteralPath $tp) { $t = ([IO.File]::ReadAllText($tp, [Text.Encoding]::UTF8)) -replace "`a", ' ' } }
    $laBanKy = ($f.Extension -match '^\.(pdf|jpe?g|png)$')
    $loai = if ($f.Extension -eq '.doc') { [void]$canhBao.Add("${k}: định dạng .doc cũ, chưa đọc được – mở bằng Word và lưu lại thành .docx"); 'KHAC' } else { Get-LoaiTheoTen $f.Name $f.Extension }
    if ($m -and $daThay.ContainsKey($m.hash)) { $loai = 'TRUNG' } elseif ($m) { $daThay[$m.hash] = $k }
    $head = if ($t.Length -gt 900) { $t.Substring(0, 900) } else { $t }
    # Hồ sơ có cấu trúc: thư mục loại văn bản quyết định loại (trừ phiếu trình, đôn đốc, tờ trình, dự thảo – nhận theo tên tệp);
    # riêng công văn của DN để lẫn trong thư mục (tiêu đề "CÔNG TY …") nhận là hồ sơ DN. Thư mục được ưu tiên hơn đoán theo nội dung
    # vì thư trả lời của LN cũng nhắc "xin ý kiến về doanh nghiệp".
    # Tệp văn thư quét đặt tên "số_vc_năm" (vd "477_vc_2026.pdf", "3586_vc_2026.pdf"): công văn đến của doanh nghiệp
    $vanThuDN = ($laBanKy -and $loai -notin 'TRUNG', 'PHIEU_TRINH', 'DON_DOC_LN', 'TRINH_UB', 'DU_THAO' -and ($f.Name -match '(?<!\d)\d{2,6}_vc_\d{4}' -or (Remove-Dau $f.Name) -match 'giay nop tien|nop tien vao ngan sach|nsnn|uy nhiem chi|bien lai|giai trinh' -or $head.Substring(0, [Math]::Min(200, $head.Length)) -cmatch 'GIẤY NỘP TIỀN'))   # cả giấy nộp tiền NSNN / ủy nhiệm chi: chứng từ của DN
    if ($vanThuDN) { $loai = 'DN_HO_SO_KEM' }
    elseif ($goiY -and $laBanKy -and $loai -notin 'PHIEU_TRINH', 'DON_DOC_LN', 'TRINH_UB', 'DU_THAO', 'TRUNG') { $loai = if (Test-CongVanDN $head) { 'DN_HO_SO_KEM' } else { $goiY } }
    $soTxt = ''; $kh = ''
    if ($loai -notin 'PHIEU_TRINH', 'DU_THAO' -and $head -match 'Số\s*:?\s*(\d*)\s*/\s*([A-ZĐ][A-ZĐ0-9\.\-&+]*(?:\s?-\s?[A-ZĐ0-9]+)*)') { $soTxt = $Matches[1]; $kh = ($Matches[2] -replace '\s', '') }   # phiếu trình/dự thảo: không lấy số (chỉ là số văn bản được nhắc tới)
    if ($loai -eq 'KHAC' -and $laBanKy) { $loai = Get-LoaiTheoNoiDung $head $kh; if ($loai -eq 'KHAC' -and (Test-CongVanDN $head)) { $loai = 'DN_HO_SO_KEM' }; if ($loai -ne 'KHAC') { [void]$canhBao.Add("${k}: tên file không nói rõ loại, đã đoán theo nội dung là $loai – nên đặt tên lại cho rõ") } }
    $ngTxt = ''; $hd = if ($t.Length -gt 1600) { $t.Substring(0, 1600) } else { $t }
    if ($hd -match 'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})') { $ngTxt = '{0:00}/{1:00}/{2}' -f [int]$Matches[1], [int]$Matches[2], $Matches[3] }
    $sn = Get-SoNgayTuTen $f.Name
    $so = ''; $ngay = ''; $nguon = @()
    if ($tay.ContainsKey($k)) { $so = [string]$tay[$k].so; $ngay = [string]$tay[$k].ngay; if ($tay[$k].ky_hieu) { $kh = [string]$tay[$k].ky_hieu }; $nguon += 'nhập tay' }
    if (-not $so -and $sn) { $so = $sn.so; $nguon += 'tên file' }
    if (-not $so -and $loai -notin 'PHIEU_TRINH', 'DU_THAO', 'DN_HO_SO_KEM', 'TRUNG') {
      $sp = Get-SoTuTenPhu $f.Name
      if ($sp -and -not $sp.yeu) { $so = $sp.so; if (-not $kh -and $sp.ky_hieu) { $kh = $sp.ky_hieu }; if (-not $ngay -and $sp.ngay) { $ngay = $sp.ngay }; $nguon += 'tên file' }
    }
    if (-not $so -and $soTxt) { $so = $soTxt; $nguon += 'chữ' }
    if (-not $ngay -and $sn) { $ngay = $sn.ngay }
    if (-not $ngay -and $ngTxt) { $ngay = $ngTxt; if ($nguon -notcontains 'chữ') { $nguon += 'chữ' } }
    # OCR ảnh đầu trang (số đóng tay/đóng dấu chỉ có trên ảnh): lấy số đứng trước ký hiệu, và ngày "ngày .. tháng .. năm .."
    $ocrPath = if ($m -and $m.anh -and $m.anh -notmatch '^LỖI') { Join-Path $cache ($m.anh -replace '\.(png|jpe?g)$', '.ocr.txt') } else { '' }
    if ((-not $so -or -not $ngay) -and $loai -ne 'DN_HO_SO_KEM' -and $ocrPath -and (Test-Path -LiteralPath $ocrPath)) {
      $ocr = [IO.File]::ReadAllText($ocrPath, [Text.Encoding]::UTF8); $khKey = ((Remove-Dau $kh) -replace '[^a-z]', '').ToUpper()
      if (-not $so) {
        # Số đóng dấu thường dạng "1234/NV-LS"; OCR hay làm mất dấu "/" → "1284 NV-LS": cho phép chỉ có khoảng trắng (cùng dòng)
        # Thêm: số dính liền ký hiệu "8270VP-LSLT" (chỉ nhận khi ngay sau là VP-/NV-/NP-)
        foreach ($mm in [regex]::Matches($ocr, '(?<![\d/])(\d{2,6})(?:\s*(?:(?:[^\w\s]|[1lI])\s*)+|([ \t]+)|(?=(?:VP|NV|NP)-))([A-Z]{2,})')) {
          $g2 = ($mm.Groups[3].Value -replace '^(NP|UP|YP|WP)$', 'VP')   # OCR hay đọc "VP" thành "NP"
          if ($loai -eq 'LN_TRA_LOI' -and $g2 -match '^(VP|NV)') { continue }   # bỏ số CV của phòng được viện dẫn
          # số quá ngắn: không tin – trừ ký hiệu NV-LS của Sở Ngoại vụ (đánh số lại từ 7/2026 nên có số 2 chữ số, vd "82 NV-LS")
          if ($mm.Groups[1].Value.Length -lt 3 -and -not ($g2 -match '^NV' -and $mm.Groups[1].Value -notmatch '^0')) { continue }
          if ($mm.Groups[2].Success -and $mm.Groups[1].Value -match '^20\d\d$') { continue }
          if ($mm.Groups[1].Value -match '^0') { continue }   # số bắt đầu bằng 0: OCR có thể đã mất chữ số đầu (vd "sæ0754VP-LSLT") – để nhập tay   # "2026 UBND…": năm, không phải số văn bản
          $khop = if ($khKey) { $khKey.StartsWith($g2.Substring(0, 2)) -or $g2.StartsWith($khKey.Substring(0, [Math]::Min(2, $khKey.Length))) } else { $g2 -match '^(CAHN|HQ|BHXH|BC|HAN|TCS|TMDT|TB|UBND|TTR)' }
          if ($khop) { $so = $mm.Groups[1].Value; $nguon += 'OCR'; break }
        }
        # Dự phòng: có nhãn "Số:" – số và ký hiệu ở 2 dòng ("so: 131" ⏎ "NP-LSLT") hoặc dính liền ("sæ10754VP-LSLT"); bỏ số bắt đầu bằng 0 (OCR có thể đã mất chữ số đầu)
        if (-not $so -and $ocr -match '(?im)^\s*s(?:[^\s\d]{0,2}\s*[:\.]?|\S{1,2}\s*:)\s*([1-9]\d{1,5})\s*(?:[\r\n]+\s*)?([A-Z]{2,})') {
          $g2 = ($Matches[2] -replace '^(NP|UP|YP|WP)', 'VP')
          $khop = if ($khKey) { $khKey.StartsWith($g2.Substring(0, 2)) } else { $g2 -match '^(VP|NV|CAHN|HQ|BHXH|UBND|TTR)' }
          if ($khop) { $so = $Matches[1]; $nguon += 'OCR' }
        }
      }
      if (-not $ngay -and $ocr -match 'ng\S{0,3}y\s*([\dIlO|]{1,2})\s*th\S{0,3}ng\s*([\dIlO|]{1,2})\s*n\S{0,3}m\s*(20\d{2})') { $dg = ($Matches[1] -replace '[Il|]', '1' -replace 'O', '0'); $mg = ($Matches[2] -replace '[Il|]', '1' -replace 'O', '0'); $ngay = '{0:00}/{1:00}/{2}' -f [int]$dg, [int]$mg, $Matches[3]; if ($nguon -notcontains 'OCR') { $nguon += 'OCR' } }
    }
    # ---- Văn bản ký số / bản quét của LN (Thuế, Hải quan, Công an…) ----
    if ($laBanKy -and $loai -in 'LN_TRA_LOI', 'KQ_UB', 'TRA_HO_SO', 'YC_BO_SUNG', 'DON_DOC_LN', 'XIN_Y_KIEN_LN') {
      $ocrTxt = if ($ocrPath -and (Test-Path -LiteralPath $ocrPath)) { [IO.File]::ReadAllText($ocrPath, [Text.Encoding]::UTF8) } else { '' }
      $nguonDoc = ($head + "`n" + $ocrTxt)
      # (a) Số theo MỐC KÝ HIỆU: số đứng ngay trước "…QLDN / NVHQ / ANKT / QLT" trên cùng dòng (hoặc dòng trên), chịu được ký hiệu bị OCR làm méo:
      #     "311f /T--IAN-QLDNI" → 311 · "16603 rrCS8-QLDN1" → 16603 · "sé: 22 /HQKVI-NVHQ" → 22. Bỏ số kèm "?" (chữ số đọc không ra), số bắt đầu bằng 0, năm 20xx.
      if (-not $so -and $loai -eq 'LN_TRA_LOI') {
        $dongs = @($nguonDoc -split "`n")
        for ($li = 0; $li -lt $dongs.Count -and -not $so; $li++) {
          $mk = [regex]::Match($dongs[$li], '(?i)(QLDN|NV[HI1l]{1,2}Q|ANKT|QLT\b|BHXH-)')
          if (-not $mk.Success) { continue }
          $truoc = $dongs[$li].Substring(0, $mk.Index); if ($truoc.Trim().Length -lt 3 -and $li -gt 0) { $truoc = $dongs[$li - 1] + ' ' + $truoc }
          $ms = @([regex]::Matches($truoc, '(?<![\dA-Za-zĐđ])(\d{2,6})(?![\d\?])') | Where-Object { $_.Value -notmatch '^0' -and $_.Value -notmatch '^20\d\d$' })
          if ($ms.Count) { $so = $ms[-1].Value; $nguon += 'OCR' }
        }
      }
      if (-not $so -and $loai -eq 'LN_TRA_LOI') { $sp2 = Get-SoTuTenPhu $f.Name; if ($sp2 -and $sp2.yeu) { $so = $sp2.so; $nguon += 'tên file' } }
      # (b) Ngày theo CHỮ KÝ SỐ: "Ký bởi: THUẾ THÀNH PHỐ HÀ NỘI – Thời gian ký: 09.01.2026 14:32" (đáng tin hơn dòng "ngày … tháng …" bị OCR làm hỏng)
      if (-not $ngay) {
        foreach ($nd in @($hd, $ocrTxt)) {
          if ($nd -match '(?i)Th\S{0,3}i\s*gian\s*(?:k\S{0,3})?\s*[:\.]?\s*(\d{1,2})[\./]?(\d{1,2})[\./](20\d{2})') {
            $dd0 = [int]$Matches[1]; $mm0 = [int]$Matches[2]
            if ($dd0 -ge 1 -and $dd0 -le 31 -and $mm0 -ge 1 -and $mm0 -le 12) { $ngay = '{0:00}/{1:00}/{2}' -f $dd0, $mm0, $Matches[3]; $nguon += 'thời gian ký số'; break }
          }
        }
      }
      # (c) Ngày chứng thực bản sao điện tử: "SAO Y; Công an Thành phố Hà Nội; 04/11/2025; +07:00" (dự phòng cuối)
      if (-not $ngay -and $nguonDoc -match '(?i)SAO\s*Y\s*;[^;\n]{0,80};\s*(\d{1,2})/(\d{1,2})/(20\d{2})') {
        $ngay = '{0:00}/{1:00}/{2}' -f [int]$Matches[1], [int]$Matches[2], $Matches[3]; $nguon += 'ngày sao y'
      }
    }
    $coQuan = if ($loai -in 'PHIEU_TRINH', 'DU_THAO') { 'Phòng' } else { Get-CoQuan $f.Name $kh $t }
    $canNhap = ($loai -in 'XIN_Y_KIEN_LN', 'LN_TRA_LOI', 'TRINH_UB', 'KQ_UB', 'YC_BO_SUNG', 'DON_DOC_LN', 'TRA_HO_SO') -and (-not $so -or -not $ngay)
    if ($loai -eq 'PHIEU_TRINH' -and -not $ngay) { $canNhap = $true }
    $docs += [pscustomobject]@{ file = $k; thu_muc = $tp0.ThuMuc; ext = $f.Extension.ToLower(); loai = $loai; co_quan = $coQuan; so = $so; ky_hieu = $kh; ngay = $ngay
      so_ky_hieu = $(if ($so -and $kh) { "$so/$kh" } elseif ($so) { $so } else { '' }); nguon_so = ($nguon -join ' + '); anh = $(if ($m) { $m.anh } else { '' })
      can_nhap = $canNhap; noi_dung = ''; text = $t; khac_dn = $false }
  }
  # ---- 1b. Số/ngày CV xin ý kiến LN chưa đọc được (OCR mất dòng số, số viết tay): lấy từ văn bản khác TRÍCH DẪN lại ----
  # LN trả lời: "Phúc đáp Công văn số 12345/VP-LSLT ngày 10/8/2026…"; tờ trình / VB UB / phiếu trình cũng nhắc lại. Chọn số được trích nhiều nhất,
  # khớp ký hiệu và (nếu đã biết) khớp số hoặc ngày của CV xin; bỏ số đã là của văn bản khác (đôn đốc, bổ sung…).
  $daBiet = @{}; foreach ($d in $docs) { if ($d.so -and $d.loai -ne 'XIN_Y_KIEN_LN') { $daBiet["$($d.so)"] = 1 } }
  $trich = @(foreach ($d in @($docs | Where-Object { $_.text -and $_.loai -in 'LN_TRA_LOI', 'DON_DOC_LN', 'TRINH_UB', 'KQ_UB', 'PHIEU_TRINH', 'DU_THAO', 'TRA_HO_SO' })) {
    foreach ($mm in [regex]::Matches($d.text, '(?:[Cc]ông văn|[Vv]ăn bản|CV)\s+số\s*:?\s*(\d{2,6})\s*/\s*(VP-LSLT|NV-LS)\s*,?\s*(?:ngày\s+)?(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})')) {
      [pscustomobject]@{ so = $mm.Groups[1].Value; kh = $mm.Groups[2].Value; ngay = ('{0:00}/{1:00}/{2}' -f [int]$mm.Groups[3].Value, [int]$mm.Groups[4].Value, $mm.Groups[5].Value); tu = $d.file }
    }
  })
  foreach ($x in @($docs | Where-Object { $_.loai -eq 'XIN_Y_KIEN_LN' -and (-not $_.so -or -not $_.ngay) } | Sort-Object { $d0 = ConvertTo-Date $_.ngay; if ($d0) { $d0 } else { [datetime]::MaxValue } })) {
    $ung = @($trich | Where-Object { -not $daBiet.ContainsKey($_.so) -and (-not $x.ky_hieu -or $_.kh -eq $x.ky_hieu) })
    if ($x.so) { $ung = @($ung | Where-Object { $_.so -eq $x.so }) } elseif ($x.ngay) { $ung = @($ung | Where-Object { $_.ngay -eq $x.ngay }) }
    if (-not $ung.Count) { continue }
    $nhom = $ung | Group-Object so | Sort-Object @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { ConvertTo-Date $_.Group[0].ngay } } | Select-Object -First 1
    $chon = $nhom.Group[0]
    if (-not $x.so) { $x.so = $chon.so }; if (-not $x.ky_hieu) { $x.ky_hieu = $chon.kh }; if (-not $x.ngay) { $x.ngay = $chon.ngay }
    $x.so_ky_hieu = "$($x.so)/$($x.ky_hieu)"; $x.can_nhap = (-not $x.so -or -not $x.ngay); $daBiet["$($x.so)"] = 1
    $x.nguon_so = ((@($x.nguon_so, "trích dẫn trong $($nhom.Count) VB khác") | Where-Object { $_ }) -join ' + ')
  }
  # ---- 1c. Số/ngày văn bản LN TRẢ LỜI: đối chiếu với tờ trình / phiếu trình kết quả / VB UB / dự thảo – nơi cán bộ gõ lại đủ
  #      "… Hải quan (Công văn số 1234/HQKV1-NVHQ ngày 03/12/2025), BHXH (1234/BHXH-BPT ngày 03/12/2025), Công an (1234/CAHN-ANKT ngày 24/12/2025), Thuế (1234/HAN-QLDN1 ngày 30/01/2026)"
  #      Ưu tiên hơn số/ngày đọc bằng OCR, chữ ký số, ngày sao y (chữ gõ đáng tin hơn); không đè số nhập tay / tên file / lớp chữ.
  $refLN = @(foreach ($d in @($docs | Where-Object { $_.text -and $_.loai -in 'TRINH_UB', 'PHIEU_TRINH', 'KQ_UB', 'DU_THAO' -and -not $_.khac_dn })) {
    foreach ($mm in [regex]::Matches($d.text, '(?<![\d/])(\d{2,6})\s*/\s*((?:[A-Z0-9Đ]+-)*(?:HAN|TCS\d*|CS\d+|HQKV\w*|HQ\w*|CAHN|BHXH)[A-Z0-9Đ\-]*)\s*,?\s*(?:ngày\s*)?(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})')) {
      $cq0 = Get-CoQuan '' $mm.Groups[2].Value ''
      if ($cq0 -in 'CA', 'THUE', 'BHXH', 'HQ') { [pscustomobject]@{ cq = $cq0; so = $mm.Groups[1].Value; kh = $mm.Groups[2].Value; ngay = ('{0:00}/{1:00}/{2}' -f [int]$mm.Groups[3].Value, [int]$mm.Groups[4].Value, $mm.Groups[5].Value) } }
    }
  })
  # Mỗi số một bản ghi; ngày lấy theo đa số (văn bản có thể gõ nhầm ngày ở một chỗ)
  $mocXin = ($docs | Where-Object { $_.loai -eq 'XIN_Y_KIEN_LN' } | ForEach-Object { ConvertTo-Date $_.ngay } | Where-Object { $_ } | Measure-Object -Minimum).Minimum
  $refLN = @($refLN | Where-Object { $_.so -notmatch '^0' -and (-not $mocXin -or (ConvertTo-Date $_.ngay) -ge $mocXin.AddDays(-1)) })   # văn bản trước ngày xin ý kiến không thể là thư trả lời
  $refLN = @($refLN | Group-Object so | ForEach-Object { $g = $_.Group; $nd = ($g | Group-Object ngay | Sort-Object Count -Descending | Select-Object -First 1).Name; [pscustomobject]@{ cq = $g[0].cq; so = $_.Name; kh = $g[0].kh; ngay = $nd } })
  $nguonYeu = { param($d) -not $d.nguon_so -or -not (@($d.nguon_so -split ' \+ ' | Where-Object { $_ -notmatch '^(OCR|thời gian ký số|ngày sao y)$' }).Count) }
  foreach ($cq0 in 'CA', 'THUE', 'BHXH', 'HQ') {
    $refs = @($refLN | Where-Object { $_.cq -eq $cq0 }); if (-not $refs.Count) { continue }
    $lns = @($docs | Where-Object { $_.loai -eq 'LN_TRA_LOI' -and $_.co_quan -eq $cq0 -and -not $_.khac_dn })
    $daDung = @{}
    # 1) khớp theo số, rồi theo ngày
    $ghep = @{}
    foreach ($d in $lns) {
      $r0 = $null
      if ($d.so) { $r0 = $refs | Where-Object { $_.so -eq "$($d.so)" } | Select-Object -First 1 }
      if (-not $r0 -and $d.ngay) { $cung = @($refs | Where-Object { $_.ngay -eq $d.ngay }); if ($cung.Count -eq 1) { $r0 = $cung[0] } }
      if ($r0 -and -not $daDung.ContainsKey($r0.so)) { $ghep[$d.file] = $r0; $daDung[$r0.so] = 1 }
    }
    # 2) còn đúng 1 văn bản chưa ghép và 1 số chưa dùng → ghép (kể cả khi OCR đọc ra số khác: số gõ trong tờ trình đáng tin hơn)
    $conDoc = @($lns | Where-Object { -not $ghep.ContainsKey($_.file) }); $conRef = @($refs | Where-Object { -not $daDung.ContainsKey($_.so) })
    $gan = { param($a, $b) $x1 = ConvertTo-Date $a; $x2 = ConvertTo-Date $b; (-not $x1) -or ($x2 -and [Math]::Abs(($x1 - $x2).TotalDays) -le 3) }
    if ($conDoc.Count -eq 1 -and $conRef.Count -eq 1 -and ((-not $conDoc[0].so) -or (& $nguonYeu $conDoc[0])) -and (& $gan $conDoc[0].ngay $conRef[0].ngay)) { $ghep[$conDoc[0].file] = $conRef[0] }
    foreach ($d in $lns) {
      if (-not $ghep.ContainsKey($d.file)) { continue }
      $r0 = $ghep[$d.file]; $yeu = & $nguonYeu $d; $doi = @()
      if (-not $d.so -or ($yeu -and "$($d.so)" -ne $r0.so -and (& $gan $d.ngay $r0.ngay))) { if ($d.so -and "$($d.so)" -ne $r0.so) { [void]$canhBao.Add("$($d.file): số đọc bằng OCR ($($d.so)) khác số ghi trong tờ trình/phiếu trình ($($r0.so)/$($r0.kh)) – đã dùng số trong tờ trình") }; $d.so = $r0.so; $d.ky_hieu = $r0.kh; $doi += 'số' }
      if (-not $d.ky_hieu) { $d.ky_hieu = $r0.kh }
      if (-not $d.ngay -or ($yeu -and $d.ngay -ne $r0.ngay -and "$($d.so)" -eq $r0.so)) { $d.ngay = $r0.ngay; $doi += 'ngày' }
      if ($doi.Count) {
        $d.so_ky_hieu = "$($d.so)/$($d.ky_hieu)"; $d.can_nhap = (-not $d.so -or -not $d.ngay)
        $d.nguon_so = ((@(@($d.nguon_so -split ' \+ ' | Where-Object { $_ -and -not ($yeu -and $_ -match '^(OCR|thời gian ký số|ngày sao y)$') }) + 'theo tờ trình/phiếu trình') | Where-Object { $_ }) -join ' + ')
      }
    }
  }
  # ---- 1d. Văn bản LN/đôn đốc vẫn thiếu ngày: ngày ghi trong tên tệp ("cahn 17.9 (an phú)", "THUẾ 11.9 Công ty S2") – năm theo CV xin ý kiến ----
  if ($mocXin) {
    foreach ($d in @($docs | Where-Object { $_.loai -in 'LN_TRA_LOI', 'DON_DOC_LN' -and -not $_.ngay })) {
      $ten0 = ([IO.Path]::GetFileNameWithoutExtension(($d.file -split '\\')[-1])) -replace '^\s*\d{1,2}\s*[\.\)]\s*', ''
      $md = [regex]::Matches($ten0, '(?<![\d\.])(\d{1,2})\.(\d{1,2})(?![\d\.])')
      if ($md.Count -ne 1) { continue }
      $dd0 = [int]$md[0].Groups[1].Value; $mm0 = [int]$md[0].Groups[2].Value; if ($dd0 -lt 1 -or $dd0 -gt 31 -or $mm0 -lt 1 -or $mm0 -gt 12) { continue }
      $nam = $mocXin.Year; if ($mm0 -lt $mocXin.Month - 1) { $nam++ }   # trả lời luôn sau CV xin ý kiến
      $d.ngay = '{0:00}/{1:00}/{2}' -f $dd0, $mm0, $nam; $d.can_nhap = (-not $d.so)
      $d.nguon_so = ((@($d.nguon_so, 'ngày theo tên file') | Where-Object { $_ }) -join ' + ')
    }
  }
  function TextOf($loai, [switch]$Docx) { ($docs | Where-Object { $_.loai -eq $loai -and ($Docx -or $_.ext -in '.pdf', '.jpg', '.jpeg', '.png') -and $_.text } | Select-Object -First 1).text }
  $tXin = ($docs | Where-Object { $_.loai -eq 'XIN_Y_KIEN_LN' -and $_.text } | Sort-Object { $d0 = ConvertTo-Date $_.ngay; if ($d0) { $d0 } else { [datetime]::MaxValue } } | Select-Object -First 1).text   # CV xin LN ĐẦU TIÊN
  if (-not $tXin) { $tXin = ($docs | Where-Object { $_.loai -eq 'DU_THAO' -and $_.file -match 'xin' } | Select-Object -First 1).text }
  $tTtr = TextOf 'TRINH_UB'; $tKq = TextOf 'KQ_UB'; $tTra = TextOf 'TRA_HO_SO'
  $tBs = TextOf 'YC_BO_SUNG'   # CV đề nghị DN bổ sung cũng nêu "nhận được văn bản số …" và "đối với 01 nhân sự …: bà X, chức vụ"
  $tPt = ($docs | Where-Object { $_.loai -eq 'PHIEU_TRINH' } | Sort-Object { ConvertTo-Date $_.ngay } | Select-Object -First 1).text
  $nguonDN = @($tXin, $tTtr, $tKq, $tTra, $tPt, $tBs) | Where-Object { $_ }   # CV bổ sung xếp cuối: chỉ bù chỗ còn thiếu
  $all = ($nguonDN -join "`n")

  # ---- 2. Thông tin doanh nghiệp ----
  $soVbDn = ''; $ngayVbDn = ''; $ngayNhan = ''; $dn = ''; $dc = ''; $mst = ''
  foreach ($t in $nguonDN) { if (-not $soVbDn -and $t -match '(?:văn bản|[Cc]ông văn|CV|[Tt]ờ trình|[Đđ]ơn)(?:\s+đề nghị)?(?:\s+của\s+(?:doanh nghiệp|Công ty))?\s+(?:số|số hiệu)\s*:?\s*([0-9][^\s,;]*?)[\.]?\s+(?:đề\s+)?(?:ký\s+)?ngày\s+(\d{1,2}/\d{1,2}/\d{4})') { $soVbDn = $Matches[1]; $ngayVbDn = Format-Ngay $Matches[2] } }
  # Văn bản của DN ghi ngày bằng chữ: "số 12/2026/CV-ECO ngày 23 tháng 10 năm 2026"
  if (-not $soVbDn) { foreach ($t in $nguonDN) { if ($t -match '(?:văn bản|[Cc]ông văn)\s+số\s+([0-9][^\s,;]*?)[\.]?\s+ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})') { $soVbDn = $Matches[1]; $ngayVbDn = '{0:00}/{1:00}/{2}' -f [int]$Matches[2], [int]$Matches[3], $Matches[4]; break } } }
  # Ngày nhận: ưu tiên dấu nhận trên phiếu trình "Số: 4318 (nhận hồ sơ ngày 26/10/2026)", sau đó mới tới câu "Ngày …, Sở Ngoại vụ nhận được …"
  if ($all -match '(?:nhận|tiếp nhận)\s+hồ sơ(?:\s+ngày)?\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})|\((?:đã\s+)?(?:nhận|tiếp nhận)(?:\s+ngày)?\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})\)|[Nn]gày nhận(?: hồ sơ)?\s*:\s*(\d{1,2}/\d{1,2}/\d{4})') { $ngayNhan = Format-Ngay (($Matches[1], $Matches[2], $Matches[3]) | Where-Object { $_ } | Select-Object -First 1) }
  elseif ($all -match 'Ngày\s+(\d{1,2}/\d{1,2}/\d{4}),\s*(?:Văn phòng|Sở Ngoại vụ)[^\n]{0,60}?nhận được') { $ngayNhan = Format-Ngay $Matches[1] }
  if ($all -match 'Tên doanh nghiệp:\s*([^\n]+?)\.?\s*(?:\n|$)') { $dn = $Matches[1].Trim() }
  if (-not $dn -and $all -match 'của\s+(Công ty[^\(\n,/]{3,120}?|Tổng Công ty[^\(\n,/]{3,120}?)\s*(?:\(|đề nghị|về việc|v/v|,|\.|;|thuộc|có|là|kèm)') { $dn = $Matches[1].Trim() }   # dừng trước "/" (ô bảng), "về việc…", "thuộc thẩm quyền…"
  $dn = (($dn -split '\s/\s')[0] -replace '\s+', ' ').Trim().TrimEnd('.')
  # Cắt phần mô tả dính sau tên ("… Công ty A hoạt động theo Giấy chứng nhận đăng ký doanh nghiệp số …", "…, mã số thuế …", "… có trụ sở tại …")
  $dn = ($dn -split '(?i)\s*[,;(\-–]?\s*(?:hoạt động theo|theo Giấy chứng nhận|Giấy chứng nhận|GCN\s*ĐKDN|mã số (?:doanh nghiệp|thuế|DN)|\bMST\b|do Sở Kế hoạch|có trụ sở|trụ sở chính|địa chỉ\s*:)')[0].Trim().TrimEnd('.', ',', ';', '-', '–').Trim()
  if ($all -match 'doanh nghiệp số\s*(\d{10,13})') { $mst = $Matches[1] }
  if ($all -match '(?:Địa chỉ trụ sở chính|Trụ sở chính)(?:\s*\([^)\n]{0,60}\))?\s*:\s*([\s\S]{5,220}?)(?:\.\s*\n|\n\s*-|\s*Người đại diện|\n\s*\d\.|\n\s*Hoạt động)') { $dc = (($Matches[1] -replace '\s+', ' ').Trim()).TrimEnd('.') }
  $dc = (($dc -split '\s/\s')[0] -replace '\s+', ' ').Trim().TrimEnd('.', ';', ',')   # dừng trước "/" (ngắt dòng trong ô bảng)
  # Ghi đè nhập tay cho cả hồ sơ (ghi_tay.json, khóa "_ho_so": doanh_nghiep / ten_viet_tat / can_bo) – người dùng sửa trong ứng dụng khi hệ thống đoán sai
  $tayHS = $tay['_ho_so']; $nhapTayHS = @(); $canBo = ''
  if ($tayHS) {
    if ([string]$tayHS.doanh_nghiep) { $dn = ([string]$tayHS.doanh_nghiep).Trim(); $nhapTayHS += 'doanh_nghiep' }
    if ([string]$tayHS.can_bo) { $canBo = ([string]$tayHS.can_bo).Trim(); $nhapTayHS += 'can_bo' }
  }
  $vt = if ($tayHS -and [string]$tayHS.ten_viet_tat) { $nhapTayHS += 'ten_viet_tat'; ([string]$tayHS.ten_viet_tat).Trim() } else { Get-VietTat $dn }
  $loi = Remove-Dau $vt
  $loiKey = (@($loi -split ' ' | Where-Object { $_ }) | Select-Object -Last 2) -join ' '
  # Văn bản của DN khác (vd phiếu trình chép từ hồ sơ khác)
  # So khớp trên chuỗi chỉ còn chữ/số (bỏ dấu, khoảng trắng, gạch nối, dấu "/" của ô bảng) để "DH - Công ty K" vẫn khớp "DH-Công ty K"
  $loiKeyAs = $loiKey -replace '[^a-z0-9]', ''
  # Ngoài 2 từ cuối, chấp nhận từ cuối (≥3 ký tự) – vd "Công ty T2": phiếu trình ghi "Tư vấn Dịch vụ Công ty S2" vẫn là cùng DN
  $loiCuoi = ((@($loi -split ' ' | Where-Object { $_ }) | Select-Object -Last 1) -replace '[^a-z0-9]', '')
  if ($loiKeyAs) { foreach ($d in $docs) { $ascD = ((Remove-Dau $d.text) -replace '[^a-z0-9]', ''); if ($d.text -and $d.loai -in 'PHIEU_TRINH', 'DU_THAO' -and $ascD -notlike "*$loiKeyAs*" -and -not ($loiCuoi.Length -ge 3 -and $ascD -like "*$loiCuoi*") -and $d.text -match 'Công ty') { $d.khac_dn = $true; [void]$canhBao.Add("$($d.file): không thấy tên '$vt' – có thể là văn bản của doanh nghiệp khác, đã bỏ qua") } } }
  foreach ($t in ($docs | Where-Object { $_.co_quan -in 'Phòng', 'Sở NV' -or $_.loai -in 'PHIEU_TRINH', 'XIN_Y_KIEN_LN', 'TRINH_UB' }).text) {
    # "Lưu: VT, LS(An)" / "LSLT(An)" / "LSAn" (không ngoặc, viết liền) → tên cán bộ; quy kiểu bỏ dấu cũ "Thuỷ" về "Thủy" cho khớp Excel
    if (-not $canBo -and $t -match 'L(?:SLT|S)\s*\(\s*(\p{L}+)\s*\)') { $canBo = $Matches[1] }
    elseif (-not $canBo -and $t -match '(?:Lưu|LƯU)\s*:?[^\n|]{0,40}?\bL(?:SLT|S)\s*[-–:]?\s*(\p{Lu}[\p{Ll}\p{Mn}]+)') { $canBo = $Matches[1] }
  }
  if ($canBo) {
    $canBo = $canBo.Normalize([Text.NormalizationForm]::FormC)
    foreach ($p in @(@('uỷ', 'ủy'), @('uý', 'úy'), @('uỳ', 'ùy'), @('uỹ', 'ũy'), @('uỵ', 'ụy'), @('oà', 'òa'), @('oá', 'óa'), @('oả', 'ỏa'), @('oã', 'õa'), @('oạ', 'ọa'))) { $canBo = $canBo.Replace($p[0], $p[1]) }
  }

  # ---- 3. Doanh nhân + ngày sinh theo từng nguồn ----
  $nguoi = [ordered]@{}
  function AddNguoi($ten, $cv, $ns, $src) {
    $ten = (($ten -replace '\s+', ' ').Trim().TrimEnd(',', '.', ';') -replace '^(?:Ông|Bà|ông|bà)\s+', ''); if ($ten.Length -lt 5 -or $ten.Length -gt 40) { return }
    if ($cv -match '^\s*(?:sinh ngày|ngày sinh|năm sinh)') { $cv = '' }
    $k = Remove-Dau $ten
    if (-not $nguoi.Contains($k)) { $nguoi[$k] = [pscustomobject]@{ ho_ten = $ten; chuc_vu = ''; ngay_sinh = [ordered]@{} } }
    if ($cv -and -not $nguoi[$k].chuc_vu) { $nguoi[$k].chuc_vu = ($cv -replace '\s+', ' ').Trim().TrimEnd('.', ';', ',') }
    if ($ns) { $nguoi[$k].ngay_sinh[$src] = Format-Ngay $ns }
  }
  $reNS = '(\d{1,2}/\d{1,2}/\d{4})'
  if ($tKq) { foreach ($mm in [regex]::Matches($tKq, "Họ và tên:\s*([^\n]+?)\s*\n[\s\S]{0,40}?Ngày tháng năm sinh:\s*$reNS[\s\S]{0,300}?Chức vụ:\s*([^\n]+)")) { AddNguoi $mm.Groups[1].Value $mm.Groups[3].Value $mm.Groups[2].Value 'KQ UB' } }
  if ($tTtr) { foreach ($mm in [regex]::Matches($tTtr, "(?:Ông|Bà)\s+([^,;\n]{3,40}?),\s*sinh ngày\s*$reNS;\s*Chức (?:danh|vụ):\s*([^;\n]+)")) { AddNguoi $mm.Groups[1].Value $mm.Groups[3].Value $mm.Groups[2].Value 'Tờ trình' } }
  # Văn bản UB / tờ trình mẫu mới (luôn đọc, cộng dồn): "Bà Trần Thị B, sinh ngày: 01/01/1980; Hộ chiếu số …; Chức danh: Chủ tịch Hội đồng quản trị."
  foreach ($src in @(@('KQ UB', $tKq), @('Tờ trình', $tTtr))) {
    if (-not $src[1]) { continue }
    foreach ($mm in [regex]::Matches($src[1], "(?:Ông|Bà)\s+(\p{Lu}\p{Ll}*(?:\s+\p{Lu}\p{Ll}*){1,5})\s*,\s*sinh ngày\s*:?\s*$reNS[^\n]{0,260}?Chức (?:danh|vụ)\s*:\s*([^;\n]+?)\s*(?:[;\.]|\n|$)")) { AddNguoi $mm.Groups[1].Value $mm.Groups[3].Value $mm.Groups[2].Value $src[0] }
  }
  $xinDocx = ($docs | Where-Object { $_.ext -eq '.docx' -and $_.file -match 'xin' } | Select-Object -First 1).text
  foreach ($t in @($xinDocx, $tPt) | Where-Object { $_ }) { foreach ($mm in [regex]::Matches($t, "\|\s*\d+\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*$reNS")) { AddNguoi $mm.Groups[1].Value $mm.Groups[2].Value $mm.Groups[3].Value 'CV xin LN' } }
  if ($nguoi.Count -eq 0 -and $tPt) { foreach ($mm in [regex]::Matches($tPt, "(?:Ông|Bà)\s+([^,;\n/]{3,40}?)\s*/?\s*Sinh ngày:?\s*$reNS")) { AddNguoi $mm.Groups[1].Value '' $mm.Groups[2].Value 'Phiếu trình' } }
  # Phiếu trình dạng gạch đầu dòng (ô bảng ngắt dòng bằng "/"): "- Họ và tên: X. / - Ngày sinh: dd/mm/yyyy. / - Số định danh… / - Chức vụ: Y."
  # Tổng quát: cắt theo từng "Họ và tên:" rồi tìm ngày sinh ("Ngày sinh" / "Ngày tháng năm sinh" / "Sinh ngày") và chức vụ trong đoạn đó (không phụ thuộc thứ tự dòng, nhiều người vẫn đúng)
  foreach ($src in @(@('Phiếu trình', $tPt), @('văn bản', $all))) {
    if ($nguoi.Count -gt 0 -or -not $src[1]) { continue }
    $parts = [regex]::Split($src[1], 'Họ (?:và )?tên\s*:')   # "Họ và tên:" hoặc "Họ tên:" (tờ trình, văn bản UB, phiếu trình mẫu mới)
    for ($pi = 1; $pi -lt $parts.Count; $pi++) {
      $seg = $parts[$pi]; if ($seg.Length -gt 900) { $seg = $seg.Substring(0, 900) }
      $ten = ($seg -split '[\n/\|]|\s[-–]\s|\.\s|\sNgày|\sSinh')[0]
      $ns = ''; if ($seg -match "(?:Ngày(?: tháng năm)? sinh|Sinh ngày|Năm sinh)\s*:?\s*$reNS") { $ns = $Matches[1] }
      $cv = ''; if ($seg -match 'Chức (?:vụ|danh)\s*:\s*([^\n/\|;]+)') { $cv = $Matches[1] }
      if ($ns) { AddNguoi $ten $cv $ns $src[0] }
    }
  }
  # Phiếu trình dạng đánh số: "(1) Nguyễn Văn A, sinh ngày 01/01/1980; / - Chức danh: Phó Tổng Giám đốc; / - Số định danh…"
  if ($nguoi.Count -eq 0 -and $tPt) {
    foreach ($mm in [regex]::Matches($tPt, "(?:^|[\s/])\(?\d{1,2}[\)\.]\s*(?:(?:Ông|Bà|ông|bà)\s+)?(\p{Lu}\p{L}+(?:\s+\p{Lu}\p{L}+){1,5})\s*,\s*(?:sinh ngày|ngày sinh)\s*:?\s*$reNS([\s\S]{0,300})")) {
      $duoi = ($mm.Groups[3].Value -split '\(\d{1,2}\)')[0]; $cv = ''
      if ($duoi -match 'Chức (?:danh|vụ)\s*:\s*([^;/\n|]+)') { $cv = $Matches[1] }
      AddNguoi $mm.Groups[1].Value $cv $mm.Groups[2].Value 'Phiếu trình'
    }
  }
  # Bảng trong PDF (Word trích mỗi ô một dòng) – chỉ dùng khi các cách trên không ra ai (ô xuống dòng có thể cắt đôi tên): "TT ⏎ Họ tên ⏎ Chức vụ ⏎ Năm sinh ⏎ … ⏎ 1 ⏎ Lê Văn C ⏎ Trưởng phòng Tài chính ⏎ 01/01/1980"
  if ($nguoi.Count -eq 0) {
    foreach ($src in @(@('CV xin LN', $tXin), @('Tờ trình', $tTtr), @('Phiếu trình', $tPt))) {
      if (-not $src[1]) { continue }
      foreach ($mm in [regex]::Matches(($src[1] -replace "[ \t]+\n", "`n"), "(?m)^\s*(\d{1,2})[\.\)]?\s*\n\s*(\p{Lu}\p{Ll}*(?:\s+\p{Lu}\p{Ll}*){1,5})\s*\n\s*((?:[^\n\d]{2,60}\n\s*){0,2}[^\n\d]{2,60}?)\s*\n\s*$reNS")) { AddNguoi $mm.Groups[2].Value ($mm.Groups[3].Value -replace '\s*\n\s*', ' ') $mm.Groups[4].Value $src[0] }
      if ($nguoi.Count) { break }
    }
  }
  # Dự phòng cuối: "đối với bà Nguyễn Văn A, Giám đốc và ông Trần Văn B, Trưởng phòng…" (văn bản trả hồ sơ / xin ý kiến không có bảng)
  if ($nguoi.Count -eq 0) { foreach ($mm in [regex]::Matches($all, "(?:ông|bà|Ông|Bà)\s+(\p{Lu}[\p{L}]+(?:\s\p{Lu}[\p{L}]+){1,5}),\s*([^;,\n.]{3,60}?)(?=\s+và\s+(?:ông|bà)|[;.\n]|\s+\(|\s+thuộc)")) { AddNguoi $mm.Groups[1].Value $mm.Groups[2].Value '' 'văn bản' } }
  # Ngày sinh trong văn bản liên ngành (để phát hiện lệch)
  foreach ($d in $docs | Where-Object { $_.loai -eq 'LN_TRA_LOI' -and $_.text }) {
    foreach ($k in @($nguoi.Keys)) {
      $ten = [regex]::Escape($nguoi[$k].ho_ten) -replace '\\ ', '\s+'
      if ($d.text -match "$ten[,\s]{1,3}[Ss]inh ngày:?\s*$reNS") { $nguoi[$k].ngay_sinh[$d.file] = Format-Ngay $Matches[1] }
    }
  }
  foreach ($k in $nguoi.Keys) { $vals = @($nguoi[$k].ngay_sinh.Values | Sort-Object -Unique); if ($vals.Count -gt 1) { [void]$canhBao.Add("Ngày sinh $($nguoi[$k].ho_ten) không thống nhất: " + (($nguoi[$k].ngay_sinh.GetEnumerator() | ForEach-Object { "$($_.Value) ($($_.Key))" }) -join '; ')) } }
  $soLuong = $nguoi.Count
  if ($all -match 'đối với\s+(\d{1,2})\s+doanh nhân') { $sl = [int]$Matches[1]; if ($soLuong -eq 0) { $soLuong = $sl } elseif ($sl -ne $soLuong) { [void]$canhBao.Add("Số doanh nhân nhận diện được ($soLuong) khác số ghi trong văn bản ($sl)") } }
  if ($nguoi.Count -eq 0) { [void]$canhBao.Add('Chưa nhận diện được danh sách doanh nhân – cần nhập tay cột G') }

  # ---- 3b. QUY TẮC: quét kỹ TOÀN BỘ phiếu trình của mọi hồ sơ để lấy đủ thông tin (địa chỉ, GCN đăng ký DN, lĩnh vực, người đại diện, CCCD/hộ chiếu…) ----
  # Phiếu trình là bản đã được cán bộ đối chiếu nên dùng để điền chỗ còn thiếu; không ghi đè thông tin đã đọc từ văn bản khác.
  $pt = [ordered]@{}
  $ptAll = (($docs | Where-Object { $_.loai -eq 'PHIEU_TRINH' -and $_.text -and -not $_.khac_dn } | ForEach-Object { $_.text }) -join "`n")
  if ($ptAll) {
    $ptn = (($ptAll -replace '\s/\s+', "`n") -replace '\|', "`n")      # ô bảng ngắt dòng bằng "/" → mỗi mục một dòng
    function PtGet([string]$re) { if ($ptn -match $re) { return (($Matches[1] -replace '\s+', ' ').Trim().TrimEnd('.', ';', ',').Trim()) }; '' }
    $pt.so_den           = PtGet 'Số\s*:\s*(\d+)\s*\((?:đã\s+)?(?:nhận|tiếp nhận)'
    $pt.van_de_trinh     = PtGet 'Vấn đề trình\s*\n\s*([^\n]+)'
    $pt.van_ban_trinh    = PtGet 'Văn bản trình\s*\n\s*([^\n]+)'
    $pt.ten_doanh_nghiep = PtGet 'Tên doanh nghiệp\s*:\s*([^\n]+)'
    $pt.gcn_dkdn         = PtGet '(Giấy chứng nhận đăng ký (?:doanh nghiệp|kinh doanh)[^\n]*?\d{8,13}[^\n]*)'
    $pt.ma_so_dn         = PtGet 'Giấy chứng nhận đăng ký (?:doanh nghiệp|kinh doanh)[^\n]*?số\s*:?\s*(\d{8,13})'
    $pt.dia_chi          = PtGet 'Địa chỉ(?: trụ sở(?: chính)?)?(?:\s*\([^)\n]{0,60}\))?\s*:\s*([^\n]+)'
    $pt.linh_vuc         = PtGet 'Lĩnh vực[^:\n]{0,40}:\s*([^\n]+)'
    $pt.nguoi_dai_dien   = PtGet 'Người đại diện theo pháp luật\s*:\s*([^\n]+)'
    # Thông tin bổ sung từng doanh nhân (danh sách gạch đầu dòng): CCCD / hộ chiếu
    $pp = [regex]::Split($ptn, 'Họ (?:và )?tên\s*:')
    for ($pi = 1; $pi -lt $pp.Count; $pi++) {
      $seg = $pp[$pi]; if ($seg.Length -gt 900) { $seg = $seg.Substring(0, 900) }
      $k = Remove-Dau ((($seg -split '[\n]')[0]) -replace '\s+', ' ').Trim().TrimEnd('.', ',', ';')
      if (-not $nguoi.Contains($k)) { continue }
      $o = $nguoi[$k]
      if ($seg -match '(?:Số định danh(?: cá nhân)?|CCCD|Căn cước(?: công dân)?)(?:\s*số)?\s*:?\s*(\d{9,12})') { $o | Add-Member -NotePropertyName cccd -NotePropertyValue $Matches[1] -Force }
      if ($seg -match 'Hộ chiếu(?: số)?\s*:?\s*([A-Z]\d{6,9})') { $o | Add-Member -NotePropertyName ho_chieu -NotePropertyValue $Matches[1] -Force }
      if ($seg -match 'cấp ngày\s*(\d{1,2}/\d{1,2}/\d{4})') { $o | Add-Member -NotePropertyName ho_chieu_cap -NotePropertyValue (Format-Ngay $Matches[1]) -Force }
      if ($seg -match '(?:có giá trị đến|ngày hết hạn|hết hạn)\s*(?:ngày)?\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})') { $o | Add-Member -NotePropertyName ho_chieu_het_han -NotePropertyValue (Format-Ngay $Matches[1]) -Force }
    }
    # Điền chỗ còn thiếu
    if (-not $dc  -and $pt.dia_chi)         { $dc = $pt.dia_chi }
    if (-not $mst -and $pt.ma_so_dn)        { $mst = $pt.ma_so_dn }
    if (-not $dn  -and $pt.ten_doanh_nghiep) { $dn = $pt.ten_doanh_nghiep }
    foreach ($key in @($pt.Keys)) { if (-not $pt[$key]) { $pt.Remove($key) } }
    $thieu = @(); if (-not $dc) { $thieu += 'địa chỉ' }; if (-not $mst) { $thieu += 'mã số DN' }
    if ($thieu.Count) { [void]$canhBao.Add("Đã quét kỹ phiếu trình nhưng vẫn thiếu: " + ($thieu -join ', ') + " – cần nhập tay") }
  }

  # ---- 4. Liên ngành ----
  $hoi = @(); $kg = if ($tXin) { $tXin.Substring(0, [Math]::Min(1500, $tXin.Length)) } else { '' }
  if ($kg -match 'Công an') { $hoi += 'CA' }; if ($kg -match '[Tt]huế') { $hoi += 'THUE' }; if ($kg -match 'Bảo hiểm xã hội') { $hoi += 'BHXH' }; if ($kg -match 'Hải quan') { $hoi += 'HQ' }
  if (-not $hoi.Count) { $hoi = @('CA', 'THUE', 'BHXH', 'HQ') }
  $ln = [ordered]@{}; $luuY = @()
  foreach ($cq in 'CA', 'THUE', 'BHXH', 'HQ') {
    # Có văn bản LN là tính "đã trả lời" (kể cả chưa đọc được ngày); văn bản có ngày muộn nhất làm đại diện
    $ds = @($docs | Where-Object { $_.loai -eq 'LN_TRA_LOI' -and $_.co_quan -eq $cq -and -not $_.khac_dn } | Sort-Object { $d0 = ConvertTo-Date $_.ngay; if ($d0) { $d0 } else { [datetime]::MinValue } })
    if (-not $ds.Count) { $ln[$cq] = [pscustomobject]@{ so = ''; ngay = ''; ket_luan = $(if ($hoi -contains $cq) { 'chưa có' } else { 'không hỏi' }) }; continue }
    $last = $ds[-1]
    $kl = @([regex]::Matches($last.text, '(?:không|chưa)\s[^.\n]{0,70}?(?:nợ|vi phạm)[^.\n;]{0,40}|[Nn]ợ phí hải quan[^.\n;]{0,40}|[Vv]i phạm[^.\n;]{0,30}:\s*không|xử phạt[^.\n]{0,80}|còn thiếu[^.\n]{0,60}') | Select-Object -First 3 | ForEach-Object { ($_.Value -replace '\s+', ' ').Trim() })
    foreach ($d in $ds) { $d.noi_dung = (($kl -join '; ')) }
    $ln[$cq] = [pscustomobject]@{ so = $last.so_ky_hieu; ngay = $last.ngay; ket_luan = ($kl -join '; '); co = $true }
    $txtAll = ($ds | ForEach-Object { $_.text }) -join "`n"
    if ($txtAll -match '[Nn]ợ phí hải quan:?\s*([\d\.]+)\s*đ' -and $Matches[1] -ne '0') { $luuY += "HQ nợ phí $($Matches[1])đ" }
    if ($txtAll -match '(?<!không có\s)(?<!Không có\s)Quyết định xử phạt[^.\n]{0,60}' -and $txtAll -notmatch 'không có Quyết định xử phạt') { $luuY += 'Thuế/HQ nêu quyết định xử phạt – kiểm tra'; [void]$canhBao.Add("$($cq): văn bản có nêu quyết định xử phạt – cần đánh giá") }
    foreach ($d in $ds) { $cnt = [regex]::Matches($d.text, '(?m)^\s*\d+\.\s*(?:Công ty|Tổng Công ty)').Count; if ($cnt -ge 2) { [void]$canhBao.Add("$($d.file): văn bản trả lời chung $cnt doanh nghiệp – kiểm tra các hồ sơ khác được nhắc tới") } }
  }

  # ---- 4b. Hồ sơ có cấu trúc: ngày cuối tên folder DN (vd "1. Công ty N2 16.1") ----
  # Tháng 1–8/2026: là ngày có KẾT QUẢ UB; từ tháng 9 (có thư mục "Tờ trình UB"): là ngày TỜ TRÌNH. Khớp một trong hai là đúng.
  $ngFolder = if (Test-HoSoCoCauTruc $dir.FullName) { Get-NgayTuTenFolder $dir.Name } else { $null }
  if ($ngFolder) {
    $kqDocs = @($docs | Where-Object { $_.loai -eq 'KQ_UB' -and -not $_.khac_dn }); $ttrDocs = @($docs | Where-Object { $_.loai -eq 'TRINH_UB' -and -not $_.khac_dn })
    $khopNgay = { param($d) $dt = ConvertTo-Date $d.ngay; $dt -and $dt.Day -eq $ngFolder.d -and $dt.Month -eq $ngFolder.m }
    # Văn bản UB chưa đọc được ngày, hồ sơ không có tờ trình → ngày theo tên folder
    if (-not $ttrDocs.Count) {
      foreach ($d in $kqDocs | Where-Object { -not (ConvertTo-Date $_.ngay) }) {
        $moc = ($docs | ForEach-Object { ConvertTo-Date $_.ngay } | Where-Object { $_ } | Measure-Object -Maximum).Maximum
        $nam = if ($moc) { $moc.Year } else { (Get-Date).Year }
        if ($moc -and [datetime]::new($nam, $ngFolder.m, $ngFolder.d) -lt $moc.AddDays(-31)) { $nam++ }   # kết quả UB luôn sau các văn bản khác
        $d.ngay = '{0:00}/{1:00}/{2}' -f $ngFolder.d, $ngFolder.m, $nam
        $d.nguon_so = (@($d.nguon_so, 'ngày theo tên folder') | Where-Object { $_ }) -join ' + '; $d.can_nhap = (-not $d.so)
      }
    }
    $coNgay = @(@($kqDocs) + @($ttrDocs) | Where-Object { ConvertTo-Date $_.ngay })
    if ($coNgay.Count -and -not @($coNgay | Where-Object { & $khopNgay $_ }).Count) {
      [void]$canhBao.Add("Ngày ghi ở tên folder ($($ngFolder.d).$($ngFolder.m)) không khớp ngày tờ trình / văn bản UB (" + (($coNgay | ForEach-Object { $_.ngay }) -join ', ') + ') – kiểm tra lại')
    }
  }
  # ---- 5. Sự kiện + trạng thái ----
  $hieuLuc = @($docs | Where-Object { $_.loai -notin 'TRUNG', 'DU_THAO', 'KHAC', 'DN_HO_SO_KEM' -and -not $_.khac_dn })
  foreach ($d in $hieuLuc | Where-Object { $_.can_nhap }) { [void]$canhBao.Add("$($d.file): chưa đọc được " + $(if (-not $d.so) { 'số' } else { '' }) + $(if (-not $d.so -and -not $d.ngay) { ' và ' } else { '' }) + $(if (-not $d.ngay) { 'ngày' } else { '' }) + ' – xem ảnh đầu trang và nhập tay') }
  foreach ($d in $docs | Where-Object { $_.loai -eq 'TRUNG' }) { [void]$canhBao.Add("$($d.file): trùng nội dung với file khác – đã bỏ qua") }
  $ev = @($hieuLuc | Where-Object { ConvertTo-Date $_.ngay } | Sort-Object { ConvertTo-Date $_.ngay })
  $xin = @($ev | Where-Object loai -eq 'XIN_Y_KIEN_LN'); $tra = @($ev | Where-Object loai -eq 'TRA_HO_SO')
  $bs = @($ev | Where-Object loai -eq 'YC_BO_SUNG'); $dd = @($ev | Where-Object loai -eq 'DON_DOC_LN')
  $ttr = @($ev | Where-Object loai -eq 'TRINH_UB'); $kq = @($ev | Where-Object loai -eq 'KQ_UB')
  $daTraLoi = @($hoi | Where-Object { $ln[$_].co }); $du = ($daTraLoi.Count -eq $hoi.Count)
  $ngayDu = $null
  if ($du -and -not @($daTraLoi | Where-Object { -not (ConvertTo-Date $ln[$_].ngay) }).Count) { $ngayDu = ($daTraLoi | ForEach-Object { ConvertTo-Date $ln[$_].ngay } | Measure-Object -Maximum).Maximum }
  # Phiếu trình kết quả = phiếu trình sau ý kiến LN cuối cùng (nếu thiếu ngày LN: sau ngày LN muộn nhất đã biết)
  $mocLN = if ($ngayDu) { $ngayDu } else { ($hoi | ForEach-Object { ConvertTo-Date $ln[$_].ngay } | Where-Object { $_ } | Measure-Object -Maximum).Maximum }
  $ptKq = @($ev | Where-Object { $_.loai -eq 'PHIEU_TRINH' -and $du -and $mocLN -and (ConvertTo-Date $_.ngay) -ge $mocLN.AddDays(-1) })
  # Trả hồ sơ MỘT PHẦN: thư mục/tên tệp nêu riêng một người, vd "Không đủ điều kiện (ông Hải)", hoặc hồ sơ vẫn xin ý kiến/có kết quả UB
  $tenTraMotPhan = { param($d) $s = "$($d.thu_muc) $($d.file)"; if ($s -match '\(\s*((?:ông|bà|Ông|Bà)\s+[^)]+?)\s*\)') { $Matches[1] } else { '' } }
  $traMotPhan = @($tra | Where-Object { $_.text -match 'đang (được )?xin ý kiến|tiếp tục' -or $xin.Count -or $kq.Count -or (& $tenTraMotPhan $_) })
  $tt = 'S0_MOI'
  if ($bs.Count) { $tt = 'S1_CHO_BO_SUNG' }
  if ($xin.Count) { $tt = 'S3_XIN_Y_KIEN_LN' }; if ($dd.Count) { $tt = 'S4_DON_DOC_LN' }
  if ($du) { $tt = 'S5_DU_Y_KIEN_LN' }; if ($ptKq.Count) { $tt = 'S5b_DANG_TRINH_SO' }
  if ($ttr.Count) { $tt = 'S6_DA_TRINH_UB' }
  if ($kq.Count) { $tt = $(if ($kq[-1].text -match 'không đồng ý|không cho phép') { 'S8_UB_KHONG_CHO_PHEP' } else { 'S7_UB_CHO_PHEP' }) }
  if ($tra.Count -and -not $xin.Count -and -not $kq.Count) { $tt = 'S9_TRA_HO_SO' }
  if ($traMotPhan.Count -and $xin.Count) { foreach ($d in $traMotPhan) { [void]$canhBao.Add("$($d.file): thông báo kết quả MỘT PHẦN (hồ sơ vẫn tiếp tục) – kiểm tra số doanh nhân bị loại") } }
  if (-not $du -and ($ttr.Count -or $kq.Count)) { [void]$canhBao.Add('Trình/có kết quả khi chưa đủ ý kiến liên ngành: thiếu ' + ((@($hoi | Where-Object { -not $ln[$_].co })) -join ', ')) }
  $ocrDocs = @($hieuLuc | Where-Object { $_.nguon_so -match 'OCR' } | ForEach-Object { $_.file })
  if ($ocrDocs.Count) { [void]$canhBao.Add("Số/ngày đọc bằng OCR (nên liếc ảnh để kiểm): " + ($ocrDocs -join '; ')) }
  $suKien = @(); if ($ngayNhan) { $suKien += [pscustomobject]@{ ngay = $ngayNhan; su_kien = 'DN_DE_NGHI'; so_van_ban = $soVbDn; trang_thai_sau = 'S0_MOI' } }
  # Trạng thái SAU từng sự kiện (chạy dần theo ngày) – dùng cho nhật ký và báo cáo tuần; trạng thái cuối $tt vẫn là kết luận chính
  $st = 'S0_MOI'; $daCo = @{}
  foreach ($d in $ev) {
    switch ($d.loai) {
      'YC_BO_SUNG'    { $st = 'S1_CHO_BO_SUNG' }
      'XIN_Y_KIEN_LN' { $st = 'S3_XIN_Y_KIEN_LN' }
      'DON_DOC_LN'    { $st = 'S4_DON_DOC_LN' }
      'LN_TRA_LOI'    { if ($d.co_quan) { $daCo[$d.co_quan] = 1 }; if (-not @($hoi | Where-Object { -not $daCo.ContainsKey($_) }).Count) { $st = 'S5_DU_Y_KIEN_LN' } }
      'PHIEU_TRINH'   { if ($st -eq 'S5_DU_Y_KIEN_LN') { $st = 'S5b_DANG_TRINH_SO' } }
      'TRINH_UB'      { $st = 'S6_DA_TRINH_UB' }
      'KQ_UB'         { $st = $(if ($d.text -match 'không đồng ý|không cho phép') { 'S8_UB_KHONG_CHO_PHEP' } else { 'S7_UB_CHO_PHEP' }) }
      'TRA_HO_SO'     { if (-not $xin.Count) { $st = 'S9_TRA_HO_SO' } }
    }
    $suKien += [pscustomobject]@{ ngay = $d.ngay; su_kien = $(if ($d.loai -eq 'LN_TRA_LOI') { "LN_TRA_LOI ($($d.co_quan))" } else { $d.loai }); so_van_ban = $d.so_ky_hieu; trang_thai_sau = $st }
  }
  if ($du -and $ngayDu) { $suKien += [pscustomobject]@{ ngay = $ngayDu.ToString('dd/MM/yyyy'); su_kien = 'DU_Y_KIEN_LN'; so_van_ban = ''; trang_thai_sau = 'S5_DU_Y_KIEN_LN' } }

  # ---- 6. Giá trị cột Excel mong muốn ----
  $vb = { param($d) "$($d.so_ky_hieu) ngày $($d.ngay)" }
  $col = [ordered]@{ B = ''; C = $canBo; D = $dn; E = $dc; F = $(if ($soLuong) { "$soLuong" } else { '' }); G = (($nguoi.Values | ForEach-Object { if ($_.chuc_vu) { "$($_.ho_ten), $($_.chuc_vu)" } else { $_.ho_ten } }) -join "`n"); H = ''; I = ''; J = ''; K = ''; L = ''; M = ''; N = ''; O = ''; P = '' }
  if ($soVbDn) { $col.B = "$soVbDn ngày $ngayVbDn" + $(if ($ngayNhan) { " (nhận ngày $ngayNhan)" } else { '' }) }
  if ($tt -eq 'S9_TRA_HO_SO') { $col.H = (& $vb $tra[-1]) + ' (không đủ ĐK)' }
  elseif ($traMotPhan.Count) { $col.H = (($traMotPhan | ForEach-Object { $ten = & $tenTraMotPhan $_; (& $vb $_) + ' (không đủ ĐK' + $(if ($ten) { " – $ten" } else { ' – một phần' }) + ')' }) -join "`n") }
  if ($bs.Count) { $col.I = (($bs | ForEach-Object { (& $vb $_) + ' (đề nghị bổ sung HS)' }) -join "`n") }
  if ($xin.Count) { $col.J = (& $vb $xin[0]); if ($xin.Count -gt 1) { $col.J += "`n" + (($xin | Select-Object -Skip 1 | ForEach-Object { (& $vb $_) + ' (xin ý kiến bổ sung)' }) -join "`n") } }
  if ($dd.Count) { $i = 0; $col.K = (($dd | ForEach-Object { $i++; "Lần ${i}: " + (& $vb $_) }) -join "`n") }
  if ($ttr.Count) { $col.L = (& $vb $ttr[-1]) }
  if ($kq.Count) { $col.M = (& $vb $kq[-1]) + $(if ($tt -eq 'S8_UB_KHONG_CHO_PHEP') { ' (không đồng ý)' } else { '' }) }
  $chua = @($hoi | Where-Object { -not $ln[$_].ngay })
  $col.N = switch ($tt) {
    'S0_MOI' { "Mới tiếp nhận $ngayNhan; đang rà soát" }
    'S1_CHO_BO_SUNG' { "Chờ DN bổ sung (YC $($bs[-1].so_ky_hieu) ngày $($bs[-1].ngay))" }
    'S3_XIN_Y_KIEN_LN' { "Đang xin ý kiến LN ($($daTraLoi.Count)/$($hoi.Count)): " + ((@($(if ($daTraLoi.Count) { "đã có $($daTraLoi -join ', ')" }), $(if ($chua.Count) { "chờ $($chua -join ', ')" })) | Where-Object { $_ }) -join '; ') }
    'S4_DON_DOC_LN' { "Đã đôn đốc lần $($dd.Count) ngày $($dd[-1].ngay); chờ $($chua -join ', ')" }
    'S5_DU_Y_KIEN_LN' { "Đủ ý kiến LN ($($hoi.Count)/$($hoi.Count)); chờ trình" }
    'S5b_DANG_TRINH_SO' { "Đang trình BGĐ Sở (phiếu trình ngày $($ptKq[-1].ngay))" }
    'S6_DA_TRINH_UB' { "Đã trình UB ($($ttr[-1].so_ky_hieu) ngày $($ttr[-1].ngay)); chờ kết quả" }
    'S7_UB_CHO_PHEP' { $hl = (ConvertTo-Date $kq[-1].ngay); "Hoàn thành – UB cho phép $('{0:00}' -f $soLuong)/$('{0:00}' -f $soLuong) doanh nhân" + $(if ($hl) { "; VB hiệu lực đến $($hl.AddMonths(6).ToString('dd/MM/yyyy'))" } else { '' }) }
    'S8_UB_KHONG_CHO_PHEP' { 'Hoàn thành – UB không cho phép' }
    'S9_TRA_HO_SO' { 'Đã trả hồ sơ – không đủ ĐK' }
    default { '' }
  }
  if ($luuY.Count) { $col.N += ' | Lưu ý: ' + (($luuY | Sort-Object -Unique) -join '; ') }

  # ---- 7. Ghép dòng Excel + đề xuất thay đổi ----
  $ex = Read-ExcelRows $Excel; $dong = $null; $cachGhep = ''
  $soXin = if ($xin.Count) { $xin[0].so } else { '' }
  if ($soXin) { foreach ($r in ($ex.rows.Keys | Sort-Object)) { if ($r -ge 5 -and $ex.rows[$r]['J'] -match "(?<!\d)$soXin(?!\d)") { $dong = $r; $cachGhep = "cột J có số CV xin LN $soXin"; break } } }
  if (-not $dong -and $soVbDn) { foreach ($r in ($ex.rows.Keys | Sort-Object)) { if ($r -ge 5 -and $ex.rows[$r]['B'] -and (Remove-Dau $ex.rows[$r]['B']) -like ('*' + (Remove-Dau $soVbDn) + '*')) { $dong = $r; $cachGhep = "cột B có số VB DN $soVbDn"; break } } }
  if (-not $dong -and $loiKey -and $nguoi.Count) {
    foreach ($r in ($ex.rows.Keys | Sort-Object -Descending)) { $rr = $ex.rows[$r]; if ($r -lt 5) { continue }
      if ((Remove-Dau $rr['D']) -like "*$loiKey*") { $g = Remove-Dau $rr['G']; if (@($nguoi.Keys | Where-Object { $g -like "*$_*" }).Count) { $dong = $r; $cachGhep = 'tên DN + tên doanh nhân'; break } } }
  }
  $moi = -not $dong; if ($moi) { $dong = [Math]::Max($ex.maxRow + 1, 5); [void]$canhBao.Add("Không tìm thấy dòng Excel khớp – đề xuất thêm dòng mới $dong (kiểm tra kỹ trước khi ghi)") }
  $cur = if ($moi) { @{} } else { $ex.rows[$dong] }
  $deXuat = @(); $khacBiet = @()
  $flat = { param($s) (([string]$s) -replace "`r", '' -replace "`n", ' / ').Trim() }
  if ($moi) { $deXuat += [pscustomobject]@{ o = "A$dong"; cot = 'A'; cu = ''; moi = "$($ex.maxStt + 1)"; ly_do = 'STT dòng mới'; chon = $true } }
  foreach ($c in $col.Keys) {
    $want = [string]$col[$c]; if (-not $want) { continue }
    $have = [string]$cur[$c]
    if ((& $flat $have) -eq (& $flat $want)) { continue }
    if (-not $have) { $deXuat += [pscustomobject]@{ o = "$c$dong"; cot = $c; cu = ''; moi = $want; ly_do = 'Ô đang trống'; chon = ($c -notin 'C', 'D', 'E', 'F', 'G' -or $moi) }; continue }
    if ($c -in 'J', 'L', 'M', 'H', 'I', 'K') {
      $num = if ($want -match '^(\d+)') { $Matches[1] } else { '' }
      if ($num -and $have -match "(?<!\d)$num(?!\d)") {
        $khW = if ($want -match '^\d+/([^\s]+)') { $Matches[1] } else { '' }
        if ($khW -and $have -notmatch [regex]::Escape($khW)) { $deXuat += [pscustomobject]@{ o = "$c$dong"; cot = $c; cu = $have; moi = $want; ly_do = "Ký hiệu khác văn bản gốc ($khW)"; chon = $false }; $khacBiet += "Cột ${c}: Excel '$have', văn bản gốc '$want'" }
        elseif ($c -in 'J', 'I', 'K' -and ($want -split "`n").Count -gt ($have -split "( / |`n)").Count) { $deXuat += [pscustomobject]@{ o = "$c$dong"; cot = $c; cu = $have; moi = $want; ly_do = 'Bổ sung văn bản còn thiếu'; chon = $false } }
        continue
      }
      $deXuat += [pscustomobject]@{ o = "$c$dong"; cot = $c; cu = $have; moi = $want; ly_do = 'Khác văn bản gốc'; chon = $false }; $khacBiet += "Cột ${c}: Excel '$have', văn bản gốc '$want'"; continue
    }
    if ($c -eq 'B') { if ($ngayNhan -and $have -notmatch [regex]::Escape($ngayNhan) -and $have -notmatch 'nhận|số đến') { $deXuat += [pscustomobject]@{ o = "B$dong"; cot = 'B'; cu = $have; moi = $want; ly_do = 'Bổ sung ngày nhận'; chon = $false } }; continue }
    if ($c -eq 'N') {
      # Ghi chú đã có: nếu ghi chú mới chỉ là bản cũ + phần thêm -> thay cả ô (không lặp); nếu bản cũ đã chứa nội dung mới -> bỏ qua; còn lại -> nối thêm
      $fh = & $flat $have; $fw = & $flat $want
      if ($fh.Contains($fw)) { continue }
      if ($fw.StartsWith($fh)) { $deXuat += [pscustomobject]@{ o = "N$dong"; cot = 'N'; cu = $have; moi = $want; ly_do = 'Ghi chú đã có – thêm phần mới'; chon = $false } }
      else { $deXuat += [pscustomobject]@{ o = "N$dong"; cot = 'N'; cu = $have; moi = "$have | $want"; ly_do = 'Ghi chú đã có – nối thêm'; chon = $false } }
      continue
    }
    if ($c -eq 'F' -and $have -eq $want) { continue }
    $khacBiet += "Cột ${c}: Excel '$have', văn bản '$want'"
  }

  # Gợi ý tên folder chuẩn "<năm>, ABTC, CTY <VIẾT TẮT>" khi tên folder chưa chứa tên viết tắt (chỉ gợi ý, không tự đổi)
  $folderGoiY = ''
  if ($vt) {
    $a = (Remove-Dau $dir.Name) -replace '[^a-z0-9]', ''; $b = (Remove-Dau $vt) -replace '[^a-z0-9]', ''
    $nam = if ($dir.Name -match '^(20\d{2})') { $Matches[1] } else { (Get-Date).Year }
    if ($b -and -not $a.Contains($b)) {
      # Tên folder gợi ý: bỏ ký tự Windows không cho dùng, tối đa ~60 ký tự (cắt theo từ)
      $vtF = (($vt -replace '[\\/:*?"<>|]', ' ' -replace '\s+', ' ').Trim().TrimEnd('.', ' ', '-'))
      if ($vtF.Length -gt 45) { $vtF = ($vtF.Substring(0, 45) -replace '\s+\S*$', '').TrimEnd(',', '-', ' ') }
      $folderGoiY = "$nam, ABTC, CTY " + $vtF.ToUpper()
    }
  }
  $ketQua = [pscustomobject]@{
    folder = $dir.Name; duong_dan = $dir.FullName; nguon = 'quy tắc'; doanh_nghiep = $dn; ten_viet_tat = $vt; ma_so_dn = $mst; dia_chi = $dc
    so_vb_dn = $soVbDn; ngay_vb_dn = $ngayVbDn; ngay_nhan = $ngayNhan; can_bo = $canBo; so_luong = $soLuong; ds_ca_nhan = @($nguoi.Values)
    nhap_tay_ho_so = $nhapTayHS; folder_goi_y = $folderGoiY; thong_tin_phieu_trinh = $pt
    van_ban = @($docs | ForEach-Object { [pscustomobject]@{ file = $_.file; loai = $(if ($_.khac_dn) { 'KHAC (DN khác)' } else { $_.loai }); co_quan = $_.co_quan; so_ky_hieu = $_.so_ky_hieu; ngay = $_.ngay; nguon_so = $_.nguon_so; noi_dung = $_.noi_dung; anh = $_.anh; can_nhap = $_.can_nhap
      so = $_.so; ky_hieu = $_.ky_hieu; ten_goi_y = $(if ($_.khac_dn) { '' } else { Get-TenChuan $_.file $_.so $_.ngay $_.loai }) } })
    co_quan_duoc_hoi = $hoi; ln = $ln; don_doc = @($dd | ForEach-Object { [pscustomobject]@{ so = $_.so_ky_hieu; ngay = $_.ngay } }); bo_sung = @($bs | ForEach-Object { & $vb $_ })
    trang_thai = $tt; su_kien = $suKien; cot_excel = [pscustomobject]$col; dong_excel_hien_co = $(if ($moi) { $null } else { $dong }); dong_ghi = $dong; cach_ghep = $cachGhep
    excel_hien_co = [pscustomobject]$cur; de_xuat = $deXuat; khac_biet_voi_excel = $khacBiet; canh_bao = @($canhBao)
  }
  $json = $ketQua | ConvertTo-Json -Depth 8
  [IO.File]::WriteAllText((Join-Path $cache 'phan_tich.json'), $json, (New-Object Text.UTF8Encoding($false)))
  if ($Out) { [IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding($false))) }
  if ($Im) { return $ketQua }
  "{0}: {1} | dòng Excel {2}{3} | {4} đề xuất | {5} cảnh báo" -f $dir.Name, $tt, $dong, $(if ($moi) { ' (mới)' } else { " ($cachGhep)" }), $deXuat.Count, $canhBao.Count
}

# =====================================================================================
#  6. TRỢ GIÚP + BỘ CHIA LỆNH
# =====================================================================================
function Start-App { & (Join-Path $Script:HE_THONG 'app\server.ps1') }   # mở ứng dụng giao diện web cục bộ

function New-PhimTat {
  # Tạo phím tắt "Hồ sơ ABTC" (có icon) ở thư mục gốc và trên Desktop. Chạy lại lệnh này khi chuyển sang máy khác.
  # Dùng API IShellLinkW của Windows (WScript.Shell không lưu được đường dẫn có dấu tiếng Việt).
  param([switch]$KhongDesktop)
  $ico = Join-Path $Script:HE_THONG 'app\icon.ico'
  if (-not (Test-Path $ico)) { & (Join-Path $Script:HE_THONG 'app\tao_icon.ps1') | Out-Null }
  if (-not ('AbtcPhimTat' -as [type])) {
    Add-Type -TypeDefinition @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Runtime.InteropServices.ComTypes;
[ComImport, Guid("00021401-0000-0000-C000-000000000046")] class CShellLink {}
[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
interface IShellLinkW {
  void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder f, int c, IntPtr fd, int fl); void GetIDList(out IntPtr p); void SetIDList(IntPtr p);
  void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder s, int c); void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string s);
  void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder s, int c); void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string s);
  void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder s, int c); void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string s);
  void GetHotkey(out short h); void SetHotkey(short h); void GetShowCmd(out int s); void SetShowCmd(int s);
  void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder s, int c, out int i); void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string s, int i);
  void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string s, int r); void Resolve(IntPtr h, int f); void SetPath([MarshalAs(UnmanagedType.LPWStr)] string s);
}
public static class AbtcPhimTat {
  public static void Tao(string lnk, string target, string args, string wd, string icon, string desc, int show) {
    IShellLinkW l = (IShellLinkW)new CShellLink();
    l.SetPath(target); l.SetArguments(args); l.SetWorkingDirectory(wd); l.SetIconLocation(icon, 0); l.SetDescription(desc); l.SetShowCmd(show);
    ((IPersistFile)l).Save(lnk, true);
  }
}
"@
  }
  $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $args2 = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "' + (Join-Path $Script:HE_THONG 'app\server.ps1') + '"'
  $noi = @($Script:GOC); if (-not $KhongDesktop) { $noi += [Environment]::GetFolderPath('Desktop') }
  foreach ($d in $noi) {
    $lnk = Join-Path $d 'Hồ sơ ABTC.lnk'
    [AbtcPhimTat]::Tao($lnk, $ps, $args2, $Script:GOC, $ico, 'Ứng dụng Hồ sơ ABTC – Phòng Lãnh sự', 7)   # 7 = cửa sổ máy chủ thu nhỏ
    "Đã tạo phím tắt: $lnk"
  }
}
$Script:LENH = [ordered]@{
  'kiem-tra'       = @('Test-MoiTruong',       'Kiểm tra máy (Excel, Word, tùy chọn PDF, file mẫu). -SuaLoi để tự bật tùy chọn Word')
  'chuan-bi'       = @('Invoke-ChuanBiFolder', '-Path <folder hồ sơ hoặc folder cha> [-LamLai]  → chữ PDF/DOCX, ảnh đầu trang, tom_tat.md (chỉ file mới)')
  'phan-tich'      = @('Invoke-PhanTich',      '-Path <1 folder hồ sơ> [-Out <file.json>]  → phân tích bằng quy tắc, đề xuất thay đổi Excel (không dùng AI)')
  'doi-ten-folder' = @('Rename-HoSoFolder',    '-Path <folder hồ sơ> -TenMoi <tên mới>  → đổi tên folder, cache đi theo (không phải OCR lại)')
  'doi-ten-tep'    = @('Rename-HoSoTep',       '-Path <folder hồ sơ> -File <tên file> -TenMoi <tên mới>  → đổi tên file văn bản, cache/nhập tay đi theo')
  'app'            = @('Start-App',            'Mở ứng dụng giao diện web trên máy (tương đương bấm "Mở ứng dụng.bat")')
  'tao-phim-tat'   = @('New-PhimTat',          '[-KhongDesktop]  → tạo phím tắt "Hồ sơ ABTC" có icon ở thư mục gốc và Desktop (chạy lại trên máy mới)')
  'doc-xlsx'       = @('Read-Xlsx',            '-Path <file.xlsx> [-Out <file.txt>]')
  'doc-docx'       = @('Read-Docx',            '-Path <file.docx> [-Out <file.txt>]')
  'pdf-chu'        = @('Get-PdfText',          '-Paths <a.pdf> <b.pdf> -OutDir <thư mục> [-TimeoutSec 180]')
  'pdf-dau-trang'  = @('Get-PdfHeader',        '-Path <file.pdf> -Out <ảnh.png> [-TopRatio 0.30] [-Width 900]')
  'tao-mau-excel'  = @('New-MauExcel',         '[-OutDir <thư mục>]  (mặc định KẾT QUẢ\MẪU)')
  'tao-mau-word'   = @('New-MauWord',          '-Source <báo cáo tuần gốc.docx> [-Out <file>] [-Org1 "SỞ NGOẠI VỤ"] [-Org2 "PHÒNG LÃNH SỰ"]')
  'chuyen-du-lieu' = @('Move-DuLieuCu',        '[-Force] [-Only <tên file.xlsx>] [-FallbackDir <thư mục bản sao>] [-OutDir <thư mục>]')
  'ghi-thay-doi'   = @('Set-ThayDoi',          '-File <file.xlsx> -Changes <thay_doi.json>  (tự sao lưu, kiểm tra giá trị cũ)')
  'thong-ke'       = @('New-ThongKe',          '-Data <ket_qua.json> -Out <file.xlsx> [-NgayLap dd/MM/yyyy] [-Nguon "mô tả"]')
  'tro-giup'       = @('Show-TroGiup',         'Hiện danh sách lệnh')
}
function Show-TroGiup {
  "ABTC_Tools.ps1 – thư mục gốc: $Script:GOC"; ''
  foreach ($k in $Script:LENH.Keys) { "  {0,-15} {1}" -f $k, $Script:LENH[$k][1] }
  ''; 'Ví dụ:  & "_HE_THONG\ABTC_Tools.ps1" chuan-bi -Path "Folder Hồ sơ ABTC Mẫu"'
}
function ConvertTo-Splat([object[]]$list) {
  # Chuyển dãy "-ThamSo giá_trị ..." thành bảng tham số (hỗ trợ nhiều giá trị và công tắc không giá trị)
  $h = @{}; $key = $null; $has = $false
  foreach ($a in $list) { foreach ($t in @($a)) {
    if ($t -is [string] -and $t -match '^-[A-Za-z][A-Za-z0-9_]*$') { if ($key -and -not $has) { $h[$key] = $true }; $key = $t.Substring(1); $has = $false; continue }
    if (-not $key) { continue }
    if ($has) { $h[$key] = @($h[$key]) + $t } else { $h[$key] = $t; $has = $true }
  } }
  if ($key -and -not $has) { $h[$key] = $true }
  $h
}

if ($MyInvocation.InvocationName -ne '.') {
  $cmd = if ($args.Count) { [string]$args[0] } else { 'tro-giup' }
  if (-not $Script:LENH.Contains($cmd)) { "Không có lệnh '$cmd'."; Show-TroGiup; return }
  $splat = ConvertTo-Splat @($args | Select-Object -Skip 1)
  $fn = $Script:LENH[$cmd][0]
  & $fn @splat
}
