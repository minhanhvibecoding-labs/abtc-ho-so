# Chuyển dữ liệu cũ từ file gốc sang file kết quả (dựng từ template).
# - File kết quả = bản sao template trong KẾT QUẢ\MẪU (bỏ tiền tố "MẪU - ").
# - Chép nguyên khối dòng dữ liệu (giá trị + định dạng + chiều cao dòng) từ file gốc vào đúng vị trí dòng cũ,
#   nên số dòng trong file kết quả trùng với file gốc (vd một hồ sơ vẫn ở dòng 112 của ABTC).
# - File gốc chỉ được đọc qua bản sao tạm, không bị sửa.
# - File gốc đang bị Excel khóa: dùng bản sao cùng tên trong -FallbackDir (nếu có), nếu không thì bỏ qua và báo.
# Dùng: powershell -File migrate_data.ps1 -Root "<thư mục CƠ SỞ DỮ LIỆU 2026>" [-Force] [-Only "ABTC 2026.xlsx"] [-FallbackDir "<thư mục bản sao>"]
param([string]$Root, [switch]$Force, [string[]]$Only, [string]$FallbackDir)

$cfg = Get-Content (Join-Path $PSScriptRoot 'template_config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$tplDir = Join-Path $Root 'KẾT QUẢ\MẪU'
$outDir = Join-Path $Root 'KẾT QUẢ\CƠ SỞ DỮ LIỆU'
$tmpDir = Join-Path $env:TEMP ('migrate_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $outDir, $tmpDir | Out-Null

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$report = @()
try {
  foreach ($f in $cfg.files) {
    if ($Only -and ($Only -notcontains $f.file)) { continue }
    $dst = Join-Path $outDir $f.file
    if ((Test-Path -LiteralPath $dst) -and -not $Force) { $report += "BỎ QUA $($f.file): file kết quả đã có (dùng -Force để làm lại)"; continue }

    # Lấy bản sao của file gốc
    $srcCopy = Join-Path $tmpDir ('NGUON_' + $f.file)   # Excel không mở được 2 file trùng tên cùng lúc
    $srcNote = 'file gốc'
    try { Copy-Item -LiteralPath (Join-Path (Join-Path $Root $cfg.thu_muc_goc) $f.file) $srcCopy -Force -ErrorAction Stop }
    catch {
      $fb = if ($FallbackDir) { Join-Path $FallbackDir $f.file } else { $null }
      if ($fb -and (Test-Path -LiteralPath $fb)) { Copy-Item -LiteralPath $fb $srcCopy -Force; $srcNote = "bản sao dự phòng ($fb)" }
      else { $report += "LỖI $($f.file): file gốc đang bị khóa, chưa chuyển. Hãy lưu và đóng file trong Excel rồi chạy lại với -Only"; continue }
    }

    $wbS = $null; $wbD = $null
    try {
      Copy-Item -LiteralPath (Join-Path $tplDir ('MẪU - ' + $f.file)) $dst -Force -ErrorAction Stop
      $wbS = $xl.Workbooks.Open($srcCopy, 0, $true)
      $wbD = $xl.Workbooks.Open($dst)
      foreach ($s in $f.sheets) {
        $wsS = $wbS.Worksheets.Item($s.sheet); $wsD = $wbD.Worksheets.Item($s.sheet)
        $start = [int]$s.dong_mau
        $lastCell = $wsS.Cells.Find('*', [Type]::Missing, -4123, 2, 1, 2)   # dòng cuối có nội dung
        $last = if ($lastCell) { $lastCell.Row } else { 0 }
        if ($last -lt $start) { $report += "$($f.file) | $($s.sheet) | không có dữ liệu"; continue }
        [void]$wsS.Range("$($start):$last").EntireRow.Copy($wsD.Range("A$start"))
        $n = 0; for ($r = $start; $r -le $last; $r++) { if ($xl.WorksheetFunction.CountA($wsS.Rows($r)) -gt 0) { $n++ } }
        $report += "$($f.file) | $($s.sheet) | dòng $start-$last | $n dòng có dữ liệu | nguồn: $srcNote"
      }
      $wbD.Worksheets.Item(1).Activate(); [void]$wbD.Worksheets.Item(1).Range('A1').Select()
      $wbD.Save()
    } catch {
      $report += "LỖI $($f.file): $($_.Exception.Message) @ $($_.InvocationInfo.Line.Trim())"
    } finally {
      if ($wbD) { $wbD.Close($false) }
      if ($wbS) { $wbS.Close($false) }
    }
  }
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
}
$report
