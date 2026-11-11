# Tạo template Excel rỗng từ các file cơ sở dữ liệu gốc.
# Giữ nguyên: tiêu đề, ô gộp, độ rộng cột, định dạng, tên sheet.
# Xóa: toàn bộ dòng dữ liệu, chỉ giữ lại 1 "dòng mẫu" trống (đã có viền/định dạng) ngay dưới tiêu đề.
# Dùng: powershell -File build_excel_templates.ps1 -Root "<thư mục CƠ SỞ DỮ LIỆU 2026>" -OutDir "<thư mục MẪU>"
param([string]$Root, [string]$OutDir)

# Cấu hình (file gốc -> sheet -> dòng mẫu) đọc từ template_config.json, dùng chung với migrate_data.ps1
$json = Get-Content (Join-Path $PSScriptRoot 'template_config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$config = foreach ($f in $json.files) {
  $h = @{}; foreach ($s in $f.sheets) { $h[$s.sheet] = [int]$s.dong_mau }
  @{ File = $f.file; Sheets = $h }
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$report = @()
try {
  foreach ($c in $config) {
    $src = Join-Path (Join-Path $Root $json.thu_muc_goc) $c.File
    $dst = Join-Path $OutDir ("MẪU - " + $c.File)
    Copy-Item -LiteralPath $src $dst -Force          # copy được cả khi file gốc đang mở
    $wb = $xl.Workbooks.Open($dst)
    foreach ($ws in $wb.Worksheets) {
      if (-not $c.Sheets.ContainsKey($ws.Name)) { $report += "$($c.File) | $($ws.Name) | giữ nguyên (không cấu hình)"; continue }
      $start = [int]$c.Sheets[$ws.Name]
      $ur = $ws.UsedRange
      $last = $ur.Row + $ur.Rows.Count - 1
      if ($last -gt $start) { [void]$ws.Range("$($start + 1):$last").EntireRow.Delete() }
      $styleRow = $ws.Rows($start)
      [void]$styleRow.ClearContents()
      [void]$styleRow.ClearComments()
      $styleRow.Interior.ColorIndex = -4142      # bỏ tô màu đánh dấu của dữ liệu cũ
      $report += "$($c.File) | $($ws.Name) | dữ liệu từ dòng $start | đã xóa dòng $start-$last"
    }
    # Sheet "CÁN BỘ DIỆN BTV TU…": ô A1 có chữ "Y" thừa
    foreach ($ws in $wb.Worksheets) { if ($ws.Name -eq 'CÁN BỘ DIỆN BTV TU ĐI VIỆC RIÊN' -and $ws.Range('A1').Text -eq 'Y') { $ws.Range('A1').ClearContents() | Out-Null } }
    $wb.Worksheets.Item(1).Activate()
    [void]$wb.Worksheets.Item(1).Range('A1').Select()
    $wb.Save()
    $wb.Close($false)
  }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
$report
