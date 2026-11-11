# Ghi các thay đổi đã được người dùng duyệt vào một file Excel.
# - Sao lưu file vào _HE_THONG\backup\ trước khi ghi.
# - Kiểm tra giá trị hiện tại của từng ô phải đúng "cu" (giá trị lúc duyệt). Lệch bất kỳ ô nào -> dừng, không ghi gì.
# - Ghi dạng chữ (NumberFormat "@"), bật Wrap Text cho ô nhiều dòng.
# File JSON thay đổi: [ { "sheet": "ABTC 2026", "o": "L112", "cu": "", "moi": "123/TTr-NV ngày 21/09/2026" }, ... ]
#   Trong "moi", ký tự "\n" là xuống dòng trong ô.
# Dùng: powershell -File apply_changes.ps1 -File "<file.xlsx>" -Changes "<thay_doi.json>" -BackupDir "<…\_HE_THONG\backup>"
param([string]$File, [string]$Changes, [string]$BackupDir)

$list = Get-Content -LiteralPath $Changes -Raw -Encoding UTF8 | ConvertFrom-Json
New-Item -ItemType Directory -Force $BackupDir | Out-Null
$bk = Join-Path $BackupDir ('{0}_{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($File), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($File))
Copy-Item -LiteralPath $File $bk -ErrorAction Stop
"Đã sao lưu: $bk"

function Norm([string]$s) { if ($null -eq $s) { '' } else { ($s -replace "`r", '').Trim() } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $null
try {
  $wb = $xl.Workbooks.Open($File)
  if ($wb.ReadOnly) { throw "File đang mở ở nơi khác (chỉ đọc). Hãy đóng file trong Excel rồi chạy lại." }
  # 1) Kiểm tra toàn bộ trước
  $bad = @()
  foreach ($c in $list) {
    $cur = Norm ([string]$wb.Worksheets.Item($c.sheet).Range($c.o).Value2)
    if ($cur -ne (Norm $c.cu)) { $bad += "$($c.sheet)!$($c.o): hiện là [$cur], lúc duyệt là [$($c.cu)]" }
  }
  if ($bad) { throw ("Giá trị trong file đã thay đổi so với lúc duyệt, KHÔNG ghi:`n  " + ($bad -join "`n  ")) }
  # 2) Ghi
  foreach ($c in $list) {
    $r = $wb.Worksheets.Item($c.sheet).Range($c.o)
    $r.NumberFormat = '@'
    $v = ([string]$c.moi) -replace '\\n', "`n"
    $r.Value2 = $v
    if ($v.Contains("`n") -or $v.Length -gt 40) { $r.WrapText = $true }
    "Đã ghi $($c.sheet)!$($c.o)"
  }
  $wb.Save()
} finally {
  if ($wb) { $wb.Close($false) }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
