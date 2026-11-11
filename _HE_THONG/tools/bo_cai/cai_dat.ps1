# Cài đặt / cập nhật / mở ứng dụng Hồ sơ ABTC – KHÔNG dùng file .exe nên chạy được trên máy bật Smart App Control.
#  - Lần đầu: chép phần mềm vào Documents\CƠ SỞ DỮ LIỆU 2026, tạo lối tắt "Hồ sơ ABTC" trên Desktop, mở ứng dụng.
#  - Bộ cài mới hơn bản đã cài (do bộ cài này cài): hỏi rồi cập nhật phần mềm; KHÔNG đụng dữ liệu hồ sơ, kho, Excel tổng hợp, mẫu Excel đã có.
#  - Thư mục đã có phần mềm nhưng không do bộ cài cài: giữ nguyên, chỉ mở ứng dụng.
# Thử nghiệm: -Goc "<thư mục>" -Thu (không hỏi, không tạo lối tắt, không mở ứng dụng).
param([string]$Goc = '', [switch]$Thu)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$PHIEN_BAN = '2026.11.12'
$nguon = Join-Path $PSScriptRoot 'phan_mem'
if (-not $Goc) { $Goc = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CƠ SỞ DỮ LIỆU 2026' }
$tools = Join-Path $Goc '_HE_THONG\ABTC_Tools.ps1'
$verFile = Join-Path $Goc '_HE_THONG\phien_ban_app.txt'

function Chep-PhanMem([bool]$capNhat) {
  $n = 0
  foreach ($f in Get-ChildItem -LiteralPath $nguon -Recurse -File) {
    $dich = Join-Path $Goc $f.FullName.Substring($nguon.Length + 1)
    if ($capNhat -and $dich -like '*.xlsx' -and (Test-Path -LiteralPath $dich)) { continue }   # giữ mẫu Excel đã có
    New-Item -ItemType Directory -Force (Split-Path $dich) | Out-Null
    Copy-Item -LiteralPath $f.FullName -Destination $dich -Force
    try { Unblock-File -LiteralPath $dich } catch { }   # bỏ dấu "tải từ internet"
    $n++
  }
  foreach ($d in 'KẾT QUẢ\CƠ SỞ DỮ LIỆU', 'Folder Hồ sơ ABTC Mẫu') { New-Item -ItemType Directory -Force (Join-Path $Goc $d) | Out-Null }
  # File Excel tổng hợp: tạo từ mẫu rỗng nếu chưa có (không bao giờ ghi đè file đã có dữ liệu)
  $xl = Join-Path $Goc 'KẾT QUẢ\CƠ SỞ DỮ LIỆU\ABTC 2026.xlsx'; $mau = Join-Path $Goc 'KẾT QUẢ\MẪU\MẪU - ABTC 2026.xlsx'
  if (-not (Test-Path -LiteralPath $xl) -and (Test-Path -LiteralPath $mau)) { Copy-Item -LiteralPath $mau -Destination $xl; try { Unblock-File -LiteralPath $xl } catch { } }
  [IO.File]::WriteAllText($verFile, $PHIEN_BAN, (New-Object Text.UTF8Encoding($false)))
  "  Đã chép $n tệp phần mềm."
}

try {
  Write-Host "=== Hồ sơ ABTC – bộ cài bản $PHIEN_BAN ===" -ForegroundColor Cyan
  if (-not (Test-Path -LiteralPath $tools)) {
    Write-Host "Cài mới vào: $Goc"
    Chep-PhanMem $false
    if (-not $Thu) { & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tools tao-phim-tat | Out-Null; Write-Host '  Đã tạo lối tắt "Hồ sơ ABTC" trên Desktop.' }
  } elseif (Test-Path -LiteralPath $verFile) {
    $cu = (Get-Content -LiteralPath $verFile -Raw).Trim()
    if ([string]::CompareOrdinal($cu, $PHIEN_BAN) -lt 0) {
      $tl = if ($Thu) { 'C' } else { Read-Host "Máy đang có bản $cu. Cập nhật lên bản $PHIEN_BAN? Dữ liệu hồ sơ giữ nguyên. (C = có / K = không)" }
      if ($tl -match '^[cCyY]') { Write-Host 'Đang cập nhật…'; Chep-PhanMem $true } else { Write-Host 'Giữ bản đang có.' }
    } else { Write-Host "Máy đã có bản $cu – không cần cài lại." }
  } else {
    Write-Host "Thư mục $Goc đã có phần mềm (không do bộ cài này cài) – giữ nguyên, không ghi đè."
  }
  if ($Thu) { exit 0 }
  # Mở ứng dụng: đang chạy thì mở lại cửa sổ, chưa chạy thì khởi động máy chủ
  $url = $null
  foreach ($p in 8765..8775) { try { $r = Invoke-WebRequest "http://localhost:$p/api/tong-quan" -UseBasicParsing -TimeoutSec 1; if ($r.Content -match '"goc"') { $url = "http://localhost:$p/"; break } } catch { } }
  if ($url) { try { Start-Process 'msedge.exe' "--app=$url" } catch { Start-Process $url } }
  else { Start-Process -FilePath (Join-Path $Goc 'Mở ứng dụng.bat') -WorkingDirectory $Goc }
  Write-Host ''; Write-Host 'Xong. Lần sau mở bằng lối tắt "Hồ sơ ABTC" trên Desktop.' -ForegroundColor Green
  Start-Sleep -Seconds 4
} catch {
  Write-Host ''; Write-Host ('LỖI: ' + $_.Exception.Message) -ForegroundColor Red
  if (-not $Thu) { Read-Host 'Bấm Enter để đóng' }
  exit 1
}