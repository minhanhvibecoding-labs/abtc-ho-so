# Đóng gói BỘ CÀI Hồ sơ ABTC (ZIP: "Cai dat Ho so ABTC.bat" + _cai_dat\cai_dat.ps1 + _cai_dat\phan_mem\…). Không dùng .exe → chạy được trên máy bật Smart App Control.
# Cách dùng: powershell -ExecutionPolicy Bypass -File tao_bo_cai.ps1 -Nguon "<thư mục phần mềm SẠCH, vd bản clone từ GitHub>" -Ra "<file .zip>"
# LƯU Ý: -Nguon phải là bản KHÔNG có dữ liệu hồ sơ. Nhớ tăng $PHIEN_BAN trong cai_dat.ps1 khi phát hành bản mới.
param([Parameter(Mandatory)][string]$Nguon, [Parameter(Mandatory)][string]$Ra)
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$ver = if ((Get-Content (Join-Path $PSScriptRoot 'cai_dat.ps1') -Raw) -match "PHIEN_BAN = '([^']+)'") { $Matches[1] } else { 'moi' }
$tam = Join-Path $env:TEMP ('abtc_bocai_' + [guid]::NewGuid().ToString('N')); $bo = Join-Path $tam "Ho so ABTC - Bo cai $ver"
try {
  New-Item -ItemType Directory (Join-Path $bo '_cai_dat') | Out-Null
  Copy-Item -LiteralPath $Nguon -Destination (Join-Path $bo '_cai_dat\phan_mem') -Recurse
  foreach ($x in '.git', '.gitignore') { Remove-Item -LiteralPath (Join-Path $bo "_cai_dat\phan_mem\$x") -Recurse -Force -ErrorAction SilentlyContinue }
  Copy-Item (Join-Path $PSScriptRoot 'cai_dat.ps1') (Join-Path $bo '_cai_dat\')
  Copy-Item (Join-Path $PSScriptRoot 'Cai dat Ho so ABTC.bat'), (Join-Path $PSScriptRoot 'HUONG DAN.txt') $bo
  if (Test-Path -LiteralPath $Ra) { Remove-Item -LiteralPath $Ra -Force }
  $za = [IO.Compression.ZipFile]::Open($Ra, 'Create', [Text.Encoding]::UTF8)
  try { foreach ($f in Get-ChildItem -LiteralPath $bo -Recurse -File) { [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($za, $f.FullName, ($f.FullName.Substring($tam.Length + 1) -replace '\\', '/'), 'Optimal') } } finally { $za.Dispose() }
  "Đã tạo bộ cài $ver : $Ra"
} finally { Remove-Item -LiteralPath $tam -Recurse -Force -ErrorAction SilentlyContinue }