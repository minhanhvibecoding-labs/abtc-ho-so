# Đóng gói ứng dụng Hồ sơ ABTC thành 1 file .exe (trình khởi chạy C# + gói phần mềm nhúng sẵn). Không cần cài thêm gì: dùng csc.exe của .NET Framework có sẵn trong Windows.
# Cách dùng: powershell -ExecutionPolicy Bypass -File tao_exe.ps1 -Nguon "<thư mục phần mềm SẠCH, vd bản clone từ GitHub>" -Ra "<file .exe>"
# LƯU Ý: -Nguon phải là bản phần mềm KHÔNG có dữ liệu hồ sơ (chỉ mã + tài liệu + mẫu rỗng) – exe sẽ chứa mọi file trong đó.
# Nhớ tăng PHIEN_BAN và AssemblyVersion trong HoSoABTC.cs khi phát hành bản mới (exe chỉ đề nghị cập nhật khi phiên bản lớn hơn).
param([Parameter(Mandatory)][string]$Nguon, [Parameter(Mandatory)][string]$Ra)
$fw = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319'; $tam = Join-Path $env:TEMP ('abtc_exe_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tam | Out-Null
try {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $app = Join-Path $tam 'app'; Copy-Item -LiteralPath $Nguon -Destination $app -Recurse
  Remove-Item -LiteralPath (Join-Path $app '.git') -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath (Join-Path $app '.gitignore') -Force -ErrorAction SilentlyContinue
  [IO.Compression.ZipFile]::CreateFromDirectory($app, (Join-Path $tam 'app.zip'), [IO.Compression.CompressionLevel]::Optimal, $false, [Text.Encoding]::UTF8)
  & "$fw\csc.exe" /nologo /codepage:65001 /target:winexe /platform:anycpu /optimize+ /out:"$Ra" /win32icon:"$(Join-Path $app '_HE_THONG\app\icon.ico')" /resource:"$(Join-Path $tam 'app.zip')",app.zip /r:System.dll /r:System.Windows.Forms.dll /r:"$fw\System.IO.Compression.dll" /r:"$fw\System.IO.Compression.FileSystem.dll" (Join-Path $PSScriptRoot 'HoSoABTC.cs')
  if ($LASTEXITCODE -ne 0) { throw "Biên dịch lỗi ($LASTEXITCODE)" }
  "Đã tạo: $Ra"
} finally { Remove-Item -LiteralPath $tam -Recurse -Force -ErrorAction SilentlyContinue }