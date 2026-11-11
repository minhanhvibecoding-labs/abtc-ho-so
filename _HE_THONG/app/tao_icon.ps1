# Vẽ icon ứng dụng Hồ sơ ABTC (thẻ doanh nhân + dấu tick đã duyệt) và xuất:
#   icon.ico (16–256px, dùng cho phím tắt), icon.png (256px), icon.svg (cùng thiết kế, dùng trong giao diện web)
# Dùng: powershell -ExecutionPolicy Bypass -File "_HE_THONG\app\tao_icon.ps1"
Add-Type -AssemblyName System.Drawing
$out = $PSScriptRoot
$C = @{ nen1 = '#2F74D6'; nen2 = '#0B3A82'; the = '#FFFFFF'; dai = '#E8F0FC'; anh = '#BCD3F3'; dong1 = '#8FA9D2'; dong2 = '#C9D6EA'; tick = '#1E9E5A' }
function Col($hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
function RoundRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath; $d = $r * 2
  $p.AddArc($x, $y, $d, $d, 180, 90); $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90); $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90); $p.CloseFigure(); $p
}
function Ve([int]$s) {
  # Vẽ trên lưới 256x256 rồi co theo kích thước s
  $bmp = New-Object System.Drawing.Bitmap $s, $s, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'; $g.PixelOffsetMode = 'HighQuality'; $g.InterpolationMode = 'HighQualityBicubic'
  $g.ScaleTransform($s / 256.0, $s / 256.0)
  # Nền bo góc, chuyển màu xanh
  $nen = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.PointF 0, 0), (New-Object System.Drawing.PointF 256, 256), (Col $C.nen1), (Col $C.nen2))
  $g.FillPath($nen, (RoundRect 0 0 256 256 56))
  # Tấm thẻ trắng + dải đầu thẻ
  $g.FillPath((New-Object System.Drawing.SolidBrush (Col $C.the)), (RoundRect 36 62 184 128 18))
  $g.SetClip((RoundRect 36 62 184 128 18)); $g.FillRectangle((New-Object System.Drawing.SolidBrush (Col $C.dai)), 36, 62, 184, 30); $g.ResetClip()
  # Ảnh chân dung (đầu + vai)
  $anh = New-Object System.Drawing.SolidBrush (Col $C.anh)
  $g.FillEllipse($anh, 64, 106, 40, 40)
  $vai = New-Object System.Drawing.Drawing2D.GraphicsPath; $vai.AddArc(54, 150, 60, 44, 180, 180); $vai.CloseFigure()
  $g.SetClip((RoundRect 36 62 184 128 18)); $g.FillPath($anh, $vai); $g.ResetClip()
  # Các dòng thông tin
  $g.FillPath((New-Object System.Drawing.SolidBrush (Col $C.dong1)), (RoundRect 124 112 72 11 5.5))
  $g.FillPath((New-Object System.Drawing.SolidBrush (Col $C.dong2)), (RoundRect 124 133 52 11 5.5))
  $g.FillPath((New-Object System.Drawing.SolidBrush (Col $C.dong2)), (RoundRect 124 154 62 11 5.5))
  # Dấu tick "đã duyệt": hình tròn xanh lá viền trắng
  $g.FillEllipse((New-Object System.Drawing.SolidBrush (Col $C.the)), 150, 146, 88, 88)
  $g.FillEllipse((New-Object System.Drawing.SolidBrush (Col $C.tick)), 158, 154, 72, 72)
  $pen = New-Object System.Drawing.Pen ((Col $C.the), 11); $pen.StartCap = 'Round'; $pen.EndCap = 'Round'; $pen.LineJoin = 'Round'
  $g.DrawLines($pen, [System.Drawing.PointF[]]@((New-Object System.Drawing.PointF 176, 191), (New-Object System.Drawing.PointF 189, 204), (New-Object System.Drawing.PointF 212, 177)))
  $g.Dispose(); $bmp
}

# PNG từng kích thước -> gộp thành .ico (mỗi mục là ảnh PNG, Windows Vista trở lên đọc được)
$sizes = 16, 24, 32, 48, 64, 128, 256
$pngs = foreach ($s in $sizes) { $b = Ve $s; $ms = New-Object IO.MemoryStream; $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); if ($s -eq 256) { $b.Save((Join-Path $out 'icon.png'), [System.Drawing.Imaging.ImageFormat]::Png) }; $b.Dispose(); , $ms.ToArray() }
$fs = [IO.File]::Create((Join-Path $out 'icon.ico')); $bw = New-Object IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$off = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
  $s = $sizes[$i]; $bw.Write([byte]($s % 256)); $bw.Write([byte]($s % 256)); $bw.Write([byte]0); $bw.Write([byte]0)
  $bw.Write([uint16]1); $bw.Write([uint16]32); $bw.Write([uint32]$pngs[$i].Length); $bw.Write([uint32]$off); $off += $pngs[$i].Length
}
foreach ($d in $pngs) { $bw.Write($d) }; $bw.Close()

# SVG cùng thiết kế (cho biểu tượng tab và logo trong giao diện web)
$svg = @"
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="$($C.nen1)"/><stop offset="1" stop-color="$($C.nen2)"/></linearGradient><clipPath id="c"><rect x="36" y="62" width="184" height="128" rx="18"/></clipPath></defs><rect width="256" height="256" rx="56" fill="url(#g)"/><rect x="36" y="62" width="184" height="128" rx="18" fill="$($C.the)"/><g clip-path="url(#c)"><rect x="36" y="62" width="184" height="30" fill="$($C.dai)"/><path d="M54 172a30 22 0 0 1 60 0z" fill="$($C.anh)"/></g><circle cx="84" cy="126" r="20" fill="$($C.anh)"/><rect x="124" y="112" width="72" height="11" rx="5.5" fill="$($C.dong1)"/><rect x="124" y="133" width="52" height="11" rx="5.5" fill="$($C.dong2)"/><rect x="124" y="154" width="62" height="11" rx="5.5" fill="$($C.dong2)"/><circle cx="194" cy="190" r="44" fill="$($C.the)"/><circle cx="194" cy="190" r="36" fill="$($C.tick)"/><path d="M176 191l13 13 23-27" fill="none" stroke="$($C.the)" stroke-width="11" stroke-linecap="round" stroke-linejoin="round"/></svg>
"@
[IO.File]::WriteAllText((Join-Path $out 'icon.svg'), $svg, (New-Object Text.UTF8Encoding($false)))
"Đã tạo icon.ico ($($sizes -join ', ') px), icon.png, icon.svg trong $out"
