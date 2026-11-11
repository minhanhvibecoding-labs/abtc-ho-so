# Tạo template báo cáo tuần (.docx) từ một báo cáo tuần gốc.
# - Giữ: tiêu ngữ, tiêu đề, 2 dòng tiêu đề bảng, các dòng mục (I, 1, 2, 2.1, ..., II) và dòng mục con (vd "Hồ sơ hành chính").
# - Xóa chữ ở các cột nội dung của dòng mục.
# - Dưới mỗi mục chỉ giữ 1 dòng nội dung trống (dòng mẫu để nhân bản khi điền), xóa các dòng còn lại.
# - Thay khoảng ngày ở tiêu đề bằng {{TU_NGAY}} - {{DEN_NGAY}}.
# Dùng: powershell -File build_word_template.ps1 -Source "<báo cáo gốc.docx>" -Out "<template.docx>"
param([string]$Source, [string]$Out,
      [string]$Org1 = 'SỞ NGOẠI VỤ',      # dòng 1 tiêu ngữ (chữ thường)
      [string]$Org2 = 'PHÒNG LÃNH SỰ')    # dòng 2 tiêu ngữ (chữ đậm)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

Copy-Item -LiteralPath $Source $Out -Force
$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Update')
try {
  $entry = $zip.GetEntry('word/document.xml')
  $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $true
  $xml.LoadXml($sr.ReadToEnd()); $sr.Close()
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace('w', $W)

  function CellText($tc) { (($tc.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim() }
  function ClearCell($tc) {
    $ps = @($tc.SelectNodes('./w:p', $ns))
    for ($i = 1; $i -lt $ps.Count; $i++) { [void]$tc.RemoveChild($ps[$i]) }   # giữ đoạn đầu tiên
    if ($ps.Count -gt 0) { foreach ($r in @($ps[0].SelectNodes('./w:r | ./w:hyperlink', $ns))) { [void]$ps[0].RemoveChild($r) } }
  }

  # 1) Tiêu đề: "... tuần từ ngày 19/01 - 23/01/2026"
  foreach ($p in $xml.SelectNodes('//w:body/w:p', $ns)) {
    $txt = ($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    if ($txt -match 'tuần từ ngày') {
      $new = ($txt -replace 'tuần từ ngày.*$', 'tuần từ ngày {{TU_NGAY}} - {{DEN_NGAY}}')
      $ts = @($p.SelectNodes('.//w:t', $ns))
      $ts[0].InnerText = $new
      [void]$ts[0].SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
      for ($i = 1; $i -lt $ts.Count; $i++) { $ts[$i].InnerText = '' }
    }
  }

  # 1b) Tiêu ngữ cơ quan: ô đầu tiên của bảng tiêu ngữ, 2 đoạn có chữ đầu tiên -> Org1, Org2 (giữ định dạng run đầu)
  $orgCell = $xml.SelectSingleNode('//w:body/w:tbl[1]/w:tr[1]/w:tc[1]', $ns)
  $orgParas = @($orgCell.SelectNodes('./w:p', $ns) | Where-Object { (($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim() })
  $newOrg = @($Org1, $Org2)
  for ($k = 0; $k -lt [Math]::Min(2, $orgParas.Count); $k++) {
    $ts = @($orgParas[$k].SelectNodes('.//w:t', $ns))
    $ts[0].InnerText = $newOrg[$k]
    for ($j = 1; $j -lt $ts.Count; $j++) { $ts[$j].InnerText = '' }
  }

  # 2) Bảng chính = bảng thứ 2 trong thân văn bản
  $tbl = @($xml.SelectNodes('//w:body/w:tbl', $ns))[1]
  $rows = @($tbl.SelectNodes('./w:tr', $ns))
  $keptBlank = $false
  $stats = @{ muc = 0; dong_mau = 0; xoa = 0 }
  for ($i = 2; $i -lt $rows.Count; $i++) {         # bỏ qua 2 dòng tiêu đề bảng
    $tr = $rows[$i]
    $cells = @($tr.SelectNodes('./w:tc', $ns))
    $c1 = CellText $cells[0]; $c2 = if ($cells.Count -gt 1) { CellText $cells[1] } else { '' }
    if ($c1 -or $c2) {
      for ($k = 2; $k -lt $cells.Count; $k++) { ClearCell $cells[$k] }   # dòng mục: giữ STT + tên mục
      $keptBlank = $false; $stats.muc++
    } elseif (-not $keptBlank) {
      foreach ($tc in $cells) { ClearCell $tc }                         # dòng mẫu trống
      $keptBlank = $true; $stats.dong_mau++
    } else {
      [void]$tbl.RemoveChild($tr); $stats.xoa++
    }
  }

  # 3) Chuẩn hóa dòng mẫu: lấy dòng trống 8 ô đầu tiên làm chuẩn
  $rows = @($tbl.SelectNodes('./w:tr', $ns))
  $std = $null
  for ($i = 2; $i -lt $rows.Count; $i++) {
    $cs = @($rows[$i].SelectNodes('./w:tc', $ns))
    if ($cs.Count -eq 8 -and -not (CellText $cs[0]) -and -not (CellText $cs[1])) { $std = $rows[$i]; break }
  }
  $stats.thay_the = 0; $stats.them = 0
  for ($i = 2; $i -lt $rows.Count; $i++) {
    $tr = $rows[$i]; $cs = @($tr.SelectNodes('./w:tc', $ns))
    $isHead = (CellText $cs[0]) -or ($cs.Count -gt 1 -and (CellText $cs[1]))
    if (-not $isHead -and $cs.Count -ne 8) {                  # dòng mẫu bị lỗi số ô -> thay bằng dòng chuẩn
      [void]$tbl.ReplaceChild($std.CloneNode($true), $tr); $stats.thay_the++
    }
  }
  # 4) Mục không có mục con và chưa có dòng mẫu -> thêm dòng mẫu
  $rows = @($tbl.SelectNodes('./w:tr', $ns))
  for ($i = 2; $i -lt $rows.Count; $i++) {
    $cs = @($rows[$i].SelectNodes('./w:tc', $ns))
    $c1 = CellText $cs[0]; $c2 = if ($cs.Count -gt 1) { CellText $cs[1] } else { '' }
    if (-not ($c1 -or $c2)) { continue }
    if ($c1 -match '^[IVX]+$' -and $c1 -ne 'II') { continue }            # "I" là mục lớn, có mục con
    $next = if ($i + 1 -lt $rows.Count) { $rows[$i + 1] } else { $null }
    if ($next) {
      $ncs = @($next.SelectNodes('./w:tc', $ns))
      $n1 = CellText $ncs[0]; $n2 = if ($ncs.Count -gt 1) { CellText $ncs[1] } else { '' }
      if (-not ($n1 -or $n2)) { continue }                               # đã có dòng mẫu
      if ($c1 -and ($n1 -like "$c1.*" -or (-not $n1 -and $n2))) { continue }  # có mục con (2 -> 2.1, 7 -> "Hồ sơ hành chính")
    }
    [void]$tbl.InsertAfter($std.CloneNode($true), $rows[$i]); $stats.them++
  }

  # 5) Cắt cột "ma" thừa trong khung bảng (do dòng lỗi của bản gốc tạo ra), giữ đúng độ rộng 8 cột thật
  $maxCols = 0
  foreach ($tr in $tbl.SelectNodes('./w:tr', $ns)) {
    $n = 0
    foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) { $gs = $tc.SelectSingleNode('./w:tcPr/w:gridSpan', $ns); $n += $(if ($gs) { [int]$gs.GetAttribute('val', $W) } else { 1 }) }
    if ($n -gt $maxCols) { $maxCols = $n }
    foreach ($tag in 'gridAfter', 'wAfter', 'gridBefore', 'wBefore') { $el = $tr.SelectSingleNode("./w:trPr/w:$tag", $ns); if ($el) { [void]$el.ParentNode.RemoveChild($el) } }
  }
  $grid = $tbl.SelectSingleNode('./w:tblGrid', $ns)
  $cols = @($grid.SelectNodes('./w:gridCol', $ns))
  for ($k = $maxCols; $k -lt $cols.Count; $k++) { [void]$grid.RemoveChild($cols[$k]) }
  $sum = 0; foreach ($gc in $grid.SelectNodes('./w:gridCol', $ns)) { $sum += [int]$gc.GetAttribute('w', $W) }
  $tblW = $tbl.SelectSingleNode('./w:tblPr/w:tblW', $ns)
  if ($tblW) { [void]$tblW.SetAttribute('w', $W, "$sum"); [void]$tblW.SetAttribute('type', $W, 'dxa') }
  $stats.so_cot = $maxCols

  $entry.Delete()
  $newEntry = $zip.CreateEntry('word/document.xml')
  $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object Text.UTF8Encoding($false)))
  $xml.Save($sw); $sw.Close()
  "Dòng mục: $($stats.muc) | Dòng mẫu trống: $($stats.dong_mau) | Dòng đã xóa: $($stats.xoa) | Dòng lỗi đã thay: $($stats.thay_the) | Dòng mẫu thêm mới: $($stats.them) | Số cột bảng: $($stats.so_cot)"
} finally { $zip.Dispose() }
