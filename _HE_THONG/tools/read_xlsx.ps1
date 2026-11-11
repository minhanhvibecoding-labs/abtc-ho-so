# Đọc toàn bộ file .xlsx ra text (mỗi ô dạng "B5=giá trị"). Tự copy ra file tạm nên đọc được cả khi Excel đang mở file.
# Dùng: powershell -File read_xlsx.ps1 -Path "ABTC 2026.xlsx" -OutFile "out.txt"
param([string]$Path, [string]$OutFile)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmp = [System.IO.Path]::GetTempFileName()
Copy-Item -LiteralPath $Path $tmp -Force
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
function ReadEntry($name) {
  $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
  if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
  $t = $sr.ReadToEnd(); $sr.Close(); return [xml]$t
}
$ss = @()
$sst = ReadEntry 'xl/sharedStrings.xml'
if ($sst) {
  foreach ($si in $sst.sst.si) {
    if ($si.t) { if ($si.t -is [string]) { $ss += $si.t } else { $ss += $si.t.'#text' } }
    else { $ss += (($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '') }
  }
}
$wb = ReadEntry 'xl/workbook.xml'
$rels = ReadEntry 'xl/_rels/workbook.xml.rels'
$out = New-Object System.Text.StringBuilder
foreach ($s in $wb.workbook.sheets.sheet) {
  $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
  $target = $target -replace '^/xl/', '' -replace '^/', ''
  $sx = ReadEntry ("xl/" + $target)
  [void]$out.AppendLine("===== SHEET: " + $s.name + " =====")
  foreach ($row in $sx.worksheet.sheetData.row) {
    $cells = @()
    foreach ($c in $row.c) {
      if ($c.t -eq 's') { $v = $ss[[int]$c.v] } elseif ($c.t -eq 'inlineStr') { $v = $c.is.t } else { $v = $c.v }
      if ($v -ne $null -and "$v" -ne '') { $cells += ($c.r + "=" + (("$v") -replace "`r?`n", ' / ')) }
    }
    if ($cells.Count) { [void]$out.AppendLine(($cells -join ' | ')) }
  }
}
$zip.Dispose(); Remove-Item $tmp
[System.IO.File]::WriteAllText($OutFile, $out.ToString(), [System.Text.Encoding]::UTF8)
