# Đọc file .docx ra text, giữ cấu trúc bảng (dòng bảng dạng "| ô 1 | ô 2 |", xuống dòng trong ô thành " / ").
# Dùng: powershell -File read_docx.ps1 -Path "van_ban.docx" -OutFile "out.txt"
param([string]$Path, [string]$OutFile)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmp = [System.IO.Path]::GetTempFileName()
Copy-Item -LiteralPath $Path $tmp -Force
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
$e = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
$sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
$xml = [xml]$sr.ReadToEnd(); $sr.Close(); $zip.Dispose(); Remove-Item $tmp
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$out = New-Object System.Text.StringBuilder
foreach ($node in $xml.SelectSingleNode('//w:body', $ns).ChildNodes) {
  if ($node.LocalName -eq 'p') {
    $t = ($node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    if ($t.Trim()) { [void]$out.AppendLine($t) }
  } elseif ($node.LocalName -eq 'tbl') {
    [void]$out.AppendLine('[TABLE]')
    foreach ($tr in $node.SelectNodes('./w:tr', $ns)) {
      $cells = foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) {
        (($tc.SelectNodes('.//w:p', $ns) | ForEach-Object { ($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '' }) -join ' / ').Trim()
      }
      [void]$out.AppendLine('| ' + ($cells -join ' | ') + ' |')
    }
    [void]$out.AppendLine('[/TABLE]')
  }
}
[System.IO.File]::WriteAllText($OutFile, $out.ToString(), [System.Text.Encoding]::UTF8)
