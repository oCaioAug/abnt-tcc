# Opens each build/*.docx in Microsoft Word, updates the table of contents,
# figure/table lists and fields, then saves to dist/:
#   <name>.dotx  (Word template, also opens in LibreOffice)
#   <name>.docx  (ready-to-edit document)
#   <name>.pdf   (preview)
$ErrorActionPreference = "Stop"
$root  = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root "build"
$dist  = Join-Path $root "dist"
New-Item -ItemType Directory -Force $dist | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  foreach ($f in Get-ChildItem $build -Filter *.docx) {
    $doc = $word.Documents.Open($f.FullName, $false, $false)
    $doc.Repaginate()
    foreach ($s in $doc.StoryRanges) { [void]$s.Fields.Update() }
    foreach ($t in $doc.TablesOfContents) { $t.Update() }
    foreach ($t in $doc.TablesOfFigures) { $t.Update() }    # \c "Figura" / "Tabela" lists
    $doc.Repaginate()
    foreach ($t in $doc.TablesOfContents) { $t.Update() }   # 2nd pass: page numbers settle
    foreach ($t in $doc.TablesOfFigures) { $t.Update() }

    $base = Join-Path $dist $f.BaseName
    $doc.SaveAs2("$base.pdf", 17)    # wdFormatPDF
    $doc.SaveAs2("$base.docx", 16)   # wdFormatXMLDocument
    $doc.SaveAs2("$base.dotx", 14)   # wdFormatXMLTemplate
    $doc.Close($false)
    Write-Output "ok: $base.{dotx,docx,pdf}"
  }
} finally {
  $word.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
