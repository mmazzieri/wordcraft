# Builds the synthetic reproducer with Word (COM): a header and a footer taller than the margins,
# with the top and bottom margins negative ("exact") in one document and positive in the other.
# Saves .docx, .doc and Word's PDF of each. Text is invented.
param([string] $Out = $PSScriptRoot)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
function New-Case([string] $name, [single] $top, [single] $bottom) {
    $d = $word.Documents.Add()
    $d.RemovePersonalInformation = $true
    $ps = $d.Sections(1).PageSetup
    $ps.PaperSize = 7          # A4
    $ps.HeaderDistance = 36
    $ps.FooterDistance = 36
    $ps.LeftMargin = 72; $ps.RightMargin = 72
    $ps.TopMargin = $top
    $ps.BottomMargin = $bottom
    $h = $d.Sections(1).Headers(1).Range
    $h.Text = "HEADER LINE 1`rHEADER LINE 2`rHEADER LINE 3`rHEADER LINE 4`rHEADER LINE 5`rHEADER LINE 6"
    $h.Font.Size = 14
    $f = $d.Sections(1).Footers(1).Range
    $f.Text = "FOOTER LINE 1`rFOOTER LINE 2`rFOOTER LINE 3`rFOOTER LINE 4`rFOOTER LINE 5`rFOOTER LINE 6"
    $f.Font.Size = 14
    $body = $d.Content
    $lines = 1..100 | ForEach-Object { "Body line $_ of the invented text." }
    $body.Text = ($lines -join "`r")
    $body.Font.Size = 11
    $body.ParagraphFormat.SpaceAfter = 0
    $body.ParagraphFormat.LineSpacingRule = 0
    $base = Join-Path $Out $name
    $d.SaveAs2("$base.docx", 16)
    $d.SaveAs2("$base.doc", 0)
    $d.ExportAsFixedFormat("$base-word.pdf", 17)
    $n = $d.ComputeStatistics(2)
    $d.Close(0)
    "{0}: Word {1} pages" -f $name, $n
}
try {
    New-Case 'negative-margins' -60 -60
    New-Case 'positive-margins' 60 60
} finally { $word.Quit() }
