# Synthetic reproducer: one text line, then a paragraph holding only an inline picture of height H,
# Normal style 1.15 lines (line="276" auto), A4, top 70.85 pt, bottom 56.7 pt. One document per H
# and compatibility mode; Word's page count is printed. Text and picture are invented.
param([string] $Out = $PSScriptRoot, [double[]] $Heights = @(696, 697, 698, 698.5, 699), [int[]] $Modes = @(11, 15))
$png = Join-Path $Out 'grey.png'
if (-not (Test-Path $png)) { & magick -size 200x200 xc:'#b0b0b0' $png }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($mode in $Modes) {
        foreach ($h in $Heights) {
            $d = $word.Documents.Add()
            $d.RemovePersonalInformation = $true
            $d.SetCompatibilityMode($mode)
            $ps = $d.Sections(1).PageSetup
            $ps.PageWidth = 595.3; $ps.PageHeight = 841.9
            $ps.TopMargin = 70.85; $ps.BottomMargin = 56.7; $ps.LeftMargin = 56.7; $ps.RightMargin = 56.7
            $n = $d.Styles.Item(-1)   # Normal
            $n.Font.Name = 'Calibri'; $n.Font.Size = 11
            $n.ParagraphFormat.SpaceAfter = 0; $n.ParagraphFormat.SpaceBefore = 0
            $n.ParagraphFormat.LineSpacingRule = 5; $n.ParagraphFormat.LineSpacing = 13.8   # multiple 1.15
            $d.Content.Text = "A line of invented text.`r"
            $r = $d.Paragraphs(2).Range
            $s = $d.InlineShapes.AddPicture($png, $false, $true, $r)
            $s.LockAspectRatio = 0
            $s.Height = $h; $s.Width = 200
            $name = [string]::Format([cultureinfo]::InvariantCulture, 'pic-m{0}-h{1:0.0}', $mode, $h)
            $d.SaveAs2((Join-Path $Out "$name.docx"), 16)
            $pages = $d.ComputeStatistics(2)
            $d.Close(0)
            '{0} {1}' -f $name, $pages
        }
    }
} finally { $word.Quit() }
