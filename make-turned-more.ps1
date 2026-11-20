# Synthetic reproducer, part 2: turned text (btLr) in a vertical merge and in rows that split
# across pages. Invented text. Word's positions are read back through COM.
#   merge-<n>     : 3 rows; column 1 merged over the 3 rows with a long turned label; column 2
#                   has <n> lines per row; row heights auto.
#   split-auto    : 1 row; column 1 turned label; column 2 has 70 lines, so the row runs on
#                   to page 2; height auto.
#   split-atleast : 40 filler paragraphs, then a row at least 120 pt whose column 2 has 2 lines,
#                   placed so that the page ends inside the row's least height.
param([string] $Out = (Join-Path $PSScriptRoot 'more'))
New-Item -ItemType Directory -Force $Out | Out-Null
$label = 'TURNED LABEL OF SEVERAL INVENTED WORDS THAT IS LONGER THAN ONE ROW'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
function New-Doc([int] $filler) {
    $d = $word.Documents.Add()
    $d.RemovePersonalInformation = $true
    $d.Content.Text = (@(1..([Math]::Max(1, $filler)) | ForEach-Object { "Filler paragraph $_ before the table." }) -join "`r")
    $d.Content.InsertParagraphAfter()
    $d
}
function Set-Turned($cell) {
    $cell.Range.Text = $label
    $cell.Range.Font.Size = 9
    $cell.Range.Orientation = 2   # wdTextOrientationUpward (btLr)
    $cell.VerticalAlignment = 1
}
function Report($d, $t, $name) {
    $d.SaveAs2((Join-Path $Out "$name.docx"), 16)
    $d.ExportAsFixedFormat((Join-Path $Out "$name-word.pdf"), 17)
    $d.Repaginate()
    $top = $t.Range.Characters(1).Information(6)
    $topPage = $t.Range.Characters(1).Information(3)
    $last = $d.Paragraphs($d.Paragraphs.Count).Range
    '{0}: table from page {1} at {2:0.0}; next paragraph on page {3} at {4:0.0}; pages {5}' -f $name, $topPage, $top, $last.Information(3), $last.Information(6), $d.ComputeStatistics(2)
    $d.Close(0)
}
try {
    foreach ($n in 1, 3) {
        $d = New-Doc 1
        $t = $d.Tables.Add($d.Paragraphs($d.Paragraphs.Count).Range, 3, 2)
        $t.Borders.Enable = $true
        $t.Columns(1).Width = 40
        $t.Columns(2).Width = 300
        foreach ($r in 1..3) { $t.Cell($r, 2).Range.Text = ((1..$n | ForEach-Object { "Row $r line $_" }) -join "`r") }
        $t.Cell(1, 1).Merge($t.Cell(3, 1))
        Set-Turned $t.Cell(1, 1)
        $d.Content.InsertParagraphAfter()
        Report $d $t ('merge-{0}' -f $n)
    }
    foreach ($turned in $true, $false) {
        $suffix = if ($turned) { '' } else { '-upright' }
        $d = New-Doc 1
        $t = $d.Tables.Add($d.Paragraphs($d.Paragraphs.Count).Range, 1, 2)
        $t.Borders.Enable = $true; $t.Columns(1).Width = 40; $t.Columns(2).Width = 300
        Set-Turned $t.Cell(1, 1)
        if (-not $turned) { $t.Cell(1, 1).Range.Orientation = 0 }
        # Turning text through COM also sets cantSplit and an at-least height: undo both.
        $t.Rows.AllowBreakAcrossPages = -1; $t.Rows.HeightRule = 0
        $t.Cell(1, 2).Range.Text = ((1..70 | ForEach-Object { "Line $_ of the other cell" }) -join "`r")
        $d.Content.InsertParagraphAfter()
        Report $d $t "split-auto$suffix"
        foreach ($fill in 27, 28, 29) {
            $d = New-Doc $fill
            $t = $d.Tables.Add($d.Paragraphs($d.Paragraphs.Count).Range, 1, 2)
            $t.Borders.Enable = $true; $t.Columns(1).Width = 40; $t.Columns(2).Width = 300
            Set-Turned $t.Cell(1, 1)
            if (-not $turned) { $t.Cell(1, 1).Range.Orientation = 0 }
            $t.Cell(1, 2).Range.Text = "Line 1 of the other cell`rLine 2 of the other cell"
            $t.Rows.AllowBreakAcrossPages = -1
            $t.Rows(1).HeightRule = 1; $t.Rows(1).Height = 120
            $d.Content.InsertParagraphAfter()
            Report $d $t "split-atleast-$fill$suffix"
        }
    }
} finally { $word.Quit() }
