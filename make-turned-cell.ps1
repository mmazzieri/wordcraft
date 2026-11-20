# Synthetic reproducer: a two-column table row whose narrow first cell holds turned text (btLr)
# longer than the row the second cell makes. Variants: row height auto, or at least 78 pt;
# N lines in the second cell. Invented text. Word's row height is read back through COM.
param([string] $Out = $PSScriptRoot, [int[]] $Lines = @(3, 6, 10), [string[]] $Rules = @('auto', 'atleast'))
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($rule in $Rules) {
        foreach ($n in $Lines) {
            $d = $word.Documents.Add()
            $d.RemovePersonalInformation = $true
            $d.Content.Text = 'Before the table.'
            $r = $d.Paragraphs(1).Range
            $r.InsertParagraphAfter()
            $t = $d.Tables.Add($d.Paragraphs(2).Range, 1, 2)
            $t.Borders.Enable = $true
            $t.Columns(1).Width = 40
            $t.Columns(2).Width = 300
            $c1 = $t.Cell(1, 1)
            $c1.Range.Text = 'TURNED LABEL OF SEVERAL INVENTED WORDS'
            $c1.Range.Font.Size = 9
            $c1.Range.Orientation = 2   # wdTextOrientationUpward (btLr)
            $c1.VerticalAlignment = 1
            $c2 = $t.Cell(1, 2)
            $c2.Range.Text = ((1..$n | ForEach-Object { "Line $_ of the other cell" }) -join "`r")
            if ($rule -eq 'atleast') { $t.Rows(1).HeightRule = 1; $t.Rows(1).Height = 78 }
            $name = 'turned-{0}-{1}' -f $rule, $n
            $d.SaveAs2((Join-Path $Out "$name.docx"), 16)
            $d.ExportAsFixedFormat((Join-Path $Out "$name-word.pdf"), 17)
            $d.Repaginate()
            $top = $t.Rows(1).Range.Information(6)   # wdVerticalPositionRelativeToPage
            $after = $d.Paragraphs($d.Paragraphs.Count).Range.Information(6)
            $d.Close(0)
            '{0}: row from {1:0.0} to the next paragraph at {2:0.0} ({3:0.0} pt)' -f $name, $top, $after, ($after - $top)
        }
    }
} finally { $word.Quit() }
