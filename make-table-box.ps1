# Synthetic reproducer: a floating text box holding a fixed-width table (green row borders,
# three inline pictures), Wrap text in shape off and Resize shape to fit text on, made with
# Word for Windows via COM. The saved size is then edited by hand to a narrow, stale one.
# All text and pictures are invented.
param([string]$Out = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Three plain synthetic pictures (a circle, a square, a triangle).
$pics = @()
foreach ($k in 0..2) {
    $bmp = New-Object Drawing.Bitmap 60, 60
    $g = [Drawing.Graphics]::FromImage($bmp); $g.Clear([Drawing.Color]::White)
    $brush = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(255, 30 + 80 * $k, 60, 160 - 50 * $k))
    switch ($k) {
        0 { $g.FillEllipse($brush, 5, 5, 50, 50) }
        1 { $g.FillRectangle($brush, 8, 8, 44, 44) }
        2 { $g.FillPolygon($brush, [Drawing.Point[]]@((New-Object Drawing.Point 30, 4), (New-Object Drawing.Point 56, 55), (New-Object Drawing.Point 4, 55))) }
    }
    $g.Dispose(); $p = Join-Path $Out "logo$k.png"; $bmp.Save($p, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose(); $pics += $p
}

$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
try {
    $doc = $w.Documents.Add()
    $doc.RemovePersonalInformation = $true   # no author in the files
    $doc.PageSetup.PaperSize = 7   # A4
    $doc.Content.Text = "Attestato di partecipazione (testo inventato). Il riquadro qui sotto contiene una tabella di tre colonne a larghezza fissa: due linee verdi e tre immagini."
    # Left 57.65 pt from the page, 500 pt from the top, as wide as the table for now.
    $s = $doc.Shapes.AddTextbox(1, 57.65, 500, 490, 80, $doc.Paragraphs(1).Range)
    $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1   # page
    $s.Left = 57.65; $s.Top = 500
    $s.Line.Visible = 0; $s.Fill.Visible = 0
    $tf = $s.TextFrame; $tf.MarginLeft = 5.4; $tf.MarginRight = 5.4; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $r = $tf.TextRange
    $t = $doc.Tables.Add($r, 3, 3)
    $t.AllowAutoFit = $false
    $widths = 170.85, 139.3, 168.25
    for ($c = 1; $c -le 3; $c++) { $t.Columns($c).Width = $widths[$c - 1] }
    $t.Borders.Enable = 0
    for ($c = 1; $c -le 3; $c++) {
        $b = $t.Cell(1, $c).Borders(-3)   # wdBorderBottom
        $b.LineStyle = 1; $b.LineWidth = 12; $b.Color = 0x008000   # BGR: green
        $b = $t.Cell(2, $c).Borders(-1)   # wdBorderTop
        $b.LineStyle = 1; $b.LineWidth = 12; $b.Color = 0x008000
        $t.Cell(3, $c).Range.ParagraphFormat.Alignment = 1
        $null = $t.Cell(3, $c).Range.InlineShapes.AddPicture($pics[$c - 1], $false, $true)
    }
    $t.Cell(1, 2).Range.Text = "ENTE FORMATORE"
    $t.Cell(1, 2).Range.ParagraphFormat.Alignment = 1
    $tf.WordWrap = 0      # Wrap text in shape off
    $tf.AutoSize = -1     # Resize shape to fit text
    $doc.SaveAs2((Join-Path $Out 'word-saved.docx'), 16)
    $doc.ExportAsFixedFormat((Join-Path $Out 'word-saved.pdf'), 17)
    $doc.Close(0)
} finally { $w.Quit() }
