# From word-saved.docx (made by make-table-box.ps1): the box's saved width is set by hand to a
# stale 19.85 pt (wp:extent, a:ext and the VML fallback), as in a file saved by another Word,
# in four variants of Wrap text in shape (wrap) and Resize shape to fit text (autofit).
# Each variant is exported to PDF by Word and rendered by the WordCraft CLI given.
param([string]$Dir = $PSScriptRoot, [string]$Cli, [string]$Tag = 'wc')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem, System.Drawing
$src = Join-Path $Dir 'word-saved.docx'
$variants = [ordered]@{
    'wide-none-autofit' = @{ wrap = 'none';   fit = '<a:spAutoFit/>'; stale = $false }
    'none-autofit'      = @{ wrap = 'none';   fit = '<a:spAutoFit/>'; stale = $true }
    'none-fixed'        = @{ wrap = 'none';   fit = '<a:noAutofit/>'; stale = $true }
    'square-autofit'    = @{ wrap = 'square'; fit = '<a:spAutoFit/>'; stale = $true }
    'square-fixed'      = @{ wrap = 'square'; fit = '<a:noAutofit/>'; stale = $true }
}
function Set-Variant($out, $v) {
    Copy-Item $src $out -Force
    $z = [IO.Compression.ZipFile]::Open($out, 'Update'); $e = $z.GetEntry('word/document.xml')
    $r = New-Object IO.StreamReader($e.Open()); $t = $r.ReadToEnd(); $r.Close()
    # The table's preferred width is automatic (its columns keep their fixed widths).
    $t = $t.Replace('<w:tblW w:w="9568" w:type="dxa"/>', '<w:tblW w:w="0" w:type="auto"/>')
    # AutoFit layout, as Word writes a table whose AutoFit was never turned off.
    $t = $t.Replace('<w:tblLayout w:type="fixed"/>', '')
    if ($v.stale) { $t = $t.Replace('cx="6223000"', 'cx="252095"').Replace('width:490pt', 'width:19.85pt') }
    $t = $t.Replace('wrap="none"', "wrap=""$($v.wrap)""").Replace('<a:spAutoFit/>', $v.fit)
    if ($v.wrap -ne 'none') { $t = $t.Replace('mso-wrap-style:none', 'mso-wrap-style:square') }
    if ($v.fit -ne '<a:spAutoFit/>') { $t = $t.Replace('mso-fit-shape-to-text:t', 'mso-fit-shape-to-text:f') }
    $e.Delete(); $n = $z.CreateEntry('word/document.xml')
    $w = New-Object IO.StreamWriter($n.Open(), (New-Object Text.UTF8Encoding $false)); $w.Write($t); $w.Close(); $z.Dispose()
}
# Green row border and picture span on page 1 at 72 dpi, x in pt.
function Measure-Png($png) {
    $b = [Drawing.Bitmap]::FromFile($png); $gmin = 9999; $gmax = -1; $green = 0; $pmin = 9999; $pmax = -1; $pics = 0
    for ($y = 480; $y -lt 640; $y++) { for ($x = 0; $x -lt $b.Width; $x++) {
        $c = $b.GetPixel($x, $y)
        if ($c.G -gt 90 -and $c.R -lt 80 -and $c.B -lt 80) { $green++; $gmin = [Math]::Min($gmin, $x); $gmax = [Math]::Max($gmax, $x) }
        elseif (([Math]::Max($c.R, $c.B) - $c.G) -gt 40) { $pics++; $pmin = [Math]::Min($pmin, $x); $pmax = [Math]::Max($pmax, $x) }
    } }
    $b.Dispose()
    [pscustomobject]@{ green = $green; greenX = "$gmin-$gmax"; picPx = $pics; picX = "$pmin-$pmax" }
}
$gs = Get-ChildItem 'C:\Program Files\gs' -Recurse -Filter gswin64c.exe | Select-Object -First 1 -ExpandProperty FullName
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0
$rows = @()
try {
    foreach ($k in @('saved') + $variants.Keys) {
        $docx = Join-Path $Dir "$k.docx"
        if ($k -eq 'saved') { Copy-Item $src $docx -Force } else { Set-Variant $docx $variants[$k] }
        $wpdf = Join-Path $Dir "$k.word.pdf"
        if (-not (Test-Path $wpdf)) { $d = $word.Documents.Open($docx, $false, $true, $false); $d.ExportAsFixedFormat($wpdf, 17); $d.Close(0) }
        & $gs -q -dNOPAUSE -dBATCH -sDEVICE=png16m -r72 -dFirstPage=1 -dLastPage=1 -o (Join-Path $Dir "$k.word.png") $wpdf
        if ($Cli) { $null = & $Cli render $docx (Join-Path $Dir "$k.$Tag.png") --page 1 --scale 1 }
        $mw = Measure-Png (Join-Path $Dir "$k.word.png")
        $row = [ordered]@{ variant = $k; 'Word green' = "$($mw.green) x $($mw.greenX)"; 'Word pics' = "$($mw.picPx) x $($mw.picX)" }
        if ($Cli) { $mc = Measure-Png (Join-Path $Dir "$k.$Tag.png"); $row["$Tag green"] = "$($mc.green) x $($mc.greenX)"; $row["$Tag pics"] = "$($mc.picPx) x $($mc.picX)" }
        $rows += [pscustomobject]$row
    }
} finally { $word.Quit() }
$rows | Format-Table -AutoSize | Out-String -Width 200
