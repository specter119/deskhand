# Live old-vs-new fingerprint golden for the PowerPoint bridge.
#
# Run from WSL:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File \\wsl.localhost\<distro>\<repo>\tests\live\ppt_fingerprint_golden.ps1
#
# It creates a temporary deck, exercises text/table/chart/picture/notes and slide
# duplicate/move/delete, then compares Get-*FingerprintLegacy with the snapshot readers in the
# same PowerPoint process. It never touches an existing deck and quits PowerPoint only when it
# started PowerPoint itself.

param(
    [string]$BridgePath = (Join-Path $PSScriptRoot '..\..\src\deskhand\ppt\bridge.ps1')
)

$ErrorActionPreference = 'Stop'

$source = [IO.File]::ReadAllText($BridgePath, [Text.Encoding]::UTF8)
$marker = $source.IndexOf('# --- main')
if ($marker -lt 0) { throw "Could not find the main marker in $BridgePath" }

$tmp = Join-Path ([IO.Path]::GetTempPath()) ('deskhand-ppt-golden-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$functionsPath = Join-Path $tmp 'functions.ps1'
[IO.File]::WriteAllText($functionsPath, $source.Substring(0, $marker), [Text.UTF8Encoding]::new($false))
. $functionsPath

Add-Type -AssemblyName System.Drawing
$png = Join-Path $tmp 'pixel.png'
$bitmap = New-Object Drawing.Bitmap 16, 16
$graphics = [Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([Drawing.Color]::CornflowerBlue)
$bitmap.Save($png, [Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()

$started = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -eq 0
$app = $null
$pres = $null
$failures = @()
$result = @{}

function Add-Text($slide, [string]$text, [double]$x, [double]$y, [double]$w, [double]$h) {
    $shape = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $shape.TextFrame.TextRange.Text = $text
    return $shape
}

try {
    $app = New-Object -ComObject PowerPoint.Application
    if ($started) { $app.Visible = -1 }
    $pres = $app.Presentations.Add($false)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $slide = $pres.Slides.Add(1, 12)
    Add-Text $slide "alpha`rbeta" 40 40 300 80 | Out-Null
    $tableShape = $slide.Shapes.AddTable(2, 3, 40, 150, 360, 100)
    $tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'h1'
    $tableShape.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = 'h2'
    $tableShape.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = 'a'
    $tableShape.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = 'b'
    $chartShape = $slide.Shapes.AddChart2(-1, 51, 40, 270, 360, 180)
    $chart = $chartShape.Chart
    $chart.ChartData.Activate()
    $sheet = $chart.ChartData.Workbook.Worksheets.Item(1)
    $sheet.Cells.Item(2, 1).Value2 = 'Q1'
    $sheet.Cells.Item(2, 2).Value2 = 10
    $sheet.Cells.Item(3, 1).Value2 = 'Q2'
    $sheet.Cells.Item(3, 2).Value2 = 20
    $chart.ChartData.Workbook.Close()
    $slide.Shapes.AddPicture($png, 0, -1, 430, 40, 100, 100) | Out-Null
    $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'speaker note'

    $duplicate = $slide.Duplicate()
    $duplicate.MoveTo(1)
    $pres.Slides.Add(2, 12) | Out-Null
    $pres.Slides.Item(3).Delete()

    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $current = $pres.Slides.Item($i)
        $old = Get-SlideFingerprintLegacy $current
        $new = (Get-SlideSnapshot $current $true $true).fingerprint
        if ($old -ne $new) { $failures += "slide $i old=$old new=$new" }
        $withoutNotes = (Get-SlideSnapshot $current $false $false).fingerprint
        if ($withoutNotes -ne $new) { $failures += "slide $i notes changed the fingerprint" }
    }

    $oldDeck = Get-DeckFingerprintLegacy $pres
    $newDeck = (Get-DeckSnapshot $pres $true).fingerprint
    if ($oldDeck -ne $newDeck) { $failures += "deck old=$oldDeck new=$newDeck" }

    $perfSlide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    for ($i = 1; $i -le 100; $i++) {
        Add-Text $perfSlide "shape $i" (($i % 10) * 80) ([math]::Floor($i / 10) * 40) 70 30 | Out-Null
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $null = Get-SlideSnapshot $perfSlide $true $true
    $sw.Stop()
    $slideMs = $sw.ElapsedMilliseconds
    if ($slideMs -ge 1500) { $failures += "100-shape read_slide took $slideMs ms" }

    for ($i = 1; $i -le 8; $i++) {
        $perf = $pres.Slides.Add($pres.Slides.Count + 1, 12)
        for ($j = 1; $j -le 18; $j++) {
            Add-Text $perf "deck $i/$j" (($j % 6) * 140) ([math]::Floor($j / 6) * 60) 120 40 | Out-Null
        }
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $null = Get-DeckSnapshot $pres $true
    $sw.Stop()
    $deckSnapshotMs = $sw.ElapsedMilliseconds
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $null = Get-DeckFingerprintLegacy $pres
    for ($i = 1; $i -le $pres.Slides.Count; $i++) { $null = Get-SlideModel ($pres.Slides.Item($i)) $true }
    $sw.Stop()
    $deckLegacyMs = $sw.ElapsedMilliseconds
    if ($deckSnapshotMs -gt $deckLegacyMs) {
        $failures += "read_deck snapshot ${deckSnapshotMs}ms is slower than legacy ${deckLegacyMs}ms"
    }

    $result = @{
        ok = ($failures.Count -eq 0)
        failures = $failures
        slides = $pres.Slides.Count
        slide100Ms = $slideMs
        deckSnapshotMs = $deckSnapshotMs
        deckLegacyMs = $deckLegacyMs
    }
}
finally {
    if ($null -ne $pres) {
        try { $pres.Saved = $true; $pres.Close() } catch { }
    }
    if ($started -and $null -ne $app) { try { $app.Quit() } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    try { Remove-Item -LiteralPath $tmp -Recurse -Force } catch { }
}

$result | ConvertTo-Json -Depth 5
if (-not $result.ok) { exit 1 }
