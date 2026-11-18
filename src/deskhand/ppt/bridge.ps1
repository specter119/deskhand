# PowerPoint bridge for the deskhand `ppt` server. One powershell.exe runs this file for the whole life
# of the server process, started through the same tiny encoded bootstrap as the OneNote bridge.
#
#   stdin   one call stem per line
#   stdout  "deskhand-ready 1 <windows pid>" once, then "deskhand-done <stem>" after each call
#   files   req-<stem>.json in and resp-<stem>.json out, both UTF-8, inside $d
#
# Never use `exit` anywhere in this file: it would end the worker between writing a response and printing
# "deskhand-done", and the caller would take a finished write for a lost one. Stop a call with Fail or return.
#
# Write contract: every write op is default-deny. It only writes when the request carries apply=true AND a
# non-empty expectFingerprint. The fingerprint is recomputed here, immediately before the COM write, from
# the live slide or deck, and a mismatch is refused with code changed_since_preview. Without apply=true the
# op is a preview and returns the fingerprint that a later apply must present.
#
# Process contract: the worker only closes presentations it opened or created, and only quits PowerPoint
# when it started PowerPoint itself and no presentation is left open. A user's running PowerPoint is never
# quit and a user's open deck is never closed.

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:RESP = $null
$script:RESPONDED = $false
$script:MUTATED = $false
$script:CURRENT_SHAPE_MUTATED = $false
$script:SHAPE_ERROR_MODE = $false
$script:CURRENT_SHAPE_INDEX = $null
$script:STEM = ''
$script:PptApp = $null
$script:OURS = @()
$script:OUR_WINDOWS = @()
$script:THEMES = @{}
$script:STARTED_PPT = $false
$script:RESET_PENDING = $false
$script:RESET_ATTEMPTS = 0
$script:RESET_TERMINAL = $false
$script:PPT_PID = 0
$script:PPT_STARTED_AT = $null
$script:PPT_MAPPING_FAILED = $false

if (-not ('DeskhandPptNative' -as [type])) {
	Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class DeskhandPptNative {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    public static List<IntPtr> FindPowerPointWindows() {
        var found = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            var name = new StringBuilder(256);
            GetClassName(hWnd, name, name.Capacity);
            var cls = name.ToString();
            if (cls == "PPTFrameClass" || cls == "PP12FrameClass") { found.Add(hWnd); }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@
}

function Write-Response($payload) {
	if ($payload.ok -and $payload.data -is [hashtable] -and $payload.data.ContainsKey('fingerprint') -and -not $payload.data.ContainsKey('fingerprintQuality') -and $null -ne $script:LAST_FINGERPRINT_QUALITY) {
		$payload.data.fingerprintQuality = $script:LAST_FINGERPRINT_QUALITY
	}
	$json = $payload | ConvertTo-Json -Depth 40 -Compress
	[IO.File]::WriteAllText($script:RESP, $json, [System.Text.UTF8Encoding]::new($false))
	$script:RESPONDED = $true
}

function Fail([string]$message, [string]$code, $data = $null) {
	if (Test-Mutated) {
		$code = 'applied_unverified'
		$message = 'Part or all of this write was applied; do not retry, read the deck again. ' + $message
		if ($null -eq $data) {
			$data = @{ status = 'unknown'; tokenSpent = $true; appliedCount = $null; failedIndex = $null; shapeIds = $null; failedShapeMayExist = $null }
		}
		elseif ($data.status -eq 'invalid') {
			$data.status = 'unknown'
		}
	}
	$payload = @{ ok = $false; error = $message; code = $code }
	if ($null -ne $data) { $payload.data = $data }
	Write-Response $payload
	throw 'deskhand:responded'
}

function New-ShapeError([string]$message, [string]$code) {
	return @{
		deskhand_shape_error = $true
		message = $message
		code = $code
		failedIndex = $script:CURRENT_SHAPE_INDEX
	}
}

function Fail-Shape([string]$message, [string]$code) {
	if ($script:SHAPE_ERROR_MODE) { throw (New-ShapeError $message $code) }
	Fail $message $code
}

function Test-Mutated {
	return ($script:MUTATED -or $script:CURRENT_SHAPE_MUTATED)
}

function Get-Field($bag, [string]$name) {
	if ($null -eq $bag) { return $null }
	if ($bag -is [hashtable]) {
		if ($bag.ContainsKey($name)) { return $bag[$name] }
		return $null
	}
	$prop = $bag.PSObject.Properties[$name]
	if ($null -eq $prop) { return $null }
	return $prop.Value
}

function Get-Array($bag, [string]$name) {
	$value = Get-Field $bag $name
	if ($null -eq $value) { return ,@() }
	return ,@($value)
}

function Test-Apply($req) {
	return ([bool](Get-Field $req 'apply'))
}

function Assert-Fingerprint($req, [string]$actual) {
	if ($script:LAST_FINGERPRINT_QUALITY -eq 'unavailable') {
		Fail 'The deck could not be read completely, so no fingerprint is available. Nothing was written.' 'fingerprint_unavailable'
	}
	if (-not (Test-Apply $req)) { return }
	$expected = [string](Get-Field $req 'expectFingerprint')
	if ([string]::IsNullOrWhiteSpace($expected)) {
		Fail 'apply requires expectFingerprint from the preview' 'apply_requires_fingerprint'
	}
	if ($expected -ne $actual) {
		Fail 'The deck changed since the preview; read it again before retrying.' 'changed_since_preview'
	}
}

function Get-Sha256Hex([string]$text) {
	$sha = [Security.Cryptography.SHA256]::Create()
	$bytes = [Text.Encoding]::UTF8.GetBytes($text)
	return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
}

function Get-Num($value, [double]$default) {
	if ($null -eq $value) { return $default }
	return [double]$value
}

function Format-Num([double]$value) {
	return $value.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture)
}

function PtToPx([double]$pt) {
	return [math]::Round($pt * 4.0 / 3.0, 1)
}

function PxToPt([double]$px) {
	return $px * 0.75
}

function ConvertTo-Color([string]$hex) {
	if ([string]::IsNullOrWhiteSpace($hex)) { return $null }
	$h = $hex.Trim()
	if ($h.StartsWith('#')) { $h = $h.Substring(1) }
	if ($h.Length -ne 6) { Fail-Shape ('color must be #RRGGBB: ' + $hex) 'invalid_request' }
	try {
		$r = [Convert]::ToInt32($h.Substring(0, 2), 16)
		$g = [Convert]::ToInt32($h.Substring(2, 2), 16)
		$b = [Convert]::ToInt32($h.Substring(4, 2), 16)
	}
	catch { Fail-Shape ('color must be #RRGGBB: ' + $hex) 'invalid_request' }
	return ($r + ($g * 256) + ($b * 65536))
}

function Set-Fill($shape, [string]$hex) {
	$rgb = ConvertTo-Color $hex
	if ($null -eq $rgb) { return }
	$shape.Fill.Solid()
	$shape.Fill.ForeColor.RGB = $rgb
}

function Set-Line($shape, [string]$hex) {
	$rgb = ConvertTo-Color $hex
	if ($null -eq $rgb) { return }
	$shape.Line.Visible = -1
	$shape.Line.ForeColor.RGB = $rgb
	$shape.Line.Weight = 1
}

function Get-ShapeTypeName($type) {
	switch ([int]$type) {
		1 { return 'autoShape' }
		3 { return 'chart' }
		6 { return 'group' }
		11 { return 'linkedPicture' }
		13 { return 'picture' }
		14 { return 'placeholder' }
		17 { return 'textBox' }
		19 { return 'table' }
		20 { return 'media' }
		default { return 'type:' + [string]$type }
	}
}

function Get-ChartTypeName($type) {
	switch ([int]$type) {
		4 { return 'line' }
		5 { return 'pie' }
		51 { return 'column' }
		57 { return 'bar' }
		default { return 'type:' + [string]$type }
	}
}

function Get-AlignName($value) {
	switch ([int]$value) {
		1 { return 'left' }
		2 { return 'center' }
		3 { return 'right' }
		4 { return 'justify' }
		default { return '' }
	}
}

function Get-ShapeParagraphs($shape) {
	$out = @()
	try {
		if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
			$tr = $shape.TextFrame.TextRange
			$count = [int]$tr.Paragraphs().Count
			for ($p = 1; $p -le $count; $p++) {
				$para = $tr.Paragraphs($p)
				$runs = @()
				$runCount = [int]$para.Runs().Count
				for ($r = 1; $r -le $runCount; $r++) {
					$run = $para.Runs($r)
					$color = ''
					try { $color = Get-ColorHex $run.Font.Color.RGB } catch { }
					$runs += @{
						text   = [string]$run.Text
						bold   = ([int]$run.Font.Bold -ne 0)
						italic = ([int]$run.Font.Italic -ne 0)
						color  = $color
						size   = (PtToPx ([double]$run.Font.Size))
						font   = [string]$run.Font.Name
					}
				}
				$out += @{
					runs        = $runs
					align       = (Get-AlignName $para.ParagraphFormat.Alignment)
					spaceBefore = (PtToPx ([double]$para.ParagraphFormat.SpaceBefore))
					spaceAfter  = (PtToPx ([double]$para.ParagraphFormat.SpaceAfter))
				}
			}
		}
	}
	catch { }
	return ,$out
}

function Get-ShapeTextStrict($shape) {
	try {
		if ($shape.HasTextFrame -ne -1) { return @{ ok = $true; value = '' } }
		if ($shape.TextFrame.HasText -ne -1) { return @{ ok = $true; value = '' } }
		$text = [string]$shape.TextFrame.TextRange.Text
		return @{ ok = $true; value = $text.Replace("`r", "`n").Replace("`v", "`n").TrimEnd("`n") }
	}
	catch { return @{ ok = $false; value = '' } }
}

function Get-ShapeText($shape) {
	return (Get-ShapeTextStrict $shape).value
}

function Get-ShapeNameStrict($shape) {
	try { return @{ ok = $true; value = [string]$shape.Name } }
	catch { return @{ ok = $false; value = '' } }
}

function Get-ShapeParagraphsStrict($shape) {
	try {
		if ($shape.HasTextFrame -ne -1) { return @{ ok = $true; value = $null } }
		if ($shape.TextFrame.HasText -ne -1) { return @{ ok = $true; value = @() } }
		$out = @()
		$tr = $shape.TextFrame.TextRange
		$count = [int]$tr.Paragraphs().Count
		for ($p = 1; $p -le $count; $p++) {
			$para = $tr.Paragraphs($p)
			$runs = @()
			$runCount = [int]$para.Runs().Count
			for ($r = 1; $r -le $runCount; $r++) {
				$run = $para.Runs($r)
				$color = ''
				try { $color = Get-ColorHex $run.Font.Color.RGB } catch { }
				$runs += @{
					text   = [string]$run.Text
					bold   = ([int]$run.Font.Bold -ne 0)
					italic = ([int]$run.Font.Italic -ne 0)
					color  = $color
					size   = (PtToPx ([double]$run.Font.Size))
					font   = [string]$run.Font.Name
				}
			}
			$out += @{
				runs        = $runs
				align       = (Get-AlignName $para.ParagraphFormat.Alignment)
				spaceBefore = (PtToPx ([double]$para.ParagraphFormat.SpaceBefore))
				spaceAfter  = (PtToPx ([double]$para.ParagraphFormat.SpaceAfter))
			}
		}
		return @{ ok = $true; value = $out }
	}
	catch { return @{ ok = $false; value = @() } }
}

function Get-ShapeFillStrict($shape) {
	try {
		if ($shape.Fill.Visible -ne -1) { return @{ ok = $true; value = $null } }
		return @{ ok = $true; value = (Get-ColorHex $shape.Fill.ForeColor.RGB) }
	}
	catch { return @{ ok = $false; value = $null } }
}

function Get-ShapeLineStrict($shape) {
	try {
		if ($shape.Line.Visible -ne -1) { return @{ ok = $true; value = $null } }
		return @{ ok = $true; value = @{ color = (Get-ColorHex $shape.Line.ForeColor.RGB); weight = [double]$shape.Line.Weight } }
	}
	catch { return @{ ok = $false; value = $null } }
}

function Get-ShapeTableStrict($shape) {
	try {
		if ($shape.HasTable -ne -1) { return @{ ok = $true; value = $null } }
		$table = $shape.Table
		$rows = @()
		$rowCount = [int]$table.Rows.Count
		$colCount = [int]$table.Columns.Count
		for ($r = 1; $r -le $rowCount; $r++) {
			$row = @()
			for ($c = 1; $c -le $colCount; $c++) {
				$cell = Get-ShapeTextStrict $table.Cell($r, $c).Shape
				if (-not $cell.ok) { return @{ ok = $false; value = $null } }
				$row += $cell.value
			}
			$rows += , $row
		}
		return @{ ok = $true; value = @{ rows = $rowCount; cols = $colCount; cells = $rows } }
	}
	catch { return @{ ok = $false; value = $null } }
}

function Get-ShapeChartStrict($shape) {
	try {
		if ($shape.HasChart -ne -1) { return @{ ok = $true; value = $null } }
		$chart = $shape.Chart
		$names = @()
		$categories = @()
		$seriesModels = @()
		$count = [int]$chart.SeriesCollection().Count
		for ($i = 1; $i -le $count; $i++) {
			$series = $chart.SeriesCollection($i)
			$names += [string]$series.Name
			$seriesModels += @{ name = [string]$series.Name; values = @($series.Values); xValues = @($series.XValues) }
			if ($i -eq 1) { $categories = @($series.XValues) }
		}
		return @{ ok = $true; value = @{ chartType = (Get-ChartTypeName $chart.ChartType); seriesNames = $names; categories = $categories; series = $seriesModels } }
	}
	catch { return @{ ok = $false; value = $null } }
}

function Get-ShapePictureStrict($shape) {
	try {
		$type = [int]$shape.Type
		if ($type -ne 13 -and $type -ne 11) { return @{ ok = $true; value = $null } }
		return @{ ok = $true; value = @{ path = [string]$shape.LinkFormat.SourceFullName; widthPx = (PtToPx ([double]$shape.Width)); heightPx = (PtToPx ([double]$shape.Height)) } }
	}
	catch { return @{ ok = $false; value = $null } }
}

function Get-NotesTextStrict($slide) {
	try {
		$placeholders = $slide.NotesPage.Shapes.Placeholders
		if ([int]$placeholders.Count -lt 2) { return @{ ok = $true; value = '' } }
		$ph = $placeholders.Item(2)
		if ($ph.HasTextFrame -ne -1) { return @{ ok = $true; value = '' } }
		if ($ph.TextFrame.HasText -ne -1) { return @{ ok = $true; value = '' } }
		$text = [string]$ph.TextFrame.TextRange.Text
		return @{ ok = $true; value = $text.Replace("`r", "`n").Replace("`v", "`n").TrimEnd("`n") }
	}
	catch { return @{ ok = $false; value = '' } }
}

function Get-NotesText($slide) {
	return (Get-NotesTextStrict $slide).value
}

function Get-SlideLayoutStrict($slide) {
	try { return @{ ok = $true; value = [string]$slide.CustomLayout.Name } }
	catch { return @{ ok = $false; value = '' } }
}

function Get-SlideTitleStrict($slide) {
	try {
		$count = [int]$slide.Shapes.Count
		for ($i = 1; $i -le $count; $i++) {
			$shape = $slide.Shapes.Item($i)
			if ([int]$shape.Type -eq 14) {
				$t = 0
				try { $t = [int]$shape.PlaceholderFormat.Type } catch { }
				if ($t -eq 1 -or $t -eq 3) {
					$text = Get-ShapeTextStrict $shape
					if (-not $text.ok) { return @{ ok = $false; value = '' } }
					return @{ ok = $true; value = $text.value }
				}
			}
		}
		return @{ ok = $true; value = '' }
	}
	catch { return @{ ok = $false; value = '' } }
}

function Get-SlideTitle($slide) {
	return (Get-SlideTitleStrict $slide).value
}

function Get-SlideShapesStrict($slide) {
	$out = @()
	try {
		$count = [int]$slide.Shapes.Count
		for ($i = 1; $i -le $count; $i++) { $out += $slide.Shapes.Item($i) }
		return @{ ok = $true; value = $out }
	}
	catch { return @{ ok = $false; value = $out } }
}

function Merge-FingerprintQuality([string]$left, [string]$right) {
	if ($left -eq 'unavailable' -or $right -eq 'unavailable') { return 'unavailable' }
	if ($left -eq 'coarse' -or $right -eq 'coarse') { return 'coarse' }
	return 'full'
}

function Get-ShapeSnapshot($shape, [bool]$includeModel) {
	$coreOk = $true
	$id = -1
	$typeRaw = 0
	$left = 0
	$top = 0
	$width = 0
	$height = 0
	try { $id = [int]$shape.Id } catch { $coreOk = $false }
	try { $typeRaw = [int]$shape.Type } catch { $coreOk = $false }
	try { $left = [double]$shape.Left } catch { $coreOk = $false }
	try { $top = [double]$shape.Top } catch { $coreOk = $false }
	try { $width = [double]$shape.Width } catch { $coreOk = $false }
	try { $height = [double]$shape.Height } catch { $coreOk = $false }
	$text = Get-ShapeTextStrict $shape
	if (-not $text.ok) { $coreOk = $false }
	if (-not $coreOk) {
		$unavailable = @{ id = -1; name = ''; type = 'unavailable'; x = 0; y = 0; w = 0; h = 0; text = '' }
		return @{ model = $(if ($includeModel) { $unavailable } else { $null }); fingerprintParts = $null; quality = 'unavailable' }
	}

	$name = Get-ShapeNameStrict $shape
	$paragraphs = Get-ShapeParagraphsStrict $shape
	$fill = Get-ShapeFillStrict $shape
	$line = Get-ShapeLineStrict $shape
	$table = Get-ShapeTableStrict $shape
	$chart = Get-ShapeChartStrict $shape
	$picture = Get-ShapePictureStrict $shape
	$type = Get-ShapeTypeName $typeRaw
	$quality = 'full'
	if (-not $name.ok -or -not $paragraphs.ok -or -not $fill.ok -or -not $line.ok -or -not $table.ok -or -not $chart.ok -or -not $picture.ok) { $quality = 'coarse' }
	if ($type -like 'type:*') { $quality = 'coarse' }

	$parts = @(
		[string]$id,
		$(if ($name.ok) { [string]$name.value } else { 'name|?' }),
		$type,
		(Format-Num (PtToPx $left)),
		(Format-Num (PtToPx $top)),
		(Format-Num (PtToPx $width)),
		(Format-Num (PtToPx $height)),
		('text|' + [string]$text.value)
	)
	if (-not $fill.ok) { $parts += 'fill|?' }
	elseif ($null -eq $fill.value) { $parts += 'fill|none' }
	else { $parts += 'fill|' + [string]$fill.value }
	if (-not $line.ok) { $parts += 'line|?' }
	elseif ($null -eq $line.value) { $parts += 'line|none' }
	else { $parts += 'line|' + [string]$line.value.color + '|' + (Format-Num ([double]$line.value.weight)) }
	if (-not $table.ok) { $parts += 'table|?' }
	elseif ($null -ne $table.value) {
		$rows = @($table.value.cells)
		for ($r = 0; $r -lt $rows.Count; $r++) {
			$row = @($rows[$r])
			for ($c = 0; $c -lt $row.Count; $c++) {
				$parts += 'cell|' + ($r + 1) + '|' + ($c + 1) + '|' + [string]$row[$c]
			}
		}
	}
	if (-not $chart.ok) { $parts += 'chart|?' }
	elseif ($null -ne $chart.value) {
		foreach ($series in @($chart.value.series)) {
			$parts += 'series|' + [string]$series.name + '|' + (($series.values) -join ',') + '|' + (($series.xValues) -join ',')
		}
	}
	if (-not $picture.ok) { $parts += 'picture|?' }
	elseif ($null -ne $picture.value) { $parts += 'picture|' + [string]$picture.value.path }

	$model = $null
	if ($includeModel) {
		$model = @{
			id   = $id
			name = [string]$name.value
			type = $type
			x    = PtToPx $left
			y    = PtToPx $top
			w    = PtToPx $width
			h    = PtToPx $height
			text = [string]$text.value
		}
		if ($paragraphs.ok -and $null -ne $paragraphs.value -and @($paragraphs.value).Count -gt 0) { $model.paragraphs = @($paragraphs.value) }
		if ($table.ok -and $null -ne $table.value) { $model.table = @{ rows = $table.value.rows; cols = $table.value.cols; cells = $table.value.cells } }
		if ($chart.ok -and $null -ne $chart.value) {
			$seriesModels = @()
			foreach ($series in @($chart.value.series)) { $seriesModels += @{ name = $series.name; values = $series.values } }
			$model.chart = @{ chartType = $chart.value.chartType; seriesNames = $chart.value.seriesNames; categories = $chart.value.categories; series = $seriesModels }
		}
		if ($picture.ok -and $null -ne $picture.value) { $model.picture = $picture.value }
	}
	return @{ model = $model; fingerprintParts = $parts; quality = $quality }
}

function Get-ShapeModel($shape) {
	return (Get-ShapeSnapshot $shape $true).model
}

function Get-SlideSnapshot($slide, [bool]$includeModel, [bool]$includeNotes = $false) {
	$layout = Get-SlideLayoutStrict $slide
	$title = Get-SlideTitleStrict $slide
	$notes = Get-NotesTextStrict $slide
	$shapesRead = Get-SlideShapesStrict $slide
	$shapeModels = @()
	$shapeParts = @()
	$quality = 'full'
	foreach ($shape in @($shapesRead.value)) {
		$snapshot = Get-ShapeSnapshot $shape $includeModel
		if ($includeModel) { $shapeModels += $snapshot.model }
		if ($null -eq $snapshot.fingerprintParts) { $quality = 'unavailable' }
		else {
			$shapeParts += , $snapshot.fingerprintParts
			$quality = Merge-FingerprintQuality $quality $snapshot.quality
		}
	}
	if (-not $layout.ok -or -not $title.ok -or -not $notes.ok -or -not $shapesRead.ok) { $quality = 'unavailable' }
	$fingerprint = $null
	if ($quality -ne 'unavailable') {
		$parts = @('slide', [string]$layout.value)
		foreach ($shapePart in $shapeParts) { $parts += ($shapePart -join '|') }
		$parts += 'notes|' + [string]$notes.value
		$fingerprint = Get-Sha256Hex ($parts -join '|')
	}
	$model = $null
	if ($includeModel) {
		$index = 0
		try { $index = [int]$slide.SlideIndex } catch { $quality = 'unavailable'; $fingerprint = $null }
		$model = @{ index = $index; layout = [string]$layout.value; title = [string]$title.value; shapes = $shapeModels }
		if ($includeNotes) { $model.notes = [string]$notes.value }
	}
	return @{ model = $model; fingerprint = $fingerprint; quality = $quality }
}

function Get-SlideModel($slide, [bool]$includeNotes) {
	return (Get-SlideSnapshot $slide $true $includeNotes).model
}

function Get-DeckSnapshot($pres, [bool]$includeSlides, $indexes = $null, [bool]$includeNotes = $false) {
	$quality = 'full'
	$fingerprint = $null
	$slides = @()
	$count = 0
	$countOk = $true
	try { $count = [int]$pres.Slides.Count } catch { $countOk = $false; $quality = 'unavailable' }
	$parts = @('deck', [string]$count)
	$selected = @{}
	if ($includeSlides) {
		if ($null -eq $indexes) { for ($i = 1; $i -le $count; $i++) { $selected[$i] = $true } }
		else { foreach ($i in @($indexes)) { $selected[[int]$i] = $true } }
	}
	for ($i = 1; $i -le $count; $i++) {
		try {
			$slide = $pres.Slides.Item($i)
			$parts += [string]$slide.SlideID
			$wantModel = $includeSlides -and $selected.ContainsKey($i)
			$snapshot = Get-SlideSnapshot $slide $wantModel $includeNotes
			if ($null -eq $snapshot.fingerprint) { $quality = 'unavailable' }
			else { $parts += $snapshot.fingerprint }
			$quality = Merge-FingerprintQuality $quality $snapshot.quality
			if ($wantModel) { $slides += $snapshot.model }
		}
		catch { $quality = 'unavailable' }
	}
	if ($countOk -and $quality -ne 'unavailable') { $fingerprint = Get-Sha256Hex ($parts -join '|') }
	return @{ fingerprint = $fingerprint; quality = $quality; slides = $slides }
}

function Get-PresentationModel($pres) {
	$path = ''
	try { $path = [string]$pres.FullName } catch { }
	$saved = $false
	try { $saved = ([int]$pres.Saved -ne 0) } catch { }
	return @{
		name     = [string]$pres.Name
		path     = $path
		slides   = [int]$pres.Slides.Count
		widthPx  = PtToPx ([double]$pres.PageSetup.SlideWidth)
		heightPx = PtToPx ([double]$pres.PageSetup.SlideHeight)
		saved    = $saved
	}
}

function Get-ColorHex($rgb) {
	$value = [int]$rgb
	$r = [int]($value % 256)
	$g = [int]([math]::Floor($value / 256) % 256)
	$b = [int]([math]::Floor($value / 65536) % 256)
	return ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
}

function Get-ShapeFingerprintLegacy($shape) {
	$parts = @(
		[string]$shape.Id,
		[string]$shape.Name,
		(Get-ShapeTypeName $shape.Type),
		(Format-Num (PtToPx ([double]$shape.Left))),
		(Format-Num (PtToPx ([double]$shape.Top))),
		(Format-Num (PtToPx ([double]$shape.Width))),
		(Format-Num (PtToPx ([double]$shape.Height))),
		('text|' + (Get-ShapeText $shape))
	)
	try {
		if ($shape.Fill.Visible -eq -1) { $parts += 'fill|' + (Get-ColorHex $shape.Fill.ForeColor.RGB) }
		else { $parts += 'fill|none' }
	}
	catch { $parts += 'fill|?' }
	try {
		if ($shape.Line.Visible -eq -1) {
			$parts += 'line|' + (Get-ColorHex $shape.Line.ForeColor.RGB) + '|' + (Format-Num ([double]$shape.Line.Weight))
		}
		else { $parts += 'line|none' }
	}
	catch { $parts += 'line|?' }
	try {
		if ($shape.HasTable -eq -1) {
			$table = $shape.Table
			for ($r = 1; $r -le $table.Rows.Count; $r++) {
				for ($c = 1; $c -le $table.Columns.Count; $c++) {
					$parts += 'cell|' + $r + '|' + $c + '|' + (Get-ShapeText $table.Cell($r, $c).Shape)
				}
			}
		}
	}
	catch { $parts += 'table|?' }
	try {
		if ($shape.HasChart -eq -1) {
			$chart = $shape.Chart
			$count = [int]$chart.SeriesCollection().Count
			for ($i = 1; $i -le $count; $i++) {
				$series = $chart.SeriesCollection($i)
				$parts += 'series|' + [string]$series.Name + '|' + (($series.Values) -join ',') + '|' + (($series.XValues) -join ',')
			}
		}
	}
	catch { $parts += 'chart|?' }
	try {
		if ([int]$shape.Type -eq 13 -or [int]$shape.Type -eq 11) {
			$parts += 'picture|' + [string]$shape.LinkFormat.SourceFullName
		}
	}
	catch { $parts += 'picture|?' }
	return ($parts -join '|')
}

function Get-SlideFingerprintLegacy($slide) {
	$parts = @('slide', [string]$slide.CustomLayout.Name)
	for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
		$parts += Get-ShapeFingerprintLegacy $slide.Shapes.Item($i)
	}
	$parts += 'notes|' + (Get-NotesText $slide)
	return Get-Sha256Hex ($parts -join '|')
}

function Get-DeckFingerprintLegacy($pres) {
	$parts = @('deck', [string]$pres.Slides.Count)
	for ($i = 1; $i -le $pres.Slides.Count; $i++) {
		$slide = $pres.Slides.Item($i)
		$parts += [string]$slide.SlideID
		$parts += Get-SlideFingerprintLegacy $slide
	}
	return Get-Sha256Hex ($parts -join '|')
}

function Get-ShapeFingerprint($shape) {
	$snapshot = Get-ShapeSnapshot $shape $false
	if ($null -eq $snapshot.fingerprintParts) { return $null }
	return ($snapshot.fingerprintParts -join '|')
}

function Get-SlideFingerprint($slide) {
	$snapshot = Get-SlideSnapshot $slide $false
	$script:LAST_FINGERPRINT_QUALITY = $snapshot.quality
	return $snapshot.fingerprint
}

function Get-DeckFingerprint($pres) {
	$snapshot = Get-DeckSnapshot $pres $false
	$script:LAST_FINGERPRINT_QUALITY = $snapshot.quality
	return $snapshot.fingerprint
}

function Get-FileFingerprint([string]$path) {
	if (Test-Path -LiteralPath $path) {
		$item = Get-Item -LiteralPath $path
		return Get-Sha256Hex ('file|' + $item.Length + '|' + $item.LastWriteTimeUtc.Ticks)
	}
	return Get-Sha256Hex 'absent'
}

function Get-PowerPointProcesses {
	try {
		$processes = @([System.Diagnostics.Process]::GetProcessesByName('POWERPNT'))
		return @{ ok = $true; processes = $processes }
	}
	catch {
		return @{ ok = $false; processes = @() }
	}
}

function Get-PptProcessSnapshot($processes) {
	$snapshot = @()
	foreach ($process in @($processes)) {
		if ($null -eq $process) { continue }
		$start = $null
		try { $start = $process.StartTime } catch { }
		$snapshot += @{ id = [int]$process.Id; startTime = $start }
	}
	return $snapshot
}

function Get-PptProcessById($processes, [int]$processId) {
	foreach ($process in @($processes)) {
		if ($null -ne $process -and [int]$process.Id -eq $processId) { return $process }
	}
	return $null
}

function Test-PptProcessAliveWithProbe($probe) {
	if ($null -eq $probe -or -not $probe.ok) { return 'unknown' }
	$processes = @($probe.processes)
	if ($processes.Count -eq 0) { return 'dead' }
	if ([int]$script:PPT_PID -le 0) { return 'unknown' }
	$process = Get-PptProcessById $processes ([int]$script:PPT_PID)
	if ($null -eq $process) { return 'dead' }
	try {
		if ([string]$process.ProcessName -ne 'POWERPNT') { return 'unknown' }
		if ($null -eq $script:PPT_STARTED_AT) { return 'unknown' }
		$start = $process.StartTime
		if ($null -eq $start) { return 'unknown' }
		if ($start.Ticks -ne ([datetime]$script:PPT_STARTED_AT).Ticks) { return 'unknown' }
		return 'alive'
	}
	catch { return 'unknown' }
}

function Test-PptProcessAlive {
	return Test-PptProcessAliveWithProbe (Get-PowerPointProcesses)
}

function Get-PptWindowPid {
	$windows = @()
	$enumOk = $true
	try { $windows = @([DeskhandPptNative]::FindPowerPointWindows()) }
	catch { $enumOk = $false }
	if ($enumOk -and $windows.Count -gt 0) {
		if ($windows.Count -ne 1) { return @{ ok = $true; pid = 0; count = $windows.Count } }
		[uint32]$windowPid = 0
		try { [void][DeskhandPptNative]::GetWindowThreadProcessId($windows[0], [ref]$windowPid) }
		catch { return @{ ok = $false; pid = 0; count = 1 } }
		if ($windowPid -eq 0) { return @{ ok = $false; pid = 0; count = 1 } }
		return @{ ok = $true; pid = [int]$windowPid; count = 1 }
	}

	$probe = Get-PowerPointProcesses
	if (-not $probe.ok) { return @{ ok = $false; pid = 0; count = 0 } }
	$handles = @()
	foreach ($process in @($probe.processes)) {
		if ($null -eq $process) { continue }
		try {
			$process.Refresh()
			if ($process.MainWindowHandle -ne [IntPtr]::Zero) { $handles += $process }
		}
		catch { }
	}
	if ($handles.Count -eq 1) { return @{ ok = $true; pid = [int]$handles[0].Id; count = 1 } }
	if ($handles.Count -eq 0) { return @{ ok = $enumOk; pid = 0; count = 0 } }
	return @{ ok = $true; pid = 0; count = $handles.Count }
}

function Set-PptMappingFromWindows($processes) {
	$window = Get-PptWindowPid
	if (-not $window.ok -or $window.count -ne 1 -or [int]$window.pid -le 0) { return $false }
	$process = Get-PptProcessById $processes ([int]$window.pid)
	if ($null -eq $process) { return $false }
	try {
		if ([string]$process.ProcessName -ne 'POWERPNT') { return $false }
		$start = $process.StartTime
		if ($null -eq $start) { return $false }
		$script:PPT_PID = [int]$process.Id
		$script:PPT_STARTED_AT = $start
		$script:PPT_MAPPING_FAILED = $false
		return $true
	}
	catch { return $false }
}

function Set-PptMappingFailure {
	$script:PPT_PID = 0
	$script:PPT_STARTED_AT = $null
	$script:PPT_MAPPING_FAILED = $true
	if ($script:STARTED_PPT) { Add-PptResetAttempt }
}

function Add-PptResetAttempt {
	$script:RESET_PENDING = $true
	$script:RESET_ATTEMPTS = [int]$script:RESET_ATTEMPTS + 1
	if ($script:RESET_ATTEMPTS -ge 2) {
		$script:RESET_TERMINAL = $true
		[Console]::Error.WriteLine('deskhand ppt: PowerPoint cleanup failed twice; restart the worker.')
	}
}

function Clear-PptResetState {
	$script:RESET_PENDING = $false
	$script:RESET_ATTEMPTS = 0
	$script:RESET_TERMINAL = $false
	$script:PPT_PID = 0
	$script:PPT_STARTED_AT = $null
	$script:PPT_MAPPING_FAILED = $false
}

function Clear-PptResetPending {
	$script:RESET_PENDING = $false
	$script:RESET_ATTEMPTS = 0
	$script:RESET_TERMINAL = $false
}

function Clear-PptReferences {
	$script:PptApp = $null
	$script:OURS = @()
	$script:OUR_WINDOWS = @()
	$script:STARTED_PPT = $false
}

function Clear-PptAll {
	Clear-PptReferences
	Clear-PptResetState
}

function Reset-PptReleaseRefs {
	$failed = $false
	$remaining = @()
	foreach ($pres in @($script:OURS)) {
		if ($null -eq $pres) { continue }
		try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
		catch { $remaining += $pres; $failed = $true }
	}
	$script:OURS = $remaining
	$remaining = @()
	foreach ($window in @($script:OUR_WINDOWS)) {
		if ($null -eq $window) { continue }
		try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($window) }
		catch { $remaining += $window; $failed = $true }
	}
	$script:OUR_WINDOWS = $remaining
	return (-not $failed)
}

function Set-PptNewMapping($beforeIds, [datetime]$launchAt) {
	$deadline = (Get-Date).AddSeconds(3)
	while ((Get-Date) -lt $deadline) {
		$probe = Get-PowerPointProcesses
		if (-not $probe.ok) { Start-Sleep -Milliseconds 100; continue }
		$newProcesses = @()
		foreach ($process in @($probe.processes)) {
			if ($null -eq $process) { continue }
			$processId = [int]$process.Id
			if ($beforeIds.ContainsKey($processId)) { continue }
			$start = $null
			try { $start = $process.StartTime } catch { continue }
			if ($null -eq $start -or $start -lt $launchAt) { continue }
			$newProcesses += $process
		}
		if ($newProcesses.Count -gt 1) { return $false }
		if ($newProcesses.Count -eq 1) {
			$window = Get-PptWindowPid
			if ($window.ok -and $window.count -eq 1 -and [int]$window.pid -eq [int]$newProcesses[0].Id) {
				$start = $null
				try { $start = $newProcesses[0].StartTime } catch { }
				if ($null -ne $start) {
					$script:PPT_PID = [int]$newProcesses[0].Id
					$script:PPT_STARTED_AT = $start
					$script:PPT_MAPPING_FAILED = $false
					return $true
				}
			}
		}
		Start-Sleep -Milliseconds 100
	}
	return $false
}

function New-PptInstance {
	$beforeProbe = Get-PowerPointProcesses
	if (-not $beforeProbe.ok) { Fail 'PowerPoint process state is unknown; restart the worker.' 'reset_pending' }
	$before = @(Get-PptProcessSnapshot $beforeProbe.processes)
	if ($before.Count -ne 0) { Fail 'PowerPoint is already running; refusing to start another instance.' 'reset_pending' }
	$beforeIds = @{}
	foreach ($item in $before) { $beforeIds[[int]$item.id] = $true }
	$launchAt = Get-Date
	try { $app = New-Object -ComObject PowerPoint.Application }
	catch {
		Start-Sleep -Seconds 1
		$retryProbe = Get-PowerPointProcesses
		if (-not $retryProbe.ok -or @($retryProbe.processes).Count -ne 0) {
			Fail ('PowerPoint COM could not start: ' + $_.Exception.Message) 'no_powerpoint'
		}
		try { $app = New-Object -ComObject PowerPoint.Application }
		catch { Fail ('PowerPoint COM could not start: ' + $_.Exception.Message) 'no_powerpoint' }
	}
	try { $app.DisplayAlerts = 1 } catch { }
	try { $app.AutomationSecurity = 3 } catch { }
	$script:PptApp = $app
	$script:STARTED_PPT = $true
	if (-not (Set-PptNewMapping $beforeIds $launchAt)) {
		Set-PptMappingFailure
		Fail 'PowerPoint started but its process could not be identified; restart the worker.' 'reset_pending'
	}
	return $app
}

function Attach-PptInstance($processes) {
	try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
	catch { Fail ('PowerPoint is running but could not be attached: ' + $_.Exception.Message) 'reset_pending' }
	try { $app.DisplayAlerts = 1 } catch { }
	try { $app.AutomationSecurity = 3 } catch { }
	$script:PptApp = $app
	$script:STARTED_PPT = $false
	if (-not (Set-PptMappingFromWindows $processes)) { Set-PptMappingFailure }
	return $app
}

function Resolve-PptNormal($probe, [string]$state) {
	if ($null -eq $probe -or -not $probe.ok) {
		Fail 'PowerPoint process state is unknown; restart the worker.' 'reset_pending'
	}
	$processes = @($probe.processes)
	if ($null -eq $script:PptApp) {
		if ((-not $script:STARTED_PPT) -and $processes.Count -eq 1) {
			return Attach-PptInstance $processes
		}
		if ($script:STARTED_PPT -and $processes.Count -gt 0) {
			Fail 'PowerPoint ownership is unknown; restart the worker.' 'reset_pending'
		}
		if ($processes.Count -eq 0) { return New-PptInstance }
		Fail 'PowerPoint is running in more than one instance; restart the worker.' 'reset_pending'
	}
	if ($state -eq 'unknown' -and [int]$script:PPT_PID -gt 0) {
		Fail 'PowerPoint process identity could not be verified; restart the worker.' 'reset_pending'
	}
	if ([int]$script:PPT_PID -eq 0) {
		if (-not (Set-PptMappingFromWindows $processes)) {
			Set-PptMappingFailure
			if ($script:STARTED_PPT) {
				Fail 'PowerPoint ownership could not be verified; restart the worker.' 'reset_pending'
			}
		}
	}
	return $script:PptApp
}

function Get-Ppt {
	for ($round = 0; $round -lt 2; $round++) {
		$probe = Get-PowerPointProcesses
		$state = Test-PptProcessAliveWithProbe $probe
		if ($state -eq 'dead') {
			Clear-PptAll
		}
		elseif ($script:RESET_TERMINAL) {
			Fail 'PowerPoint cleanup is pending; restart the worker.' 'reset_pending'
		}
		elseif ($script:RESET_PENDING) {
			Reset-Ppt -Recovery
			continue
		}
		$app = Resolve-PptNormal $probe $state
		try {
			$null = $app.Presentations.Count
			return $app
		}
		catch {
			Reset-Ppt -Recovery
			continue
		}
	}

	$probe = Get-PowerPointProcesses
	$state = Test-PptProcessAliveWithProbe $probe
	if ($state -eq 'dead') { Clear-PptAll }
	if ($script:RESET_TERMINAL -or $script:RESET_PENDING) {
		Fail 'PowerPoint cleanup is pending; restart the worker.' 'reset_pending'
	}
	$app = Resolve-PptNormal $probe $state
	try {
		$null = $app.Presentations.Count
		return $app
	}
	catch {
		Fail 'PowerPoint is not responding; restart the worker.' 'reset_pending'
	}
}

function Reset-Ppt([switch]$Recovery) {
	$state = Test-PptProcessAlive
	if ($state -eq 'dead') {
		Clear-PptAll
		return
	}

	if (-not (Reset-PptReleaseRefs)) {
		Add-PptResetAttempt
		if (-not $Recovery) { [Console]::Error.WriteLine('deskhand ppt: COM reference release failed during shutdown; PowerPoint may leak.') }
		return
	}

	if ($null -eq $script:PptApp) {
		if ($Recovery) { Clear-PptResetPending }
		else { Clear-PptAll }
		return
	}

	if (-not $script:STARTED_PPT) {
		if ($Recovery) { Clear-PptResetPending }
		else { Clear-PptAll }
		return
	}

	$count = $null
	try { $count = [int]$script:PptApp.Presentations.Count }
	catch {
		Add-PptResetAttempt
		if (-not $Recovery) { [Console]::Error.WriteLine('deskhand ppt: presentation count failed during shutdown; PowerPoint may leak.') }
		return
	}

	if ($count -gt 0) {
		if ($Recovery -and [int]$script:PPT_PID -eq 0) {
			Add-PptResetAttempt
		}
		elseif ($Recovery) {
			Clear-PptResetPending
		}
		else {
			Clear-PptAll
		}
		return
	}

	try { $script:PptApp.Quit() }
	catch {
		Add-PptResetAttempt
		if (-not $Recovery) { [Console]::Error.WriteLine('deskhand ppt: PowerPoint quit failed during shutdown; PowerPoint may leak.') }
		return
	}
	Clear-PptAll
}

function Get-Presentation($reference, [bool]$openIfPath, [bool]$editable = $true, [bool]$allowReadOnly = $false) {
	$app = Get-Ppt
	$ref = [string]$reference
	$match = $null
	if ([string]::IsNullOrWhiteSpace($ref)) {
		if ($app.Presentations.Count -eq 0) { Fail 'No presentation is open.' 'no_presentation' }
		if ($app.Presentations.Count -gt 1) { Fail 'Several presentations are open; pass presentation.' 'ambiguous' }
		$match = $app.Presentations.Item(1)
	}
	else {
		for ($i = 1; $i -le $app.Presentations.Count; $i++) {
			$pres = $app.Presentations.Item($i)
			if ([string]$pres.Name -eq $ref) { $match = $pres; break }
			$full = ''
			try { $full = [string]$pres.FullName } catch { }
			if (-not [string]::IsNullOrEmpty($full) -and $full -eq $ref) { $match = $pres; break }
		}
	}
	if ($null -ne $match) {
		if (-not $editable) { return $match }
		$readOnly = $false
		try { $readOnly = ([int]$match.ReadOnly -ne 0) } catch { }
		if (-not $readOnly) { return $match }
		$saved = $false
		try { $saved = ([int]$match.Saved -ne 0) } catch { }
		if (-not $saved) {
			Fail 'The presentation is open read-only with unsaved changes; close it in PowerPoint before editing.' 'read_only'
		}
		$name = [string]$match.Name
		try { $ref = [string]$match.FullName } catch { $ref = '' }
		$match.Close()
		# PowerPoint can return the old read-only object if Open races the close. Wait for it to leave
		# the collection before asking for an editable copy.
		for ($wait = 0; $wait -lt 50; $wait++) {
			$still = $false
			for ($j = 1; $j -le $app.Presentations.Count; $j++) {
				if ([string]$app.Presentations.Item($j).Name -eq $name) { $still = $true; break }
			}
			if (-not $still) { break }
			Start-Sleep -Milliseconds 100
		}
	}
	if ($openIfPath -and (Test-Path -LiteralPath $ref)) {
		$readOnlyFlag = -1
		if ($editable) { $readOnlyFlag = 0 }
		try { $pres = $app.Presentations.Open($ref, $readOnlyFlag, 0, 0) }
		catch { Fail ('Could not open ' + $ref + ': ' + $_.Exception.Message) 'com_error' }
		$script:OURS += $pres
		if ($editable) {
			$stillReadOnly = $false
			try { $stillReadOnly = ([int]$pres.ReadOnly -ne 0) } catch { }
			if ($stillReadOnly -and -not $allowReadOnly) {
				Fail ('PowerPoint reopened ' + $ref + ' read-only; close it in PowerPoint and retry.') 'read_only'
			}
		}
		return $pres
	}
	Fail ('No open presentation matches ' + $ref + '.') 'no_presentation'
}

function Get-Slide($pres, $index) {
	$n = [int]$index
	if ($n -lt 1 -or $n -gt $pres.Slides.Count) {
		Fail ('Slide ' + $n + ' does not exist; the deck has ' + $pres.Slides.Count + ' slide(s).') 'slide_not_found'
	}
	return $pres.Slides.Item($n)
}

function Get-ShapeById($slide, $shapeId) {
	$wanted = [int]$shapeId
	for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
		$shape = $slide.Shapes.Item($i)
		if ([int]$shape.Id -eq $wanted) { return $shape }
	}
	Fail ('Shape ' + $wanted + ' does not exist on slide ' + $slide.SlideIndex + '.') 'shape_not_found'
}

function Get-SlideIndexes($spec, [int]$count) {
	$text = [string]$spec
	if ([string]::IsNullOrWhiteSpace($text)) {
		$all = @()
		for ($i = 1; $i -le $count; $i++) { $all += $i }
		return $all
	}
	$out = @()
	foreach ($part in $text.Split(',')) {
		$p = $part.Trim()
		if ($p -eq '') { continue }
		if ($p.Contains('-')) {
			$bounds = $p.Split('-')
			if ($bounds.Count -ne 2) { Fail ('Bad slide selection: ' + $p) 'invalid_request' }
			$a = [int]$bounds[0]
			$b = [int]$bounds[1]
		}
		else {
			$a = [int]$p
			$b = $a
		}
		if ($a -lt 1 -or $b -gt $count -or $a -gt $b) { Fail ('Slide selection is outside 1-' + $count + ': ' + $p) 'slide_not_found' }
		for ($i = $a; $i -le $b; $i++) { if ($out -notcontains $i) { $out += $i } }
	}
	if ($out.Count -eq 0) { Fail 'No slides selected.' 'invalid_request' }
	return @($out | Sort-Object)
}

function Resolve-Layout([string]$name) {
	switch ($name) {
		'blank' { return 12 }
		'title' { return 1 }
		'title-content' { return 2 }
		'two-column' { return 3 }
		'section' { return 33 }
		default { Fail ('Unknown layout: ' + $name) 'invalid_request' }
	}
}

function Get-ThemeKey($theme) {
	if ($null -eq $theme) { return '' }
	return ([string](Get-Field $theme 'background')) + '|' + ([string](Get-Field $theme 'text')) + '|' + ([string](Get-Field $theme 'accent')) + '|' + ([string](Get-Field $theme 'font'))
}

function Set-Theme($pres, $theme) {
	if ($null -eq $theme) { return }
	$script:THEMES[[string]$pres.Name] = $theme
	$bg = [string](Get-Field $theme 'background')
	if ($bg -ne '') {
		$pres.SlideMaster.Background.Fill.Solid()
		$pres.SlideMaster.Background.Fill.ForeColor.RGB = ConvertTo-Color $bg
	}
}

function Get-Theme($pres) {
	$key = [string]$pres.Name
	if ($script:THEMES.ContainsKey($key)) { return $script:THEMES[$key] }
	return $null
}

function Set-SlideBackground($slide, [string]$hex) {
	$rgb = ConvertTo-Color $hex
	if ($null -eq $rgb) { return }
	$slide.Background.Fill.Solid()
	$slide.Background.Fill.ForeColor.RGB = $rgb
}

function Set-TextFrameOptions($shape, $padding, [string]$anchor) {
	if ($null -ne $padding) {
		$pt = PxToPt ([double]$padding)
		$shape.TextFrame.MarginLeft = $pt
		$shape.TextFrame.MarginRight = $pt
		$shape.TextFrame.MarginTop = $pt
		$shape.TextFrame.MarginBottom = $pt
	}
	if (-not [string]::IsNullOrWhiteSpace($anchor)) {
		switch ($anchor) {
			'top' { $shape.TextFrame.VerticalAnchor = 1 }
			'middle' { $shape.TextFrame.VerticalAnchor = 3 }
			'bottom' { $shape.TextFrame.VerticalAnchor = 4 }
			default { Fail-Shape ('Unknown anchor: ' + $anchor) 'invalid_request' }
		}
	}
}

function Set-Paragraphs($shape, $paragraphs, $theme) {
	$tf = $shape.TextFrame
	$tf.WordWrap = -1
	try { $tf.AutoSize = 0 } catch { }
	$paras = @($paragraphs)
	$texts = @()
	$spans = @()
	foreach ($para in $paras) {
		$runs = Get-Array $para 'runs'
		if ($runs.Count -eq 0) { $runs = @(@{ text = [string](Get-Field $para 'text') }) }
		$text = ''
		$runSpans = @()
		foreach ($run in $runs) {
			$t = [string](Get-Field $run 'text')
			$runSpans += @{ start = $text.Length; length = $t.Length; run = $run }
			$text += $t
		}
		$texts += $text
		$spans += , $runSpans
	}
	$tr = $tf.TextRange
	$tr.Text = ($texts -join "`r")
	$offset = 0
	for ($i = 0; $i -lt $texts.Count; $i++) {
		$fmt = $tr.Paragraphs($i + 1).ParagraphFormat
		$align = [string](Get-Field $paras[$i] 'align')
		if ($align -ne '') {
			switch ($align) {
				'left' { $fmt.Alignment = 1 }
				'center' { $fmt.Alignment = 2 }
				'right' { $fmt.Alignment = 3 }
				'justify' { $fmt.Alignment = 4 }
				default { Fail-Shape ('Unknown align: ' + $align) 'invalid_request' }
			}
		}
		$before = Get-Field $paras[$i] 'spaceBefore'
		if ($null -ne $before) { $fmt.SpaceBefore = PxToPt ([double]$before) }
		$after = Get-Field $paras[$i] 'spaceAfter'
		if ($null -ne $after) { $fmt.SpaceAfter = PxToPt ([double]$after) }
		$line = Get-Field $paras[$i] 'lineSpacing'
		if ($null -ne $line) {
			$fmt.LineRuleWithin = -1
			$fmt.SpaceWithin = PxToPt ([double]$line)
		}
		foreach ($span in $spans[$i]) {
			$start = $offset + [int]$span.start + 1
			$length = [int]$span.length
			if ($length -le 0) { continue }
			$range = $tr.Characters($start, $length)
			$run = $span.run
			$bold = Get-Field $run 'bold'
			if ($null -ne $bold) { if ([bool]$bold) { $range.Font.Bold = -1 } else { $range.Font.Bold = 0 } }
			$italic = Get-Field $run 'italic'
			if ($null -ne $italic) { if ([bool]$italic) { $range.Font.Italic = -1 } else { $range.Font.Italic = 0 } }
			$color = [string](Get-Field $run 'color')
			if ($color -eq '' -and $null -ne $theme) { $color = [string](Get-Field $theme 'text') }
			if ($color -ne '') { $range.Font.Color.RGB = ConvertTo-Color $color }
			$size = Get-Field $run 'size'
			if ($null -ne $size) { $range.Font.Size = PxToPt ([double]$size) }
			$font = [string](Get-Field $run 'font')
			if ($font -eq '' -and $null -ne $theme) { $font = [string](Get-Field $theme 'font') }
			if ($font -ne '') { $range.Font.Name = $font }
		}
		$offset += $texts[$i].Length + 1
	}
}

function Set-SlideTitle($slide, [string]$title, [string]$subtitle, $theme) {
	$textColor = ''
	$font = ''
	if ($null -ne $theme) {
		$textColor = [string](Get-Field $theme 'text')
		$font = [string](Get-Field $theme 'font')
	}
	if ($title -ne '') {
		$found = $false
		for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
			$shape = $slide.Shapes.Item($i)
			if ([int]$shape.Type -eq 14) {
				$t = 0
				try { $t = [int]$shape.PlaceholderFormat.Type } catch { }
				if ($t -eq 1 -or $t -eq 3) {
					$shape.TextFrame.TextRange.Text = $title
					if ($textColor -ne '') { $shape.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Color $textColor }
					if ($font -ne '') { $shape.TextFrame.TextRange.Font.Name = $font }
					$found = $true
					break
				}
			}
		}
		if (-not $found) {
			$box = $slide.Shapes.AddTextbox(1, (PxToPt 96), (PxToPt 64), (PxToPt 1088), (PxToPt 96))
			$box.TextFrame.TextRange.Text = $title
			$box.TextFrame.TextRange.Font.Size = PxToPt 40
			$box.TextFrame.TextRange.Font.Bold = -1
			if ($textColor -ne '') { $box.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Color $textColor }
			if ($font -ne '') { $box.TextFrame.TextRange.Font.Name = $font }
		}
	}
	if ($subtitle -ne '') {
		$found = $false
		for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
			$shape = $slide.Shapes.Item($i)
			if ([int]$shape.Type -eq 14) {
				$t = 0
				try { $t = [int]$shape.PlaceholderFormat.Type } catch { }
				if ($t -eq 4 -or $t -eq 2) {
					$shape.TextFrame.TextRange.Text = $subtitle
					if ($textColor -ne '') { $shape.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Color $textColor }
					if ($font -ne '') { $shape.TextFrame.TextRange.Font.Name = $font }
					$found = $true
					break
				}
			}
		}
		if (-not $found) {
			$box = $slide.Shapes.AddTextbox(1, (PxToPt 96), (PxToPt 180), (PxToPt 1088), (PxToPt 72))
			$box.TextFrame.TextRange.Text = $subtitle
			$box.TextFrame.TextRange.Font.Size = PxToPt 24
			if ($textColor -ne '') { $box.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Color $textColor }
			if ($font -ne '') { $box.TextFrame.TextRange.Font.Name = $font }
		}
	}
}

function Ensure-ChartWindow($pres) {
	if ($pres.Windows.Count -gt 0) { return }
	$window = $pres.NewWindow()
	try { $window.WindowState = 2 } catch { }
	$script:OUR_WINDOWS += $window
}

function Set-ChartSeries($chart, $categories, $series) {
	$wanted = $series.Count
	$existing = -1
	try {
		$existing = [int]$chart.SeriesCollection().Count
		for ($i = $existing + 1; $i -le $wanted; $i++) { $null = $chart.SeriesCollection().NewSeries() }
		for ($i = $existing; $i -gt $wanted; $i--) { $chart.SeriesCollection($i).Delete() }
		for ($i = 1; $i -le $wanted; $i++) {
			$ser = $chart.SeriesCollection($i)
			$ser.Name = [string](Get-Field $series[$i - 1] 'name')
			$ser.Values = Get-Array $series[$i - 1] 'values'
			$ser.XValues = $categories
		}
	}
	catch {
		Fail ('Could not set chart series (existing=' + $existing + ', wanted=' + $wanted + '): ' + $_.Exception.Message) 'com_error'
	}
}

function Set-ChartOptions($chart, $req) {
	$title = [string](Get-Field $req 'title')
	if ($title -ne '') {
		$chart.HasTitle = -1
		$chart.ChartTitle.Text = $title
	}
	$legend = Get-Field $req 'legend'
	if ($null -ne $legend) { if ([bool]$legend) { $chart.HasLegend = -1 } else { $chart.HasLegend = 0 } }
	$axisTitle = [string](Get-Field $req 'valueAxisTitle')
	if ($axisTitle -ne '') {
		$chart.Axes(2).HasTitle = -1
		$chart.Axes(2).AxisTitle.Text = $axisTitle
	}
	$dataLabels = Get-Field $req 'dataLabels'
	$numberFormat = [string](Get-Field $req 'numberFormat')
	if ($numberFormat -ne '') {
		$chart.Axes(2).TickLabels.NumberFormat = $numberFormat
	}
	$count = 0
	try { $count = [int]$chart.SeriesCollection().Count } catch { }
	for ($i = 1; $i -le $count; $i++) {
		$ser = $chart.SeriesCollection($i)
		if ($null -ne $dataLabels) {
			if ([bool]$dataLabels) {
				$ser.HasDataLabels = -1
				$points = 0
				try { $points = [int]$ser.Points().Count } catch { }
				for ($j = 1; $j -le $points; $j++) {
					$point = $ser.Points($j)
					$point.HasDataLabel = -1
					$point.DataLabel.ShowValue = -1
					if ($numberFormat -ne '') { $point.DataLabel.NumberFormat = $numberFormat }
				}
			}
			else {
				$ser.HasDataLabels = 0
			}
		}
	}
}

function Set-ChartColors($chart, $series, [string]$chartType, $theme) {
	$accent = ''
	if ($null -ne $theme) { $accent = [string](Get-Field $theme 'accent') }
	for ($i = 1; $i -le $series.Count; $i++) {
		$color = [string](Get-Field $series[$i - 1] 'color')
		if ($color -eq '') { $color = $accent }
		if ($color -eq '') { continue }
		$rgb = ConvertTo-Color $color
		if ($chartType -eq 'line') {
			$chart.SeriesCollection($i).Format.Line.ForeColor.RGB = $rgb
		}
		else {
			$chart.SeriesCollection($i).Format.Fill.ForeColor.RGB = $rgb
		}
	}
}

function Add-TextShape($pres, $slide, $req) {
	$x = PxToPt ([double](Get-Field $req 'x'))
	$y = PxToPt ([double](Get-Field $req 'y'))
	$w = PxToPt ([double](Get-Field $req 'w'))
	$h = PxToPt ([double](Get-Field $req 'h'))
	$script:MUTATED = $true
	$script:CURRENT_SHAPE_MUTATED = $true
	try { $shape = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h) }
	catch { Fail-Shape ('Could not add a text box: ' + $_.Exception.Message) 'com_error' }
	Set-Paragraphs $shape (Get-Field $req 'paragraphs') (Get-Theme $pres)
	Set-TextFrameOptions $shape (Get-Field $req 'padding') ([string](Get-Field $req 'anchor'))
	Set-Fill $shape ([string](Get-Field $req 'fill'))
	Set-Line $shape ([string](Get-Field $req 'line'))
	return $shape
}

function Add-TableShape($pres, $slide, $req) {
	$header = Get-Array $req 'header'
	$rows = Get-Array $req 'rows'
	$cols = $header.Count
	foreach ($row in $rows) { if (@($row).Count -gt $cols) { $cols = @($row).Count } }
	if ($cols -lt 1) { Fail-Shape 'A table needs at least one column.' 'invalid_request' }
	$rowCount = $rows.Count + $(if ($header.Count -gt 0) { 1 } else { 0 })
	if ($rowCount -lt 1) { Fail-Shape 'A table needs at least one row.' 'invalid_request' }
	$columnWidths = Get-Array $req 'columnWidths'
	if ($columnWidths.Count -gt 0 -and $columnWidths.Count -ne $cols) {
		Fail-Shape 'columnWidths must have one entry per column.' 'invalid_request'
	}
	$x = PxToPt ([double](Get-Field $req 'x'))
	$y = PxToPt ([double](Get-Field $req 'y'))
	$w = PxToPt ([double](Get-Field $req 'w'))
	$h = PxToPt ([double](Get-Field $req 'h'))
	$script:MUTATED = $true
	$script:CURRENT_SHAPE_MUTATED = $true
	try { $shape = $slide.Shapes.AddTable($rowCount, $cols, $x, $y, $w, $h) }
	catch { Fail-Shape ('Could not add a table: ' + $_.Exception.Message) 'com_error' }
	$table = $shape.Table
	$r = 1
	if ($header.Count -gt 0) {
		for ($c = 1; $c -le $cols; $c++) {
			$text = ''
			if ($c -le $header.Count) { $text = [string]$header[$c - 1] }
			$table.Cell($r, $c).Shape.TextFrame.TextRange.Text = $text
		}
		$r++
	}
	foreach ($row in $rows) {
		$values = @($row)
		for ($c = 1; $c -le $cols; $c++) {
			$text = ''
			if ($c -le $values.Count) { $text = [string]$values[$c - 1] }
			$table.Cell($r, $c).Shape.TextFrame.TextRange.Text = $text
		}
		$r++
	}
	if ($columnWidths.Count -gt 0) {
		for ($c = 1; $c -le $cols; $c++) { $table.Columns.Item($c).Width = PxToPt ([double]$columnWidths[$c - 1]) }
	}
	$headerFill = [string](Get-Field $req 'headerFill')
	$theme = Get-Theme $pres
	$themeAccent = ''
	$themeText = ''
	$themeFont = ''
	if ($null -ne $theme) {
		$themeAccent = [string](Get-Field $theme 'accent')
		$themeText = [string](Get-Field $theme 'text')
		$themeFont = [string](Get-Field $theme 'font')
	}
	if ($headerFill -eq '') { $headerFill = $themeAccent }
	if ($headerFill -ne '') {
		for ($c = 1; $c -le $cols; $c++) { Set-Fill $table.Cell(1, $c).Shape $headerFill }
	}
	$fontSize = Get-Field $req 'fontSize'
	$headerFontSize = Get-Field $req 'headerFontSize'
	for ($ri = 1; $ri -le $rowCount; $ri++) {
		for ($c = 1; $c -le $cols; $c++) {
			$cellShape = $table.Cell($ri, $c).Shape
			if ($null -ne $fontSize) { $cellShape.TextFrame.TextRange.Font.Size = PxToPt ([double]$fontSize) }
			if ($themeText -ne '') { $cellShape.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Color $themeText }
			if ($themeFont -ne '') { $cellShape.TextFrame.TextRange.Font.Name = $themeFont }
			if ($ri -eq 1 -and $header.Count -gt 0) {
				$cellShape.TextFrame.TextRange.Font.Bold = -1
				if ($null -ne $headerFontSize) { $cellShape.TextFrame.TextRange.Font.Size = PxToPt ([double]$headerFontSize) }
			}
		}
	}
	return $shape
}

function Test-ShapeSpecs($specs) {
	$index = 0
	foreach ($spec in @($specs)) {
		$index++
		$script:CURRENT_SHAPE_INDEX = $index
		$kind = [string](Get-Field $spec 'kind')
		if ($kind -ne 'text' -and $kind -ne 'table' -and $kind -ne 'picture') {
			Fail-Shape ('shapes[' + $index + '] has unknown kind: ' + $kind) 'invalid_request'
		}
		foreach ($key in @('x', 'y', 'w', 'h')) {
			if ($null -eq (Get-Field $spec $key)) { Fail-Shape ('shapes[' + $index + '] needs ' + $key) 'invalid_request' }
		}
		if ([double](Get-Field $spec 'w') -le 0 -or [double](Get-Field $spec 'h') -le 0) {
			Fail-Shape ('shapes[' + $index + '] needs positive w and h') 'invalid_request'
		}
		if ($kind -eq 'text') {
			if ((Get-Array $spec 'paragraphs').Count -eq 0) { Fail-Shape ('shapes[' + $index + '] needs paragraphs') 'invalid_request' }
		}
		elseif ($kind -eq 'table') {
			$header = Get-Array $spec 'header'
			$rows = Get-Array $spec 'rows'
			if ($header.Count -eq 0 -and $rows.Count -eq 0) { Fail-Shape ('shapes[' + $index + '] needs a header or rows') 'invalid_request' }
			$cols = $header.Count
			foreach ($row in $rows) { if (@($row).Count -gt $cols) { $cols = @($row).Count } }
			$widths = Get-Array $spec 'columnWidths'
			if ($widths.Count -gt 0 -and $widths.Count -ne $cols) { Fail-Shape ('shapes[' + $index + '].columnWidths must match the columns') 'invalid_request' }
		}
		elseif ($kind -eq 'picture') {
			$path = [string](Get-Field $spec 'path')
			$data = [string](Get-Field $spec 'dataBase64')
			if (($path -eq '') -eq ($data -eq '')) { Fail-Shape ('shapes[' + $index + '] needs exactly one of path or dataBase64') 'invalid_request' }
			if ($path -ne '' -and -not (Test-Path -LiteralPath $path)) { Fail-Shape ('Image not found: ' + $path) 'invalid_request' }
		}
	}
	$script:CURRENT_SHAPE_INDEX = $null
}

function Add-Shape($pres, $slide, $req, [string]$suffix) {
	$kind = [string](Get-Field $req 'kind')
	switch ($kind) {
		'text' { return Add-TextShape $pres $slide $req }
		'table' { return Add-TableShape $pres $slide $req }
		'picture' { return Add-PictureShape $slide $req $suffix }
		default { Fail-Shape ('Unknown shape kind: ' + $kind) 'invalid_request' }
	}
}

function Add-PictureShape($slide, $req, [string]$suffix = 'single') {
	$path = [string](Get-Field $req 'path')
	$data = [string](Get-Field $req 'dataBase64')
	$temp = ''
	if ($data -ne '') {
		try { $bytes = [Convert]::FromBase64String($data) }
		catch { Fail-Shape 'dataBase64 is not valid base64' 'invalid_request' }
		$ext = '.png'
		if ($bytes.Length -gt 2 -and $bytes[0] -eq 255 -and $bytes[1] -eq 216) { $ext = '.jpg' }
		elseif ($bytes.Length -gt 3 -and $bytes[0] -eq 71 -and $bytes[1] -eq 73 -and $bytes[2] -eq 70) { $ext = '.gif' }
		$temp = Join-Path $script:WORK ('img-' + $script:STEM + '-' + $suffix + $ext)
		[IO.File]::WriteAllBytes($temp, $bytes)
		$path = $temp
	}
	if (-not (Test-Path -LiteralPath $path)) { Fail-Shape ('Image not found: ' + $path) 'com_error' }
	$x = PxToPt ([double](Get-Field $req 'x'))
	$y = PxToPt ([double](Get-Field $req 'y'))
	$w = PxToPt ([double](Get-Field $req 'w'))
	$h = PxToPt ([double](Get-Field $req 'h'))
	$fit = [string](Get-Field $req 'fit')
	if ($fit -eq '') { $fit = 'stretch' }
	if ($fit -ne 'stretch' -and $fit -ne 'contain' -and $fit -ne 'cover') { Fail-Shape ('Unknown fit: ' + $fit) 'invalid_request' }
	try {
		$script:MUTATED = $true
		$script:CURRENT_SHAPE_MUTATED = $true
		if ($fit -eq 'stretch') {
			return $slide.Shapes.AddPicture($path, 0, -1, $x, $y, $w, $h)
		}
		$shape = $slide.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1)
		$natW = [double]$shape.Width
		$natH = [double]$shape.Height
		if ($natW -le 0 -or $natH -le 0) { return $shape }
		if ($fit -eq 'contain') {
			$scale = [math]::Min($w / $natW, $h / $natH)
			$shape.Width = $natW * $scale
			$shape.Height = $natH * $scale
			$shape.Left = $x + (($w - $shape.Width) / 2)
			$shape.Top = $y + (($h - $shape.Height) / 2)
		}
		elseif ($fit -eq 'cover') {
			$scale = [math]::Max($w / $natW, $h / $natH)
			$newW = $natW * $scale
			$newH = $natH * $scale
			$shape.PictureFormat.CropLeft = (($newW - $w) / 2) / $scale
			$shape.PictureFormat.CropRight = (($newW - $w) / 2) / $scale
			$shape.PictureFormat.CropTop = (($newH - $h) / 2) / $scale
			$shape.PictureFormat.CropBottom = (($newH - $h) / 2) / $scale
			$shape.Left = $x
			$shape.Top = $y
			$shape.Width = $w
			$shape.Height = $h
		}
		return $shape
	}
	finally {
		if ($temp -ne '') { try { Remove-Item -LiteralPath $temp -Force } catch { } }
	}
}

# --- main -------------------------------------------------------------------

function Invoke-Request([string]$stem) {
	$script:RESP = Join-Path $script:WORK ('resp-' + $stem + '.json')
	$script:RESPONDED = $false
	$script:MUTATED = $false
	$script:CURRENT_SHAPE_MUTATED = $false
	$script:SHAPE_ERROR_MODE = $false
	$script:CURRENT_SHAPE_INDEX = $null
	$script:LAST_FINGERPRINT_QUALITY = $null
	$script:STEM = $stem
	$pres = $null
	try {
		$req = [IO.File]::ReadAllText((Join-Path $script:WORK ('req-' + $stem + '.json')), [Text.Encoding]::UTF8) | ConvertFrom-Json
		$op = [string]$req.op
		$sw = [Diagnostics.Stopwatch]::StartNew()

		if ($op -eq 'ping') {
			$null = Get-Ppt
			Write-Response @{ ok = $true; data = @{ ready = $true; elapsedMs = [int]$sw.ElapsedMilliseconds } }
			return
		}

		switch ($op) {
			'presentations' {
				if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -eq 0) {
					Write-Response @{ ok = $true; data = @{ running = $false; presentations = @() } }
				}
				else {
					$app = Get-Ppt
					$out = @()
					for ($i = 1; $i -le $app.Presentations.Count; $i++) { $out += Get-PresentationModel $app.Presentations.Item($i) }
					Write-Response @{ ok = $true; data = @{ running = $true; presentations = $out } }
				}
			}

			'read_deck' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true $false
				$includeNotes = [bool](Get-Field $req 'includeNotes')
				$indexes = Get-SlideIndexes (Get-Field $req 'slides') $pres.Slides.Count
				$snapshot = Get-DeckSnapshot $pres $true $indexes $includeNotes
				$data = @{ presentation = (Get-PresentationModel $pres); slides = $snapshot.slides; fingerprintQuality = $snapshot.quality }
				if ($null -ne $snapshot.fingerprint) { $data.fingerprint = $snapshot.fingerprint }
				Write-Response @{ ok = $true; data = $data }
			}

			'read_slide' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true $false
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$includeNotes = [bool](Get-Field $req 'includeNotes')
				$snapshot = Get-SlideSnapshot $slide $true $includeNotes
				$data = @{ slide = $snapshot.model; fingerprintQuality = $snapshot.quality }
				if ($null -ne $snapshot.fingerprint) { $data.fingerprint = $snapshot.fingerprint }
				Write-Response @{ ok = $true; data = $data }
			}

			'render' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true $false
				$outDir = [string](Get-Field $req 'outDir')
				if ([string]::IsNullOrWhiteSpace($outDir)) { Fail 'outDir is required' 'invalid_request' }
				if (-not (Test-Path -LiteralPath $outDir)) { $null = New-Item -ItemType Directory -Force -Path $outDir }
				$format = [string](Get-Field $req 'format')
				if ($format -eq '') { $format = 'png' }
				$format = $format.ToLowerInvariant()
				$width = [int](Get-Num (Get-Field $req 'width') 1600)
				$height = [int](Get-Num (Get-Field $req 'height') 900)
				$overwrite = [bool](Get-Field $req 'overwrite')
				$indexes = Get-SlideIndexes (Get-Field $req 'slides') $pres.Slides.Count
				$files = @()
				if ($format -eq 'pdf') {
					if ($indexes.Count -ne $pres.Slides.Count) { Fail 'A PDF contains the whole deck; omit slides.' 'unsupported' }
					$path = Join-Path $outDir 'deck.pdf'
					if ((Test-Path -LiteralPath $path) -and -not $overwrite) { Fail ('File exists: ' + $path) 'file_exists' }
					$pres.SaveCopyAs($path, 32)
					$files += @{ path = $path; widthPx = (PtToPx ([double]$pres.PageSetup.SlideWidth)); heightPx = (PtToPx ([double]$pres.PageSetup.SlideHeight)) }
				}
				elseif ($format -eq 'png') {
					foreach ($i in $indexes) {
						$path = Join-Path $outDir ('slide-' + $i + '.png')
						if ((Test-Path -LiteralPath $path) -and -not $overwrite) { Fail ('File exists: ' + $path) 'file_exists' }
						$pres.Slides.Item($i).Export($path, 'PNG', $width, $height)
						$files += @{ path = $path; widthPx = $width; heightPx = $height }
					}
				}
				else { Fail ('Unknown format: ' + $format) 'invalid_request' }
				Write-Response @{ ok = $true; data = @{ files = $files; elapsedMs = [int]$sw.ElapsedMilliseconds; outDir = $outDir } }
			}

			'create_deck' {
				$width = [int](Get-Num (Get-Field $req 'width') 1280)
				$height = [int](Get-Num (Get-Field $req 'height') 720)
				$path = [string](Get-Field $req 'path')
				$theme = Get-Field $req 'theme'
				$fingerprint = Get-Sha256Hex ('new:' + $width + 'x' + $height + ':' + (Get-ThemeKey $theme))
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = @{ name = ''; path = $path; slides = 0; widthPx = $width; heightPx = $height; saved = $false } } }
				}
				else {
					if ($path -ne '' -and (Test-Path -LiteralPath $path)) { Fail ('File exists: ' + $path) 'file_exists' }
					$app = Get-Ppt
					try { $pres = $app.Presentations.Add(0) }
					catch { Fail ('Could not create a presentation: ' + $_.Exception.Message) 'com_error' }
					$script:OURS += $pres
					$script:MUTATED = $true
					$pres.PageSetup.SlideWidth = PxToPt $width
					$pres.PageSetup.SlideHeight = PxToPt $height
					Set-Theme $pres $theme
					if ($path -ne '') { $pres.SaveAs($path, 24) }
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-DeckFingerprint $pres); presentation = (Get-PresentationModel $pres) } }
				}
			}

			'add_slide' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$fingerprint = Get-DeckFingerprint $pres
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = (Get-PresentationModel $pres) } }
				}
				else {
					$layout = Resolve-Layout ([string](Get-Field $req 'layout'))
					$position = $pres.Slides.Count + 1
					if ($null -ne (Get-Field $req 'position')) { $position = [int](Get-Field $req 'position') }
					if ($position -lt 1 -or $position -gt ($pres.Slides.Count + 1)) { Fail ('Position is outside 1-' + ($pres.Slides.Count + 1) + '.') 'invalid_request' }
					try { $slide = $pres.Slides.Add($position, $layout) }
					catch { Fail ('Could not add a slide: ' + $_.Exception.Message) 'com_error' }
					$script:MUTATED = $true
					Set-SlideTitle $slide ([string](Get-Field $req 'title')) ([string](Get-Field $req 'subtitle')) (Get-Theme $pres)
					$bg = [string](Get-Field $req 'background')
					if ($bg -ne '') { Set-SlideBackground $slide $bg }
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-DeckFingerprint $pres); presentation = (Get-PresentationModel $pres); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'add_text' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$null = Add-TextShape $pres $slide $req
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'add_table' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$header = Get-Array $req 'header'
				$rows = Get-Array $req 'rows'
				$cols = $header.Count
				foreach ($row in $rows) { if (@($row).Count -gt $cols) { $cols = @($row).Count } }
				if ($cols -lt 1) { Fail 'A table needs at least one column.' 'invalid_request' }
				$rowCount = $rows.Count + $(if ($header.Count -gt 0) { 1 } else { 0 })
				if ($rowCount -lt 1) { Fail 'A table needs at least one row.' 'invalid_request' }
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$null = Add-TableShape $pres $slide $req
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'add_chart' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$type = [string](Get-Field $req 'type')
				$categories = Get-Array $req 'categories'
				$series = Get-Array $req 'series'
				if ($categories.Count -eq 0 -or $series.Count -eq 0) { Fail 'A chart needs categories and at least one series.' 'invalid_request' }
				foreach ($s in $series) {
					if ((Get-Array $s 'values').Count -ne $categories.Count) { Fail 'Every series needs one value per category.' 'invalid_request' }
				}
				$xlType = 51
				if ($type -eq 'bar') { $xlType = 57 }
				elseif ($type -eq 'line') { $xlType = 4 }
				elseif ($type -eq 'pie') { $xlType = 5 }
				elseif ($type -ne 'column') { Fail ('Unknown chart type: ' + $type) 'invalid_request' }
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$x = PxToPt ([double](Get-Field $req 'x'))
					$y = PxToPt ([double](Get-Field $req 'y'))
					$w = PxToPt ([double](Get-Field $req 'w'))
					$h = PxToPt ([double](Get-Field $req 'h'))
					Ensure-ChartWindow $pres
					try { $shape = $slide.Shapes.AddChart2(-1, $xlType, $x, $y, $w, $h, $false) }
					catch { Fail ('Could not add a chart: ' + $_.Exception.Message) 'com_error' }
					$script:MUTATED = $true
					$chart = $shape.Chart
					Set-ChartSeries $chart $categories $series
					Set-ChartOptions $chart $req
					Set-ChartColors $chart $series $type (Get-Theme $pres)
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'add_picture' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$null = Add-PictureShape $slide $req 'single'
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'add_shapes' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$shapes = @(Get-Array $req 'shapes')
				$applying = Test-Apply $req
				$script:SHAPE_ERROR_MODE = $true
				try {
					Test-ShapeSpecs $shapes
					$fingerprint = Get-SlideFingerprint $slide
					Assert-Fingerprint $req $fingerprint
					if (-not $applying) {
						Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; shapeIds = @(); appliedCount = 0 } }
					}
					else {
						$script:MUTATED = $false
						$script:CURRENT_SHAPE_MUTATED = $false
						$ids = @()
						$index = 0
						foreach ($spec in $shapes) {
							$index++
							$script:CURRENT_SHAPE_INDEX = $index
							$script:CURRENT_SHAPE_MUTATED = $false
							try {
								$shape = Add-Shape $pres $slide $spec ([string]$index)
								$ids += $shape.Id
							}
							catch {
								$err = $_.TargetObject
								$code = 'com_error'
								$message = $_.Exception.Message
								if ($null -ne $err -and $err.deskhand_shape_error) { $code = [string]$err.code; $message = [string]$err.message }
								$details = @{ status = 'partial'; tokenSpent = $true; appliedCount = $ids.Count; failedIndex = $index; shapeIds = $ids; failedShapeMayExist = $script:CURRENT_SHAPE_MUTATED }
								if ($ids.Count -eq 0 -and -not $script:CURRENT_SHAPE_MUTATED) { $details.status = 'invalid' }
								Fail $message $code $details
							}
						}
						if ($ids.Count -ne $shapes.Count) {
							Fail 'The bridge added fewer shapes than requested; read the slide before retrying.' 'applied_unverified' @{ status = 'unknown'; tokenSpent = $true; appliedCount = $null; failedIndex = $null; shapeIds = $null; failedShapeMayExist = $null }
						}
						Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); shapeIds = $ids; appliedCount = $ids.Count } }
					}
				}
				catch {
					if ($script:RESPONDED) { return }
					$err = $_.TargetObject
					$code = 'com_error'
					$message = $_.Exception.Message
					if ($null -ne $err -and $err.deskhand_shape_error) { $code = [string]$err.code; $message = [string]$err.message }
					$details = @{ status = 'invalid'; tokenSpent = $applying; appliedCount = 0; failedIndex = $script:CURRENT_SHAPE_INDEX; shapeIds = @(); failedShapeMayExist = $false }
					Fail $message $code $details
				}
				finally {
					$script:SHAPE_ERROR_MODE = $false
					$script:CURRENT_SHAPE_INDEX = $null
				}
			}

			'set_notes' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $true) } }
				}
				else {
					try { $ph = $slide.NotesPage.Shapes.Placeholders.Item(2) }
					catch { Fail 'This slide has no notes placeholder.' 'com_error' }
					$ph.TextFrame.TextRange.Text = [string](Get-Field $req 'text')
					$script:MUTATED = $true
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $true) } }
				}
			}

			'update_shape' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$shape = Get-ShapeById $slide (Get-Field $req 'shapeId')
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$text = Get-Field $req 'text'
					$paragraphs = Get-Field $req 'paragraphs'
					if ($null -ne $text -and $null -ne $paragraphs) { Fail 'Pass text or paragraphs, not both.' 'invalid_request' }
					if (($null -ne $text -or $null -ne $paragraphs) -and ($shape.HasTable -eq -1 -or $shape.HasChart -eq -1)) {
						Fail 'update_shape cannot change table or chart content; delete the shape and add it again.' 'invalid_request'
					}
					$script:MUTATED = $true
					if ($null -ne (Get-Field $req 'x')) { $shape.Left = PxToPt ([double](Get-Field $req 'x')) }
					if ($null -ne (Get-Field $req 'y')) { $shape.Top = PxToPt ([double](Get-Field $req 'y')) }
					if ($null -ne (Get-Field $req 'w')) { $shape.Width = PxToPt ([double](Get-Field $req 'w')) }
					if ($null -ne (Get-Field $req 'h')) { $shape.Height = PxToPt ([double](Get-Field $req 'h')) }
					if ($null -ne $paragraphs) { Set-Paragraphs $shape $paragraphs (Get-Theme $pres) }
					elseif ($null -ne $text) { $shape.TextFrame.TextRange.Text = [string]$text }
					Set-Fill $shape ([string](Get-Field $req 'fill'))
					Set-Line $shape ([string](Get-Field $req 'line'))
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'delete_shape' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$shape = Get-ShapeById $slide (Get-Field $req 'shapeId')
				$fingerprint = Get-SlideFingerprint $slide
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$script:MUTATED = $true
					$shape.Delete()
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-SlideFingerprint $slide); slide = (Get-SlideModel $slide $false) } }
				}
			}

			'move_slide' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$from = [int](Get-Field $req 'from')
				$to = [int](Get-Field $req 'to')
				$null = Get-Slide $pres $from
				if ($to -lt 1 -or $to -gt $pres.Slides.Count) { Fail ('Position is outside 1-' + $pres.Slides.Count + '.') 'invalid_request' }
				$fingerprint = Get-DeckFingerprint $pres
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = (Get-PresentationModel $pres) } }
				}
				else {
					$script:MUTATED = $true
					$pres.Slides.Item($from).MoveTo($to)
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-DeckFingerprint $pres); presentation = (Get-PresentationModel $pres) } }
				}
			}

			'delete_slide' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true
				$slide = Get-Slide $pres (Get-Field $req 'slide')
				$fingerprint = Get-DeckFingerprint $pres
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = (Get-PresentationModel $pres); slide = (Get-SlideModel $slide $false) } }
				}
				else {
					$script:MUTATED = $true
					$slide.Delete()
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-DeckFingerprint $pres); presentation = (Get-PresentationModel $pres) } }
				}
			}

			'save_deck' {
				$pres = Get-Presentation (Get-Field $req 'presentation') $true $true $true
				$path = [string](Get-Field $req 'path')
				if ([string]::IsNullOrWhiteSpace($path)) { Fail 'path is required' 'invalid_request' }
				$overwrite = [bool](Get-Field $req 'overwrite')
				$current = ''
				try { $current = [string]$pres.FullName } catch { }
				$readOnly = $false
				try { $readOnly = ([int]$pres.ReadOnly -ne 0) } catch { }
				if ($readOnly -and $current -eq $path) {
					Fail 'The presentation is open read-only; save it to a different path or close it in PowerPoint first.' 'read_only'
				}
				if ((Test-Path -LiteralPath $path) -and -not $overwrite) { Fail ('File exists: ' + $path) 'file_exists' }
				if ($current -eq $path -and -not $overwrite) { Fail ('File exists: ' + $path) 'file_exists' }
				$fingerprint = Get-Sha256Hex ((Get-FileFingerprint $path) + '|' + (Get-DeckFingerprint $pres))
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = (Get-PresentationModel $pres); path = $path } }
				}
				else {
					$script:MUTATED = $true
					if ($current -eq $path) { $pres.Save() }
					else { $pres.SaveAs($path, 24) }
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = (Get-Sha256Hex ((Get-FileFingerprint $path) + '|' + (Get-DeckFingerprint $pres))); presentation = (Get-PresentationModel $pres); path = $path } }
				}
			}

			'close_deck' {
				$discard = [bool](Get-Field $req 'discard')
				$pres = Get-Presentation (Get-Field $req 'presentation') $true (-not $discard)
				$fingerprint = Get-DeckFingerprint $pres
				$name = [string]$pres.Name
				$path = ''
				try { $path = [string]$pres.FullName } catch { }
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; presentation = (Get-PresentationModel $pres); closed = $name; path = $path } }
				}
				else {
					$script:MUTATED = $true
					if ($discard) { $pres.Saved = -1 }
					else { $pres.Save() }
					$pres.Close()
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = $fingerprint; closed = $name; path = $path } }
				}
			}

			default { Fail ('unknown op: ' + $op) 'unknown_op' }
		}
	}
	catch {
		if ($script:RESPONDED) { return }
		$code = 'com_error'
		$msg = $_.Exception.Message
		if ($null -ne $_.Exception.InnerException) { $msg = $msg + ' / ' + $_.Exception.InnerException.Message }
		if ($null -ne $_.InvocationInfo -and -not [string]::IsNullOrEmpty($_.InvocationInfo.Line)) {
			$msg = $msg + ' | at: ' + $_.InvocationInfo.Line.Trim()
		}
		if (Test-Mutated) {
			$code = 'applied_unverified'
			$msg = 'Part or all of this write was applied; do not retry, read the deck again. ' + $msg
		}
		$payload = @{ ok = $false; error = ($msg -replace "`r?`n", ' '); code = $code }
		if ($code -eq 'applied_unverified') {
			$payload.data = @{ status = 'unknown'; tokenSpent = $true; appliedCount = $null; failedIndex = $null; shapeIds = $null; failedShapeMayExist = $null }
		}
		Write-Response $payload
	}
	finally {
		$script:RESP = $null
	}
}

# --- worker loop ------------------------------------------------------------

$script:WORK = $d
$__dhUtf8 = New-Object System.Text.UTF8Encoding($false)
$__dhIn = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $__dhUtf8, $false)
$__dhOut = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $__dhUtf8)
$__dhOut.AutoFlush = $true

$null = '{}' | ConvertFrom-Json
$null = @{ a = 1 } | ConvertTo-Json -Compress

$__dhOut.WriteLine('deskhand-ready 1 ' + $PID)
while ($null -ne ($__dhLine = $__dhIn.ReadLine())) {
	$__dhStem = $__dhLine.Trim()
	if ($__dhStem -notmatch '^[0-9A-Za-z-]+$') { continue }
	try {
		$null = Invoke-Request $__dhStem *> $null
	}
	catch {
		[Console]::Error.WriteLine('deskhand worker: ' + $_.Exception.Message)
	}
	$Error.Clear()
	$__dhOut.WriteLine('deskhand-done ' + $__dhStem)
}
Reset-Ppt
