# OneNote bridge for the deskhand `onenote` server. One powershell.exe runs this file for the whole life of
# the server process, started through a tiny encoded bootstrap:
#
#   powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand <base64 UTF-16LE of>
#     $ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue'
#     $d='<work dir>'
#     Invoke-Expression ([IO.File]::ReadAllText("$d\worker-<id>.ps1",[Text.Encoding]::UTF8))
#
# worker-<id>.ps1 is this file. It defines the functions below, then serves calls until stdin closes:
#   stdin   one call stem per line
#   stdout  "deskhand-ready 1 <windows pid>" once, then "deskhand-done <stem>" after each call
#   files   req-<stem>.json in and resp-<stem>.json out, both UTF-8, inside $d
# A fresh process per call cost ~0.6 s of start-up and exit plus ~0.3 s of reloading this file; a resident
# worker answers in a few milliseconds. The stem keeps each call on its own file pair.
#
# Never use `exit` anywhere in this file: it would end the worker between writing a response and printing
# "deskhand-done", and the caller would take a finished write for a lost one. Stop a call with Fail or return.
#
# The transport is shaped by four constraints found the hard way:
#   * WSL interop does not forward the Linux environment to Windows executables and
#     ignores WSLENV in this direction, so paths cannot ride in environment variables;
#     they ride in the encoded bootstrap instead.
#   * A Windows command line caps out near 32 K characters, so this file cannot be sent
#     through -EncodedCommand itself; only the ~200-character bootstrap is.
#   * Passing a script with -File makes Windows PowerShell 5.1 decode a BOM-less file as
#     ANSI and mangle every CJK literal, so this file is never run with -File.
#   * Page XML reaches 260 KB, so request and response travel through UTF-8 files in the
#     work directory rather than through arguments.
#
# XPath uses local-name() rather than a namespace manager on purpose: both XmlNameTable and
# XmlNamespaceManager implement IEnumerable, so passing or returning one makes PowerShell
# silently unroll it into namespace prefixes.
#
# Provided by the bootstrap: $d (work dir).
#   request  { "op": "<name>", ...args }
#   response { "ok": true, "data": {...} } | { "ok": false, "error": "...", "code": "..." }
# Every call writes a response; a missing response file is the only hard failure.
#
# Write contract: every write op is default-deny. It only writes when the request carries
# apply=true AND a non-empty expectFingerprint. The fingerprint is recomputed here, right before
# the COM write, from the live page (or section / page name), and a mismatch is refused with code
# changed_since_preview. Without apply=true the op is a preview and returns the fingerprint that a
# later apply must present. The fingerprint is a content hash, never lastModifiedTime.

$ErrorActionPreference = 'Stop'

$NS = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$XP_NOTEBOOK = '//*[local-name()="Notebook"]'
$XP_SECTION = '*[local-name()="Section"]'
$XP_PAGE = '*[local-name()="Page"]'
$XP_OE = '*[local-name()="OE"]'
$ProgressPreference = 'SilentlyContinue'

# Set per call by Invoke-Request. RESPONDED makes the first response final and tells the catch in
# Invoke-Request that Fail's throw is control flow, not an error: a .NET callback wraps a thrown exception,
# so its message cannot be matched reliably.
$script:RESP = $null
$script:RESPONDED = $false
# True once a COM write call has returned. From then on any failure of the same request is reported as
# applied_unverified, because part or all of the write is already in OneNote.
$script:MUTATED = $false

function Write-Response($payload) {
	$json = $payload | ConvertTo-Json -Depth 40 -Compress
	[IO.File]::WriteAllText($script:RESP, $json, [System.Text.UTF8Encoding]::new($false))
	$script:RESPONDED = $true
}

function Fail([string]$message, [string]$code) {
	if ($script:MUTATED) {
		$code = 'applied_unverified'
		$message = 'Part or all of this write was applied; do not retry, read the page again. ' + $message
	}
	Write-Response @{ ok = $false; error = $message; code = $code }
	throw 'deskhand:responded'
}

# --- text helpers -----------------------------------------------------------

function Get-CleanText([string]$s) {
	if ([string]::IsNullOrEmpty($s)) { return '' }
	# XML 1.0 forbids most C0 control characters; OneNote rejects pages containing them.
	return [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
}

# Plain text is stored as an HTML-ish fragment inside the CDATA, so a literal "<angles>" that is put
# in raw is read back as a tag and silently disappears. Escape & < > the way OneNote does itself
# ("->" is stored as "-&gt;"), and the text survives round-tripping verbatim.
function Get-EscapedText([string]$s) {
	$t = Get-CleanText $s
	$t = $t.Replace('&', '&amp;')
	$t = $t.Replace('<', '&lt;')
	$t = $t.Replace('>', '&gt;')
	return $t
}

function New-TextCdata([string]$s) {
	$t = (Get-EscapedText $s).Replace(']]>', ']]]]><![CDATA[>')
	return '<![CDATA[' + $t + ']]>'
}

# Hashtable property access falls through to the type's members, so a request field named like a
# Hashtable method ("remove") returns that method's overloads instead of $null when it is absent.
# Read every optional request field through here.
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

# OneNote stores rich text as an HTML-ish fragment *inside* the CDATA of one:T, for
# example <one:T><![CDATA[<span style='font-family:Calibri'>hi</span>]]></one:T>.
# Entities stay escaped even inside CDATA, so decoding is required as well as tag stripping.
function Get-RichText([string]$rich) {
	if ([string]::IsNullOrEmpty($rich)) { return '' }
	$s = [regex]::Replace($rich, '(?i)<br\s*/?>', "`n")
	$s = [regex]::Replace($s, '<[^>]*>', '')
	$s = [System.Net.WebUtility]::HtmlDecode($s)
	# Neither WebUtility nor HttpUtility decodes astral-plane references on .NET Framework,
	# so OneNote's &#128640; would survive as literal text. Finish the job by hand.
	$s = [regex]::Replace($s, '&#(\d{1,7});', { param($m) ConvertFrom-CodePoint $m.Groups[1].Value })
	$s = [regex]::Replace($s, '&#[xX]([0-9A-Fa-f]{1,6});', { param($m) ConvertFrom-CodePoint ('0x' + $m.Groups[1].Value) })
	$s = $s -replace "`r`n", "`n"
	$s = $s.Replace([string][char]0xA0, ' ')
	return $s.Trim("`n")
}

function ConvertFrom-CodePoint([string]$text) {
	$value = 0
	if ($text.StartsWith('0x')) { $value = [Convert]::ToInt32($text, 16) } else { $value = [int]$text }
	if ($value -lt 0 -or $value -gt 0x10FFFF) { return '' }
	return [char]::ConvertFromUtf32($value)
}

function Get-TRaw($tNode) {
	$x = $tNode.InnerXml
	$x = $x.Replace('<![CDATA[', '')
	$x = $x.Replace(']]>', '')
	return $x
}

function Get-OeText($oeNode) {
	$parts = New-Object System.Collections.ArrayList
	foreach ($t in $oeNode.SelectNodes('*[local-name()="T"]')) {
		[void]$parts.Add((Get-RichText (Get-TRaw $t)))
	}
	return ($parts -join "`n")
}

# --- xml helpers ------------------------------------------------------------

function Get-AncestorNames($node) {
	$names = New-Object System.Collections.ArrayList
	$p = $node.ParentNode
	while ($null -ne $p -and $null -ne $p.ParentNode) {
		[void]$names.Add($p.LocalName)
		$p = $p.ParentNode
	}
	return $names
}

function Get-NearestAncestor($node, [string]$localName) {
	$p = $node.ParentNode
	while ($null -ne $p -and $null -ne $p.ParentNode) {
		if ($p.LocalName -eq $localName) { return $p }
		$p = $p.ParentNode
	}
	return $null
}

function Get-Num($node, [string]$attrName) {
	if ($null -eq $node) { return $null }
	$raw = $node.GetAttribute($attrName)
	if ([string]::IsNullOrEmpty($raw)) { return $null }
	$dbl = 0.0
	if ([double]::TryParse($raw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$dbl)) { return $dbl }
	return $null
}

function Get-PageModel($pageNode) {
	$outlines = New-Object System.Collections.ArrayList
	foreach ($o in $pageNode.SelectNodes('.//*[local-name()="Outline"]')) {
		$pos = $o.SelectSingleNode('*[local-name()="Position"]')
		$size = $o.SelectSingleNode('*[local-name()="Size"]')
		[void]$outlines.Add(@{
				objectID  = $o.GetAttribute('objectID')
				x         = Get-Num $pos 'x'
				y         = Get-Num $pos 'y'
				z         = Get-Num $pos 'z'
				width     = Get-Num $size 'width'
				height    = Get-Num $size 'height'
				inTitle   = ($null -ne (Get-NearestAncestor $o 'Title'))
				lineCount = $o.SelectNodes('.//*[local-name()="OE"]').Count
			})
	}

	$lines = New-Object System.Collections.ArrayList
	$index = 0
	foreach ($oe in $pageNode.SelectNodes('.//*[local-name()="OE"]')) {		$ancestors = Get-AncestorNames $oe
		$depth = 0
		foreach ($a in $ancestors) { if ($a -eq 'OE') { $depth++ } }

		$list = ''
		if ($null -ne $oe.SelectSingleNode('*[local-name()="List"]/*[local-name()="Bullet"]')) { $list = 'bullet' }
		elseif ($null -ne $oe.SelectSingleNode('*[local-name()="List"]/*[local-name()="Number"]')) { $list = 'number' }

		$outline = Get-NearestAncestor $oe 'Outline'

		# Font family, size and colour are cascaded text attributes: OneNote puts them on the OE, a caller
		# may have put them on the T. Either way they are what read_page reports back.
		$style = $oe.GetAttribute('style')
		if ([string]::IsNullOrEmpty($style)) {
			$tNode = $oe.SelectSingleNode('*[local-name()="T"]')
			if ($null -ne $tNode) { $style = $tNode.GetAttribute('style') }
		}

		$line = @{
			index      = $index
			objectID   = $oe.GetAttribute('objectID')
			text       = (Get-OeText $oe)
			inTitle    = ($ancestors -contains 'Title')
			inTable    = ($ancestors -contains 'Table')
			hasImage   = ($null -ne $oe.SelectSingleNode('*[local-name()="Image"]'))
			hasTable   = ($null -ne $oe.SelectSingleNode('*[local-name()="Table"]'))
			hasInk     = ($null -ne $oe.SelectSingleNode('.//*[local-name()="InkDrawing"]'))
			list       = $list
			depth      = $depth
			outlineID  = $(if ($null -ne $outline) { $outline.GetAttribute('objectID') } else { '' })
			quickStyle = $oe.GetAttribute('quickStyleIndex')
			alignment  = $oe.GetAttribute('alignment')
		}
		# Only when the paragraph carries one, so a page without formatting reads back exactly as before.
		if (-not [string]::IsNullOrEmpty($style)) { $line['style'] = $style }
		[void]$lines.Add($line)
		$index++
	}

	$images = New-Object System.Collections.ArrayList
	foreach ($img in $pageNode.SelectNodes('.//*[local-name()="Image"]')) {
		$imgPos = $img.SelectSingleNode('*[local-name()="Position"]')
		$imgSize = $img.SelectSingleNode('*[local-name()="Size"]')
		$imgOutline = Get-NearestAncestor $img 'Outline'
		[void]$images.Add(@{
				objectID  = $img.GetAttribute('objectID')
				format    = $img.GetAttribute('format')
				alt       = $img.GetAttribute('alt')
				x         = Get-Num $imgPos 'x'
				y         = Get-Num $imgPos 'y'
				width     = Get-Num $imgSize 'width'
				height    = Get-Num $imgSize 'height'
				outlineID = $(if ($null -ne $imgOutline) { $imgOutline.GetAttribute('objectID') } else { '' })
				# OneNote replaces one:Data with a one:CallbackID when reading with piBasic, so an image that
				# was written is never read back as Data. Its bits live in the page's binary store and the
				# callback id is the evidence that they are there; format is not echoed back at all.
				embedded  = ($null -ne $img.SelectSingleNode('*[local-name()="Data"]')) -or ($null -ne $img.SelectSingleNode('*[local-name()="CallbackID"]'))
				callbackID = $(
					$cb = $img.SelectSingleNode('*[local-name()="CallbackID"]')
					if ($null -ne $cb) { $cb.GetAttribute('callbackID') } else { '' }
				)
			})
	}

	$tables = New-Object System.Collections.ArrayList
	foreach ($tbl in $pageNode.SelectNodes('.//*[local-name()="Table"]')) {
		$tblOutline = Get-NearestAncestor $tbl 'Outline'
		[void]$tables.Add(@{
				objectID  = $tbl.GetAttribute('objectID')
				rows      = $tbl.SelectNodes('*[local-name()="Row"]').Count
				columns   = $tbl.SelectNodes('*[local-name()="Row"][1]/*[local-name()="Cell"]').Count
				outlineID = $(if ($null -ne $tblOutline) { $tblOutline.GetAttribute('objectID') } else { '' })
			})
	}

	return @{
		id           = $pageNode.GetAttribute('ID')
		name         = $pageNode.GetAttribute('name')
		level        = $pageNode.GetAttribute('pageLevel')
		lang         = $pageNode.GetAttribute('lang')
		created      = $pageNode.GetAttribute('dateTime')
		lastModified = $pageNode.GetAttribute('lastModifiedTime')
		lines        = @($lines)
		outlines     = @($outlines)
		images       = @($images)
		tables       = @($tables)
	}
}

# --- COM helpers ------------------------------------------------------------

function Get-HierarchyXml($onenote, [string]$startId, [int]$scope) {
	$xml = ''
	$onenote.GetHierarchy($startId, $scope, [ref]$xml)
	return $xml
}

function Get-PageContentXml($onenote, [string]$pageId) {
	$xml = ''
	$onenote.GetPageContent($pageId, [ref]$xml, 0)
	return $xml
}

# Returns an ArrayList of hashtables. Callers must re-wrap the result in @(), because
# PowerShell unrolls a returned collection into individual items.
function Get-NotebookTree($onenote) {
	[xml]$doc = Get-HierarchyXml $onenote '' 3
	$notebooks = New-Object System.Collections.ArrayList
	foreach ($nb in $doc.SelectNodes($XP_NOTEBOOK)) {
		$sections = New-Object System.Collections.ArrayList
		foreach ($sec in $nb.SelectNodes('.//' + $XP_SECTION)) {
			[void]$sections.Add(@{
					id   = $sec.GetAttribute('ID')
					name = $sec.GetAttribute('name')
					path = $sec.GetAttribute('path')
				})
		}
		[void]$notebooks.Add(@{
				id       = $nb.GetAttribute('ID')
				name     = $nb.GetAttribute('name')
				path     = $nb.GetAttribute('path')
				sections = @($sections)
			})
	}
	return $notebooks
}

function Resolve-SectionId($onenote, [string]$sectionId, [string]$notebookName, [string]$sectionName) {
	if (-not [string]::IsNullOrEmpty($sectionId)) { return $sectionId }
	if ([string]::IsNullOrEmpty($sectionName)) {
		Fail 'provide sectionId, or Notebook plus Section names' 'invalid_request'
	}
	$matches = New-Object System.Collections.ArrayList
	foreach ($nb in @(Get-NotebookTree $onenote)) {
		if (-not [string]::IsNullOrEmpty($notebookName) -and $nb.name -ne $notebookName) { continue }
		foreach ($sec in $nb.sections) {
			if ($sec.name -eq $sectionName) {
				[void]$matches.Add(@{ notebook = $nb.name; section = $sec.name; id = $sec.id })
			}
		}
	}
	if ($matches.Count -eq 0) {
		$where = ''
		if (-not [string]::IsNullOrEmpty($notebookName)) { $where = " in notebook '" + $notebookName + "'" }
		Fail ("no section named '" + $sectionName + "'" + $where) 'not_found'
	}
	if ($matches.Count -gt 1) {
		$names = (($matches | ForEach-Object { $_.notebook }) -join ', ')
		Fail ("section name '" + $sectionName + "' is ambiguous across notebooks: " + $names + '; pass Notebook too') 'ambiguous'
	}
	return $matches[0].id
}

# GetHierarchy reports a stale lastModifiedTime - on a page edited many times it kept returning the
# creation time - while the page root from GetPageContent carries the live one. Anything that
# reports or compares lastModified must go through here, otherwise the staleness guard compares two
# different clocks and refuses writes that are in fact fresh.
function Read-PageMeta($onenote, [string]$pageId) {
	[xml]$doc = Get-PageContentXml $onenote $pageId
	$node = $doc.DocumentElement
	if ($null -eq $node -or $node.LocalName -ne 'Page') { Fail ("page id not found: " + $pageId) 'not_found' }
	return @{
		id           = $node.GetAttribute('ID')
		name         = $node.GetAttribute('name')
		lastModified = $node.GetAttribute('lastModifiedTime')
	}
}

function Resolve-Page($onenote, [string]$pageId, [string]$notebookName, [string]$sectionName, [string]$pageName) {
	if (-not [string]::IsNullOrEmpty($pageId)) {
		return Read-PageMeta $onenote $pageId
	}
	if ([string]::IsNullOrEmpty($pageName)) {
		Fail 'provide pageId, or Notebook, Section and Page names' 'invalid_request'
	}
	$secId = Resolve-SectionId $onenote '' $notebookName $sectionName
	[xml]$doc = Get-HierarchyXml $onenote $secId 4
	foreach ($pg in $doc.SelectNodes('//*[local-name()="Page"]')) {
		if ($pg.GetAttribute('name') -eq $pageName) {
			return Read-PageMeta $onenote $pg.GetAttribute('ID')
		}
	}
	Fail ("no page named '" + $pageName + "' in section '" + $sectionName + "'") 'not_found'
}

function Format-Double($value) {
	return ([double]$value).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)
}

# --- fingerprints ------------------------------------------------------------

function Get-Sha256Hex([string]$text) {
	$sha = [System.Security.Cryptography.SHA256]::Create()
	try {
		$bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text))
	}
	finally { $sha.Dispose() }
	return ([BitConverter]::ToString($bytes).Replace('-', '')).ToLowerInvariant()
}

# OneNote writes the UI selection (selected="...") into the XML it hands out, and the per-element
# lastModifiedTime moves independently of content, so neither may take part in a content hash.
function Get-NormalizedXml([string]$xml) {
	return [regex]::Replace($xml, '\s+(selected|lastModifiedTime)="[^"]*"', '')
}

# A write is allowed only with apply=true and the fingerprint that the preview reported.
function Assert-Fingerprint($req, [string]$fingerprint) {
	if ((Get-Field $req 'apply') -ne $true) { return }
	$expect = [string](Get-Field $req 'expectFingerprint')
	if ([string]::IsNullOrEmpty($expect)) {
		Fail 'refusing to write without expectFingerprint: run the preview first and pass its fingerprint' 'apply_requires_fingerprint'
	}
	if ($expect -ne $fingerprint) {
		Fail 'the target changed since the preview was taken; re-read it and preview again' 'changed_since_preview'
	}
}

function Test-Apply($req) { return ((Get-Field $req 'apply') -eq $true) }

# --- tables and images ------------------------------------------------------
# Schema notes that cost real debugging time:
#   * one:Table is a child of one:OE, never of one:Outline (so it travels inside an OE);
#   * Table is Columns? then Row+, and every one:Cell MUST contain an OEChildren;
#   * one:Column needs both index and width, so Columns is emitted only when widths are given;
#   * there is no columnSpan/rowSpan in the schema at all;
#   * one:Image accepts base64 in one:Data - no callback object is involved - and its format
#     pattern is auto|png|emf|jpg, so "jpg" is spelled without the e.

function New-TableOeXml([string[]]$headers, $rows, [bool]$bordersVisible, [bool]$hasHeaderRow, $columnWidths) {
	$grid = New-Object System.Collections.ArrayList
	if ($headers.Count -gt 0) { [void]$grid.Add($headers) }
	foreach ($row in @($rows)) {
		$cells = New-Object System.Collections.ArrayList
		foreach ($cell in @($row)) { [void]$cells.Add([string]$cell) }
		[void]$grid.Add(@($cells))
	}
	if ($grid.Count -eq 0) { Fail 'a table needs at least one row' 'invalid_request' }

	$columns = 0
	foreach ($row in $grid) { if ($row.Count -gt $columns) { $columns = $row.Count } }
	if ($columns -eq 0) { Fail 'a table needs at least one column' 'invalid_request' }

	$sb = New-Object System.Text.StringBuilder
	[void]$sb.Append('<one:OE><one:Table bordersVisible="' + $(if ($bordersVisible) { 'true' } else { 'false' }) + '" hasHeaderRow="' + $(if ($hasHeaderRow) { 'true' } else { 'false' }) + '">')

	$widths = @($columnWidths)
	if ($widths.Count -gt 0) {
		[void]$sb.Append('<one:Columns>')
		for ($i = 0; $i -lt $widths.Count; $i++) {
			[void]$sb.Append('<one:Column index="' + $i + '" width="' + (Format-Double $widths[$i]) + '"/>')
		}
		[void]$sb.Append('</one:Columns>')
	}

	$rowIndex = 0
	foreach ($row in $grid) {
		[void]$sb.Append('<one:Row>')
		for ($c = 0; $c -lt $columns; $c++) {
			$value = ''
			if ($c -lt $row.Count) { $value = [string]$row[$c] }
			[void]$sb.Append('<one:Cell')
			if ($hasHeaderRow -and $rowIndex -eq 0) { [void]$sb.Append(' shadingColor="#DEEBF7"') }
			[void]$sb.Append('><one:OEChildren><one:OE><one:T>' + (New-TextCdata $value) + '</one:T></one:OE></one:OEChildren></one:Cell>')
		}
		[void]$sb.Append('</one:Row>')
		$rowIndex++
	}

	[void]$sb.Append('</one:Table></one:OE>')
	return $sb.ToString()
}

function New-ImageOeXml([string]$base64, [string]$format, $width, $height, [string]$alt) {
	$clean = [regex]::Replace($base64, '\s', '')
	if ([string]::IsNullOrEmpty($clean)) { Fail 'no image data supplied' 'invalid_request' }
	$fmt = 'auto'
	if (-not [string]::IsNullOrEmpty($format)) { $fmt = $format.ToLowerInvariant() }
	if ($fmt -eq 'jpeg') { $fmt = 'jpg' }

	$attrs = ' format="' + $fmt + '"'
	if (-not [string]::IsNullOrEmpty($alt)) { $attrs = $attrs + ' alt="' + (Get-CleanText $alt).Replace('"', '&quot;') + '"' }

	$size = ''
	if ($null -ne $width -and $null -ne $height) {
		$size = '<one:Size width="' + (Format-Double $width) + '" height="' + (Format-Double $height) + '"/>'
	}
	return '<one:OE><one:Image' + $attrs + '>' + $size + '<one:Data>' + $clean + '</one:Data></one:Image></one:OE>'
}

# OneNote rejects a page whose new element carries an objectID it did not issue, so new
# outlines and OEs must be written without one and let OneNote assign them.
function New-BlockOutlineXml([string[]]$blocks) {
	$sb = New-Object System.Text.StringBuilder
	[void]$sb.Append('<one:Outline><one:OEChildren>')
	foreach ($block in $blocks) { [void]$sb.Append($block) }
	[void]$sb.Append('</one:OEChildren></one:Outline>')
	return $sb.ToString()
}

# --- paragraphs -------------------------------------------------------------
# A caller describes a paragraph tree as a flat list that carries a depth per entry: depth d hangs
# under the nearest earlier paragraph at depth d-1. OneNote expresses the same tree by nesting
# one:OEChildren, so the flat list is expanded into that shape here. Formatting goes where OneNote
# itself puts it: font family, size and colour on the one:OE style attribute (the cascaded text
# attributes), and the character formatting that only a run can carry (bold, italic, underline,
# strike, highlight) in a span inside the CDATA.

$QUICK_STYLE_ALIASES = @{ 'normal' = 'p'; 'paragraph' = 'p'; 'title' = 'pagetitle' }

# Reads lines, which are either plain strings or { text, depth, list, quickStyle, alignment,
# spaceBefore, spaceAfter, style } objects, into the flat form the tree builder expects.
function Get-Paragraphs($req) {
	$out = New-Object System.Collections.ArrayList
	$previous = -1
	foreach ($l in @(Get-Field $req 'lines')) {
		if ($null -eq $l) { continue }
		$p = @{ text = ''; depth = 0; list = ''; quickStyle = $null; alignment = ''; spaceBefore = $null; spaceAfter = $null; style = $null }
		if ($l -is [string]) { $p.text = $l }
		else {
			$p.text = [string](Get-Field $l 'text')
			$depth = Get-Field $l 'depth'
			if ($null -ne $depth) { $p.depth = [int]$depth }
			$p.list = [string](Get-Field $l 'list')
			$p.quickStyle = Get-Field $l 'quickStyle'
			$p.alignment = [string](Get-Field $l 'alignment')
			$p.spaceBefore = Get-Field $l 'spaceBefore'
			$p.spaceAfter = Get-Field $l 'spaceAfter'
			$p.style = Get-Field $l 'style'
		}
		if ($p.depth -lt 0) { Fail 'paragraph depth cannot be negative' 'invalid_request' }
		if ($out.Count -eq 0 -and $p.depth -ne 0) { Fail 'the first paragraph must be at depth 0' 'invalid_request' }
		if ($p.depth -gt $previous + 1) { Fail ('paragraph depth ' + $p.depth + ' skips a level after depth ' + $previous) 'invalid_request' }
		$previous = $p.depth
		[void]$out.Add($p)
	}
	return @($out)
}

function Test-QuickStyleNames($paragraphs) {
	foreach ($p in $paragraphs) {
		if ($null -eq $p.quickStyle) { continue }
		if ($p.quickStyle -is [string] -and $p.quickStyle -notmatch '^\d+$') { return $true }
	}
	return $false
}

# A quickStyleIndex is an index into the page's own QuickStyleDef list, not a global enum, so a name
# means whatever the page says it means. A number is passed through: callers copy it from read_page.
function Resolve-QuickStyleIndex($pageNode, $value) {
	if ($null -eq $value) { return $null }
	if (-not ($value -is [string])) { return [int]$value }
	$name = $value.Trim()
	if ($name -match '^\d+$') { return [int]$name }
	$wanted = $name.ToLowerInvariant()
	if ($QUICK_STYLE_ALIASES.ContainsKey($wanted)) { $wanted = $QUICK_STYLE_ALIASES[$wanted] }
	$defs = @{}
	if ($null -ne $pageNode) {
		foreach ($def in $pageNode.SelectNodes('*[local-name()="QuickStyleDef"]')) {
			$key = $def.GetAttribute('name').ToLowerInvariant()
			if (-not $defs.ContainsKey($key)) { $defs[$key] = [int]$def.GetAttribute('index') }
		}
	}
	if ($defs.ContainsKey($wanted)) { return $defs[$wanted] }
	$known = (@($defs.Keys) | Sort-Object) -join ', '
	if ([string]::IsNullOrEmpty($known)) { $known = 'none' }
	Fail ('no quick style named "' + $name + '" on this page; it defines: ' + $known + '. Pass the index from read_page instead.') 'invalid_request'
}

# Returns the one:OE open tag, its list and its text, but not the closing tag: the tree builder closes
# it after any children, because a paragraph's children live inside its own OE.
function New-ParagraphOeXml($p, $pageNode) {
	$attrs = ''
	$style = $p.style
	if ($null -ne $style) {
		$css = New-Object System.Collections.ArrayList
		$family = [string](Get-Field $style 'fontFamily')
		if (-not [string]::IsNullOrEmpty($family)) { [void]$css.Add('font-family:' + $family) }
		$fontSize = Get-Field $style 'fontSize'
		if ($null -ne $fontSize) { [void]$css.Add('font-size:' + (Format-Double $fontSize) + 'pt') }
		$color = [string](Get-Field $style 'color')
		if (-not [string]::IsNullOrEmpty($color)) { [void]$css.Add('color:' + $color) }
		if ($css.Count -gt 0) { $attrs = $attrs + ' style="' + ($css -join ';') + '"' }
	}
	$quickStyle = Resolve-QuickStyleIndex $pageNode $p.quickStyle
	if ($null -ne $quickStyle) { $attrs = $attrs + ' quickStyleIndex="' + $quickStyle + '"' }
	if (-not [string]::IsNullOrEmpty($p.alignment)) { $attrs = $attrs + ' alignment="' + $p.alignment + '"' }
	if ($null -ne $p.spaceBefore) { $attrs = $attrs + ' spaceBefore="' + (Format-Double $p.spaceBefore) + '"' }
	if ($null -ne $p.spaceAfter) { $attrs = $attrs + ' spaceAfter="' + (Format-Double $p.spaceAfter) + '"' }

	$list = ''
	if ($p.list -eq 'bullet') { $list = '<one:List><one:Bullet bullet="2"/></one:List>' }
	# numberFormat is a pattern, not a name, and text is the rendered number OneNote starts from.
	elseif ($p.list -eq 'number') { $list = '<one:List><one:Number numberSequence="0" numberFormat="##." text="1."/></one:List>' }

	$inner = Get-EscapedText $p.text
	$spans = New-Object System.Collections.ArrayList
	if ($null -ne $style) {
		if ((Get-Field $style 'bold') -eq $true) { [void]$spans.Add('font-weight:bold') }
		if ((Get-Field $style 'italic') -eq $true) { [void]$spans.Add('font-style:italic') }
		$decoration = New-Object System.Collections.ArrayList
		if ((Get-Field $style 'underline') -eq $true) { [void]$decoration.Add('underline') }
		if ((Get-Field $style 'strikethrough') -eq $true) { [void]$decoration.Add('line-through') }
		if ($decoration.Count -gt 0) { [void]$spans.Add('text-decoration:' + ($decoration -join ' ')) }
		$highlight = [string](Get-Field $style 'highlight')
		if (-not [string]::IsNullOrEmpty($highlight)) { [void]$spans.Add('background:' + $highlight) }
	}
	if ($spans.Count -gt 0) { $inner = "<span style='" + ($spans -join ';') + "'>" + $inner + '</span>' }
	$inner = $inner.Replace(']]>', ']]]]><![CDATA[>')
	return '<one:OE' + $attrs + '>' + $list + '<one:T><![CDATA[' + $inner + ']]></one:T>'
}

# Expands the flat (depth) list into nested one:OEChildren and returns the whole container, so that
# both a new outline and an append into an existing one can use it.
function New-ParagraphsXml($paragraphs, $pageNode) {
	$sb = New-Object System.Text.StringBuilder
	[void]$sb.Append('<one:OEChildren>')
	# One flag per open level: whether that paragraph's OEChildren has been opened.
	$levels = New-Object System.Collections.ArrayList
	foreach ($p in $paragraphs) {
		$depth = [int]$p.depth
		if ($depth -gt $levels.Count) { Fail ('paragraph depth ' + $depth + ' skips a level') 'invalid_request' }
		if ($depth -eq $levels.Count) {
			if ($levels.Count -gt 0) {
				[void]$sb.Append('<one:OEChildren>')
				$levels[$levels.Count - 1] = $true
			}
			[void]$levels.Add($false)
		}
		else {
			for ($i = $levels.Count - 1; $i -ge $depth; $i--) {
				if ($levels[$i]) { [void]$sb.Append('</one:OEChildren>') }
				[void]$sb.Append('</one:OE>')
				$levels.RemoveAt($i)
			}
			[void]$levels.Add($false)
		}
		[void]$sb.Append((New-ParagraphOeXml $p $pageNode))
	}
	for ($i = $levels.Count - 1; $i -ge 0; $i--) {
		if ($levels[$i]) { [void]$sb.Append('</one:OEChildren>') }
		[void]$sb.Append('</one:OE>')
	}
	[void]$sb.Append('</one:OEChildren>')
	return $sb.ToString()
}

function New-OutlineXml($paragraphs, $position, $pageNode) {
	$sb = New-Object System.Text.StringBuilder
	[void]$sb.Append('<one:Outline>')
	if ($null -ne $position) {
		[void]$sb.Append('<one:Position x="' + (Format-Double $position.x) + '" y="' + (Format-Double $position.y) + '" z="' + [int]$position.z + '"/>')
		if ($null -ne $position.width) {
			[void]$sb.Append('<one:Size width="' + (Format-Double $position.width) + '" height="' + (Format-Double $position.height) + '"/>')
		}
	}
	[void]$sb.Append((New-ParagraphsXml $paragraphs $pageNode))
	[void]$sb.Append('</one:Outline>')
	return $sb.ToString()
}

# The real signature is UpdatePageContent(string xml, DateTime dateExpectedLastModified,
# XMLSchema xsSchema, bool force). There is no error output: OneNote reports a refusal by throwing.
# force is always false, so a page holding data from a newer OneNote is refused instead of
# overwritten (0x80042022 is "content loss"). xsSchema must be 2 (xs2013) because this bridge writes
# the 2013 namespace: 0 and 1 are rejected with 0x80042001.
#
# dateExpectedLastModified is deliberately always MinValue, which means "do not check". It cannot
# be turned into a concurrency guard: lastModifiedTime does not advance when content is written
# through COM (three appends in a row left it at the page's creation second) and GetHierarchy can
# report a different value than GetPageContent for the same page. Feeding back the value just read
# therefore protects nothing while refusing legitimate writes. Appends are additive and cannot
# clobber concurrent edits; delete_page guards on the page name instead.
$XSCHEMA_2013 = 2

# HRESULTs are compared as upper-case strings ('0x80042010'): a hex literal below 0x80000000 is an Int32 in
# PowerShell 5.1 and one above it is negative, so mixing literals and formatted values invites mismatches.
$HR_STALE_PAGE = '0x80042010'
# The call was turned away before OneNote ran it, so nothing was written.
$HR_BUSY = @('0x80010001', '0x8001010A')
$HR_RECORDING = '0x80042024'
# Refusals OneNote is known to make: the request was rejected and nothing was written. The rest of the
# 0x800420xx range is not assumed, because a code this bridge has never seen may follow a partial write.
# 0x80042023 is OneNote's own timeout and 0x80042024 means recording, so neither is a refusal.
$HR_REFUSED = @('0x80042001', '0x8004200E', '0x80042014', '0x80042022')
# A COM diagnostic can be arbitrarily long and a client cuts a long reply in the middle.
$DIAGNOSTIC_MAX = 300
# The server went away or the connection broke, possibly after it applied the write.
$HR_GONE = @('0x800706BA', '0x800706BE', '0x800706BF', '0x80010108', '0x80010012', '0x80010007', '0x800401FD')

# The exception that actually failed: the first one that is not a PowerShell or reflection wrapper, or $null
# when the failure never reached a COM call. MethodInvocationException is what PowerShell puts around an
# exception thrown by a method, so it is looked through. Any other exception PowerShell raised itself (a failed
# conversion, a binding error) implements IContainsErrorRecord and means no call was made, whatever it wraps; a
# COM exception does not implement it.
function Get-FailureException($errorRecord) {
	$e = $errorRecord.Exception
	while ($null -ne $e) {
		if ($e -is [Management.Automation.MethodInvocationException] -or $e -is [Reflection.TargetInvocationException] -or $e.GetType() -eq [Management.Automation.RuntimeException]) {
			$e = $e.InnerException
		}
		elseif ($e -is [Management.Automation.IContainsErrorRecord]) { return $null }
		else { break }
	}
	return $e
}

function Get-HResultHex($errorRecord) {
	$e = Get-FailureException $errorRecord
	if ($null -eq $e) { return '' }
	# Casting a negative Int32 straight to UInt32 throws in PowerShell; go through BitConverter.
	$unsigned = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$e.HResult), 0)
	return ('0x{0:X8}' -f $unsigned)
}

# Only a failure OneNote is known to have refused counts as a refusal. Anything unrecognised during a write
# is outcome_unknown, because the write may have landed.
function Get-WriteFailureCode([string]$hex) {
	if ([string]::IsNullOrEmpty($hex)) { return 'com_error' }
	if ($hex -in $HR_BUSY -or $hex -eq $HR_RECORDING) { return 'busy' }
	if ($hex -eq $HR_STALE_PAGE) { return 'stale' }
	if ($hex -in $HR_REFUSED) { return 'rejected' }
	return 'outcome_unknown'
}

# The classification and the HRESULT stay; only the free text is shortened, and never inside a surrogate pair.
function Limit-Diagnostic([string]$text) {
	if ($text.Length -le $DIAGNOSTIC_MAX) { return $text }
	$cut = $DIAGNOSTIC_MAX
	if ([char]::IsHighSurrogate($text[$cut - 1])) { $cut = $cut - 1 }
	return $text.Substring(0, $cut) + '...'
}

function Fail-Write($errorRecord, [string]$action) {
	$hex = Get-HResultHex $errorRecord
	$code = Get-WriteFailureCode $hex
	# The wrapper PowerShell puts around a COM exception is noise; the reason is the inner message.
	$failure = Get-FailureException $errorRecord
	$why = if ($null -eq $failure) { $errorRecord.Exception.Message } else { $failure.Message }
	$why = Limit-Diagnostic $why
	switch ($code) {
		'stale' { Fail ('OneNote refused the update as stale (' + $hex + '); re-read the page and retry') $code }
		'busy' { Fail ('OneNote is busy (' + $hex + ') and did not ' + $action + '; try again shortly') $code }
		'rejected' { Fail ('OneNote refused to ' + $action + ' (' + $hex + '): ' + $why) $code }
		'outcome_unknown' {
			# The guidance comes first so a client that cuts the reply cannot remove it.
			Fail ('Read the page again and preview before retrying: whether OneNote managed to ' + $action +
				' is unknown (' + $hex + '): ' + $why) $code
		}
		default { Fail ('Could not ' + $action + ': ' + $why) $code }
	}
}

function Send-PageUpdate($onenote, [string]$pageId, [string]$innerXml) {
	$xml = '<one:Page xmlns:one="' + $NS + '" ID="' + $pageId + '">' + $innerXml + '</one:Page>'

	try {
		$onenote.UpdatePageContent($xml, [DateTime]::MinValue, $XSCHEMA_2013, $false)
	}
	catch { Fail-Write $_ 'update the page' }
	$script:MUTATED = $true
}

function Read-PageById($onenote, [string]$pageId) {
	[xml]$doc = Get-PageContentXml $onenote $pageId
	return Get-PageModel $doc.DocumentElement
}

# --- in-place editing -------------------------------------------------------
# What one:OEChildren means to UpdatePageContent, all verified against the live client:
#   * the element referenced by objectID is authoritative for its children, so the one:OEChildren
#     that is sent becomes the COMPLETE child list and any paragraph left out is deleted;
#   * the text of a kept paragraph is changed by resending it under its own objectID;
#   * a new paragraph is added by sending it without an objectID, anywhere in that list;
#   * a referenced outline keeps its Position and Size when they are omitted;
#   * page elements that are not mentioned at all (the title, other outlines) are preserved;
#   * an empty one:OEChildren is rejected, so emptying an outline cannot delete it.
# The unit of edit is therefore one whole outline, resent in full. Everything the caller is not
# changing is echoed back from OneNote's own XML, so styles, lists, tables and images survive.

function Get-OeByObjectId($pageNode, [string]$objectID) {
	foreach ($oe in $pageNode.SelectNodes('.//*[local-name()="OE"]')) {
		if ($oe.GetAttribute('objectID') -eq $objectID) { return $oe }
	}
	return $null
}

function Get-OutlineByObjectId($pageNode, [string]$objectID) {
	foreach ($o in $pageNode.SelectNodes('.//*[local-name()="Outline"]')) {
		if ($o.GetAttribute('objectID') -eq $objectID) { return $o }
	}
	return $null
}

function Get-OwnerOutline($node) {
	$p = $node.ParentNode
	while ($null -ne $p) {
		if ($p.LocalName -eq 'Outline') { return $p }
		$p = $p.ParentNode
	}
	return $null
}

# $newText must stay untyped: declaring it [string] makes PowerShell convert a $null argument into
# "", which turns "leave this paragraph alone" into "replace its text with nothing".
function Get-ChildXml($childNode, $newText) {
	# OuterXml re-declares the namespace on every serialized element: valid, but noisy.
	$xml = [regex]::Replace($childNode.OuterXml, '\s+xmlns:one="[^"]*"', '')
	if ($null -eq $newText) { return $xml }
	$match = [regex]::Match($xml, '<one:T\b[^>]*>.*?</one:T>', 'Singleline')
	if (-not $match.Success) { return $null }
	return $xml.Substring(0, $match.Index) + '<one:T>' + (New-TextCdata ([string]$newText)) + '</one:T>' + $xml.Substring($match.Index + $match.Length)
}

# Returns a PSCustomObject rather than a hashtable: a hashtable is IEnumerable, so PowerShell would
# unroll one returned from a function into DictionaryEntry objects.
function New-EditPlan($pageNode, $req, [bool]$allowRawXml = $false) {
	$plan = @{}

	foreach ($item in @(Get-Field $req 'set')) {
		if ($null -eq $item) { continue }
		$oeId = [string](Get-Field $item 'objectID')
		if ([string]::IsNullOrEmpty($oeId)) { Fail 'every set entry needs objectID and text' 'invalid_request' }
		$oe = Get-OeByObjectId $pageNode $oeId
		if ($null -eq $oe) { Fail ('no paragraph with objectID ' + $oeId + ' on this page') 'not_found' }
		$outline = Get-OwnerOutline $oe
		if ($null -eq $outline) { Fail ('paragraph ' + $oeId + ' is not inside an outline') 'unsupported' }
		$outlineId = $outline.GetAttribute('objectID')
		if (-not $plan.ContainsKey($outlineId)) { $plan[$outlineId] = @{ set = @{}; remove = @{}; append = @() } }
		$plan[$outlineId].set[$oeId] = [string](Get-Field $item 'text')
	}

	foreach ($item in @(Get-Field $req 'remove')) {
		$oeId = [string]$item
		if ([string]::IsNullOrEmpty($oeId)) { continue }
		$oe = Get-OeByObjectId $pageNode $oeId
		if ($null -eq $oe) { Fail ('no paragraph with objectID ' + $oeId + ' on this page') 'not_found' }
		$outline = Get-OwnerOutline $oe
		if ($null -eq $outline) { Fail ('paragraph ' + $oeId + ' is not inside an outline') 'unsupported' }
		$outlineId = $outline.GetAttribute('objectID')
		if (-not $plan.ContainsKey($outlineId)) { $plan[$outlineId] = @{ set = @{}; remove = @{}; append = @() } }
		$plan[$outlineId].remove[$oeId] = $true
	}

	foreach ($item in @(Get-Field $req 'append')) {
		if ($null -eq $item) { continue }
		$target = [string](Get-Field $item 'objectID')
		$blocks = New-Object System.Collections.ArrayList
		$paragraphs = @(Get-Paragraphs $item)
		if ($paragraphs.Count -gt 0) {
			# New-ParagraphsXml returns the whole OEChildren; the outline being resent already has one.
			$xml = New-ParagraphsXml $paragraphs $pageNode
			$xml = $xml.Substring('<one:OEChildren>'.Length)
			$xml = $xml.Substring(0, $xml.Length - '</one:OEChildren>'.Length)
			$labels = New-Object System.Collections.ArrayList
			foreach ($p in $paragraphs) { [void]$labels.Add($p.text) }
			[void]$blocks.Add(@{ xml = $xml; labels = @($labels) })
		}
		foreach ($x in @(Get-Field $item 'xml')) {
			if ($null -eq $x) { continue }
			if (-not $allowRawXml) { Fail 'raw xml is not accepted in an edit; use lines' 'invalid_request' }
			[void]$blocks.Add(@{ xml = [string]$x; labels = @('[raw block]') })
		}
		if ($blocks.Count -eq 0) { Fail 'every append entry needs lines or xml' 'invalid_request' }

		$outline = Get-OutlineByObjectId $pageNode $target
		if ($null -eq $outline) {
			$oe = Get-OeByObjectId $pageNode $target
			if ($null -ne $oe) { $outline = Get-OwnerOutline $oe }
		}
		if ($null -eq $outline) { Fail ('no outline or paragraph with objectID ' + $target + ' on this page') 'not_found' }
		$outlineId = $outline.GetAttribute('objectID')
		if (-not $plan.ContainsKey($outlineId)) { $plan[$outlineId] = @{ set = @{}; remove = @{}; append = @() } }
		$plan[$outlineId].append += , @{ blocks = @($blocks); after = [string](Get-Field $item 'after') }
	}

	if ($plan.Count -eq 0) { Fail 'nothing to change: pass set, remove or append' 'invalid_request' }

	$inner = New-Object System.Text.StringBuilder
	$changes = New-Object System.Collections.ArrayList

	$outlineIds = [string[]]@($plan.Keys)
	[Array]::Sort($outlineIds, [StringComparer]::Ordinal)
	foreach ($outlineId in $outlineIds) {
		$outline = Get-OutlineByObjectId $pageNode $outlineId
		if ($null -eq $outline) { Fail ('outline ' + $outlineId + ' is not on the page') 'not_found' }
		$children = $outline.SelectSingleNode('*[local-name()="OEChildren"]')
		if ($null -eq $children) { Fail ('outline ' + $outlineId + ' has no OEChildren') 'unsupported' }

		$setMap = $plan[$outlineId].set
		$removeMap = $plan[$outlineId].remove
		$parts = New-Object System.Collections.ArrayList
		$changed = New-Object System.Collections.ArrayList
		$removed = New-Object System.Collections.ArrayList
		$after = New-Object System.Collections.ArrayList

		foreach ($child in $children.ChildNodes) {
			if ($child.NodeType -ne 'Element') { continue }
			$childId = $child.GetAttribute('objectID')
			$before = ''
			if ($child.LocalName -eq 'OE') { $before = Get-OeText $child } else { $before = '[' + $child.LocalName.ToLowerInvariant() + ']' }

			if (-not [string]::IsNullOrEmpty($childId) -and $removeMap.ContainsKey($childId)) {
				[void]$removed.Add(@{ objectID = $childId; text = $before })
				continue
			}

			$newText = $null
			if (-not [string]::IsNullOrEmpty($childId) -and $setMap.ContainsKey($childId)) { $newText = [string]$setMap[$childId] }
			$xml = Get-ChildXml $child $newText
			if ($null -eq $xml) {
				Fail ('paragraph ' + $childId + ' has no one:T element, so its text cannot be replaced') 'unsupported'
			}
			[void]$parts.Add($xml)
			if ($null -ne $newText -and $newText -ne $before) {
				[void]$changed.Add(@{ objectID = $childId; before = $before; after = $newText })
				[void]$after.Add($newText)
			}
			else { [void]$after.Add($before) }
		}

		foreach ($entry in $plan[$outlineId].append) {
			$anchor = [string]$entry.after
			if ([string]::IsNullOrEmpty($anchor)) {
				foreach ($block in $entry.blocks) {
					[void]$parts.Add($block.xml)
					foreach ($label in $block.labels) { [void]$after.Add($label) }
				}
				continue
			}

			# An anchor may sit at any depth, but only the outline's top-level children are rebuilt here,
			# so the block is spliced into the child that contains the anchor. The anchor's own XML is a
			# unique substring of that child (objectIDs are unique), which makes the splice exact.
			$anchorOe = Get-OeByObjectId $outline $anchor
			if ($null -eq $anchorOe) { Fail ('no paragraph with objectID ' + $anchor + ' in outline ' + $outlineId) 'not_found' }
			if ($setMap.ContainsKey($anchor) -or $removeMap.ContainsKey($anchor)) {
				Fail ('cannot insert after ' + $anchor + ' while the same call changes or removes it') 'invalid_request'
			}
			$anchorXml = [regex]::Replace($anchorOe.OuterXml, '\s+xmlns:one="[^"]*"', '')
			$blockXml = ''
			foreach ($block in $entry.blocks) { $blockXml = $blockXml + $block.xml }
			$placed = $false
			for ($i = 0; $i -lt $parts.Count; $i++) {
				$at = $parts[$i].IndexOf($anchorXml, [StringComparison]::Ordinal)
				if ($at -lt 0) { continue }
				if ($parts[$i].IndexOf($anchorXml, $at + $anchorXml.Length, [StringComparison]::Ordinal) -ge 0) {
					Fail ('paragraph ' + $anchor + ' appears more than once in its outline') 'unsupported'
				}
				$parts[$i] = $parts[$i].Substring(0, $at + $anchorXml.Length) + $blockXml + $parts[$i].Substring($at + $anchorXml.Length)
				$placed = $true
				break
			}
			if (-not $placed) { Fail ('could not find paragraph ' + $anchor + ' in the outline being resent') 'not_found' }
			foreach ($block in $entry.blocks) {
				foreach ($label in $block.labels) { [void]$after.Add($label) }
			}
		}

		if ($parts.Count -eq 0) {
			Fail ('refusing to leave outline ' + $outlineId + ' empty: OneNote rejects an empty OEChildren, so remove some paragraphs instead of all of them') 'unsupported'
		}

		[void]$inner.Append('<one:Outline objectID="' + $outlineId + '"><one:OEChildren>' + ($parts -join '') + '</one:OEChildren></one:Outline>')
		[void]$changes.Add(@{
				outlineId = $outlineId
				changed   = @($changed)
				removed   = @($removed)
				after     = @($after)
			})
	}

	return [pscustomobject]@{ inner = $inner.ToString(); changes = @($changes) }
}

function Get-PlanFingerprint([string]$pageId, $plan) {
	return Get-Sha256Hex ('outlines:' + $pageId + "`n" + (Get-NormalizedXml $plan.inner))
}

function Get-RequestedLines($req) {
	$lines = New-Object System.Collections.ArrayList
	foreach ($l in @($req.lines)) { if ($null -ne $l) { [void]$lines.Add([string]$l) } }
	return @($lines)
}

# --- OneNote connection -----------------------------------------------------

# The application object is kept for the life of the worker. When the user has not opened OneNote, COM starts
# ONENOTE.EXE for the first object and ends it when the last one is released, so creating and releasing an
# object per call restarted OneNote every time (~0.9 s) and sometimes raced its shutdown (0x800706BF). A kept
# object is probed before each call and replaced when OneNote has gone away since the last one.
$script:OneNoteApp = $null

function Reset-OneNote {
	if ($null -ne $script:OneNoteApp) {
		try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:OneNoteApp) } catch {}
		$script:OneNoteApp = $null
	}
}

function New-OneNoteApp { New-Object -ComObject OneNote.Application }

# How long a busy OneNote (a modal dialog, a sync in progress) is waited for before giving up with 'busy'.
$PROBE_BUSY_S = 5

function Get-OneNote {
	if ($null -ne $script:OneNoteApp) {
		$deadline = [DateTime]::UtcNow.AddSeconds($PROBE_BUSY_S)
		while ($true) {
			try {
				$probe = ''
				$script:OneNoteApp.GetHierarchy('', 0, [ref]$probe)
				return $script:OneNoteApp
			}
			catch {
				if ((Get-HResultHex $_) -in $HR_BUSY) {
					# A busy OneNote is alive, so the object stays; only a dead one is replaced.
					if ([DateTime]::UtcNow -ge $deadline) { Fail 'OneNote is busy (a dialog may be open); try again shortly' 'busy' }
					Start-Sleep -Milliseconds 250
					continue
				}
				Reset-OneNote
				break
			}
		}
	}
	# Starting OneNote can race its previous shutdown (0x800706BF), so a failure gets one more try.
	try { $script:OneNoteApp = New-OneNoteApp }
	catch {
		Start-Sleep -Seconds 1
		$script:OneNoteApp = New-OneNoteApp
	}
	return $script:OneNoteApp
}

# --- main -------------------------------------------------------------------

function Invoke-Request([string]$stem) {
	$script:RESP = Join-Path $script:WORK ('resp-' + $stem + '.json')
	$script:RESPONDED = $false
	$script:MUTATED = $false
	# Unassigned names resolve up the dynamic scope chain, so the worker's globals must not leak in.
	$onenote = $null
	try {
		$req = [IO.File]::ReadAllText((Join-Path $script:WORK ('req-' + $stem + '.json')), [Text.Encoding]::UTF8) | ConvertFrom-Json
		$op = [string]$req.op
		$sw = [Diagnostics.Stopwatch]::StartNew()

		if ($op -eq 'ping') {
			$onenote = Get-OneNote
			$xml = Get-HierarchyXml $onenote '' 3
			Write-Response @{ ok = $true; data = @{ ready = $true; hierarchyChars = $xml.Length; elapsedMs = [int]$sw.ElapsedMilliseconds } }
			return
		}

		$onenote = Get-OneNote

		switch ($op) {
			'notebooks' {
				Write-Response @{ ok = $true; data = @{ notebooks = @(Get-NotebookTree $onenote) } }
			}

			'pages' {
				$secId = Resolve-SectionId $onenote ([string]$req.sectionId) ([string]$req.notebook) ([string]$req.section)
				[xml]$doc = Get-HierarchyXml $onenote $secId 4
				$sectionName = ''
				$secNode = $doc.SelectSingleNode('//*[local-name()="Section"]')
				if ($null -ne $secNode) { $sectionName = $secNode.GetAttribute('name') }
				$out = New-Object System.Collections.ArrayList
				# lastModified is deliberately absent: GetHierarchy's copy of it is stale, and reporting it
				# invites a caller to pass it back as expectedModified. read_page is the only source for that.
				foreach ($pg in $doc.SelectNodes('//*[local-name()="Page"]')) {
					[void]$out.Add(@{
							id      = $pg.GetAttribute('ID')
							name    = $pg.GetAttribute('name')
							level   = $pg.GetAttribute('pageLevel')
							created = $pg.GetAttribute('dateTime')
						})
				}
				Write-Response @{ ok = $true; data = @{ sectionId = $secId; section = $sectionName; pages = @($out) } }
			}

			'search' {
				$query = ([string]$req.query).ToLowerInvariant()
				$limit = 25
				if ($null -ne $req.limit) { $limit = [int]$req.limit }
				[xml]$doc = Get-HierarchyXml $onenote '' 4
				$out = New-Object System.Collections.ArrayList
				foreach ($nb in $doc.SelectNodes($XP_NOTEBOOK)) {
					foreach ($sec in $nb.SelectNodes('.//' + $XP_SECTION)) {
						foreach ($pg in $sec.SelectNodes($XP_PAGE)) {
							if ($pg.GetAttribute('name').ToLowerInvariant().Contains($query)) {
								[void]$out.Add(@{
										pageId   = $pg.GetAttribute('ID')
										page     = $pg.GetAttribute('name')
										section  = $sec.GetAttribute('name')
										notebook = $nb.GetAttribute('name')
									})
							}
						}
					}
				}
				$trimmed = @($out | Select-Object -First $limit)
				Write-Response @{ ok = $true; data = @{ total = $out.Count; returned = $trimmed.Count; pages = @($trimmed) } }
			}

			'read_page' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				$model = Read-PageById $onenote $page.id
				if ($req.includeXml -eq $true) {
					$model['rawXml'] = Get-PageContentXml $onenote $page.id
				}
				Write-Response @{ ok = $true; data = @{ page = $model } }
			}

			'create_page' {
				$secId = Resolve-SectionId $onenote ([string]$req.sectionId) ([string]$req.notebook) ([string]$req.section)
				# GetHierarchy fails for an id that is not a node, which validates an explicit sectionId.
				[xml]$secDoc = Get-HierarchyXml $onenote $secId 0
				$sectionName = $secDoc.DocumentElement.GetAttribute('name')
				$title = [string]$req.title
				$paragraphs = @(Get-Paragraphs $req)
				$fingerprint = Get-Sha256Hex ('create:' + $secId)
				Assert-Fingerprint $req $fingerprint
				if (-not (Test-Apply $req)) {
					Write-Response @{ ok = $true; data = @{ applied = $false; fingerprint = $fingerprint; sectionId = $secId; section = $sectionName; elapsedMs = [int]$sw.ElapsedMilliseconds } }
				}
				else {
					$newId = ''
					try { $onenote.CreateNewPage($secId, [ref]$newId, 1) }
					catch { Fail-Write $_ 'create the page' }
					$script:MUTATED = $true
					if (-not [string]::IsNullOrEmpty($title)) {
						Send-PageUpdate $onenote $newId ('<one:Title><one:OE><one:T>' + (New-TextCdata $title) + '</one:T></one:OE></one:Title>')
					}
					if ($paragraphs.Count -gt 0) {
						$pageNode = $null
						# A quick style given by name is resolved against the page's own QuickStyleDef list.
						if (Test-QuickStyleNames $paragraphs) { [xml]$pageDoc = Get-PageContentXml $onenote $newId; $pageNode = $pageDoc.DocumentElement }
						Send-PageUpdate $onenote $newId (New-OutlineXml $paragraphs $null $pageNode)
					}
					Write-Response @{ ok = $true; data = @{ applied = $true; fingerprint = $fingerprint; sectionId = $secId; section = $sectionName; page = (Read-PageById $onenote $newId); elapsedMs = [int]$sw.ElapsedMilliseconds } }
				}
			}

			'append' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				$paragraphs = @(Get-Paragraphs $req)
				if ($paragraphs.Count -eq 0) { Fail 'no lines to append' 'invalid_request' }
				# Appending is additive and cannot clobber a concurrent edit, so the fingerprint only has to
				# tie the apply to the page that was previewed.
				$fingerprint = Get-Sha256Hex ('page:' + $page.id)
				Assert-Fingerprint $req $fingerprint
				$applied = (Test-Apply $req)
				if ($applied) {
					$pageNode = $null
					if (Test-QuickStyleNames $paragraphs) { [xml]$pageDoc = Get-PageContentXml $onenote $page.id; $pageNode = $pageDoc.DocumentElement }
					Send-PageUpdate $onenote $page.id (New-OutlineXml $paragraphs $req.position $pageNode)
				}
				Write-Response @{ ok = $true; data = @{
						applied     = $applied
						fingerprint = $fingerprint
						page        = (Read-PageById $onenote $page.id)
						elapsedMs   = [int]$sw.ElapsedMilliseconds
					} }
			}

			'edit' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				[xml]$doc = Get-PageContentXml $onenote $page.id
				$plan = New-EditPlan $doc.DocumentElement $req $false
				$fingerprint = Get-PlanFingerprint $page.id $plan
				Assert-Fingerprint $req $fingerprint

				$applied = (Test-Apply $req)
				if ($applied) { Send-PageUpdate $onenote $page.id $plan.inner }

				Write-Response @{ ok = $true; data = @{
						applied     = $applied
						fingerprint = $fingerprint
						changes     = $plan.changes
						page        = (Read-PageById $onenote $page.id)
						elapsedMs   = [int]$sw.ElapsedMilliseconds
					} }
			}

			# insert_table and insert_image share a shape: build one one:OE block and either put it into an
			# existing outline (rebuilding that outline in full) or hang it on the page as a new outline.
			'insert_table' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				$block = New-TableOeXml (Get-RequestedLines @{ lines = $req.headers }) $req.rows ([bool]$req.bordersVisible) ([bool]$req.hasHeaderRow) $req.columnWidths
				$outlineId = [string]$req.outline
				$after = [string]$req.after
				if (-not [string]::IsNullOrEmpty($after) -and [string]::IsNullOrEmpty($outlineId)) {
					[xml]$anchorDoc = Get-PageContentXml $onenote $page.id
					$anchorOe = Get-OeByObjectId $anchorDoc.DocumentElement $after
					if ($null -eq $anchorOe) { Fail ('no paragraph with objectID ' + $after + ' on this page') 'not_found' }
					$owner = Get-OwnerOutline $anchorOe
					if ($null -eq $owner) { Fail ('paragraph ' + $after + ' is not inside an outline') 'unsupported' }
					$outlineId = $owner.GetAttribute('objectID')
				}

				$fingerprint = ''
				if ([string]::IsNullOrEmpty($outlineId)) {
					$fingerprint = Get-Sha256Hex ('page:' + $page.id)
					Assert-Fingerprint $req $fingerprint
					if (Test-Apply $req) { Send-PageUpdate $onenote $page.id (New-BlockOutlineXml @($block)) }
				}
				else {
					[xml]$doc = Get-PageContentXml $onenote $page.id
					$plan = New-EditPlan $doc.DocumentElement @{ append = @(@{ objectID = $outlineId; xml = @($block); after = $after }) } $true
					$fingerprint = Get-PlanFingerprint $page.id $plan
					Assert-Fingerprint $req $fingerprint
					if (Test-Apply $req) { Send-PageUpdate $onenote $page.id $plan.inner }
				}

				Write-Response @{ ok = $true; data = @{
						applied     = (Test-Apply $req)
						fingerprint = $fingerprint
						kind        = 'table'
						outlineId   = $outlineId
						page        = (Read-PageById $onenote $page.id)
						elapsedMs   = [int]$sw.ElapsedMilliseconds
					} }
			}

			'insert_image' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				$block = New-ImageOeXml ([string]$req.base64) ([string]$req.format) $req.width $req.height ([string]$req.alt)
				$outlineId = [string]$req.outline
				$after = [string]$req.after
				if (-not [string]::IsNullOrEmpty($after) -and [string]::IsNullOrEmpty($outlineId)) {
					[xml]$anchorDoc = Get-PageContentXml $onenote $page.id
					$anchorOe = Get-OeByObjectId $anchorDoc.DocumentElement $after
					if ($null -eq $anchorOe) { Fail ('no paragraph with objectID ' + $after + ' on this page') 'not_found' }
					$owner = Get-OwnerOutline $anchorOe
					if ($null -eq $owner) { Fail ('paragraph ' + $after + ' is not inside an outline') 'unsupported' }
					$outlineId = $owner.GetAttribute('objectID')
				}

				$fingerprint = ''
				if ([string]::IsNullOrEmpty($outlineId)) {
					$fingerprint = Get-Sha256Hex ('page:' + $page.id)
					Assert-Fingerprint $req $fingerprint
					if (Test-Apply $req) { Send-PageUpdate $onenote $page.id (New-BlockOutlineXml @($block)) }
				}
				else {
					[xml]$doc = Get-PageContentXml $onenote $page.id
					$plan = New-EditPlan $doc.DocumentElement @{ append = @(@{ objectID = $outlineId; xml = @($block); after = $after }) } $true
					$fingerprint = Get-PlanFingerprint $page.id $plan
					Assert-Fingerprint $req $fingerprint
					if (Test-Apply $req) { Send-PageUpdate $onenote $page.id $plan.inner }
				}

				Write-Response @{ ok = $true; data = @{
						applied     = (Test-Apply $req)
						fingerprint = $fingerprint
						kind        = 'image'
						outlineId   = $outlineId
						page        = (Read-PageById $onenote $page.id)
						elapsedMs   = [int]$sw.ElapsedMilliseconds
					} }
			}

			'delete_page' {
				$page = Resolve-Page $onenote ([string]$req.pageId) ([string]$req.notebook) ([string]$req.section) ([string]$req.page)
				# The page name is the guard, not a timestamp: lastModifiedTime does not track content
				# written through COM, so comparing it would refuse correct deletes and miss changed pages.
				$expectName = [string]$req.expectName
				if ([string]::IsNullOrEmpty($expectName)) {
					Fail 'expectName is required: pass the page name you believe this id refers to' 'invalid_request'
				}
				if ($page.name -ne $expectName) {
					Fail ("refusing to delete: expected page '" + $expectName + "' but the id resolves to '" + $page.name + "'") 'name_mismatch'
				}
				$fingerprint = Get-Sha256Hex ('delete:' + $page.id + '|' + $page.name)
				Assert-Fingerprint $req $fingerprint
				$deleted = (Test-Apply $req)
				if ($deleted) {
					try { $onenote.DeleteHierarchy($page.id) }
					catch { Fail-Write $_ 'delete the page' }
					$script:MUTATED = $true
				}
				Write-Response @{ ok = $true; data = @{ applied = $deleted; deleted = $deleted; fingerprint = $fingerprint; pageId = $page.id; name = $page.name; elapsedMs = [int]$sw.ElapsedMilliseconds } }
			}

			default { Fail ('unknown op: ' + $op) 'unknown_op' }
		}
	}
	catch {
		if ($script:RESPONDED) { return }
		$code = 'com_error'
		$msg = $_.Exception.Message
		if ($null -ne $_.Exception.InnerException) { $msg = $msg + ' / ' + $_.Exception.InnerException.Message }
		# The request body is executed through Invoke-Expression, so report the failing source line.
		if ($null -ne $_.InvocationInfo -and -not [string]::IsNullOrEmpty($_.InvocationInfo.Line)) {
			$msg = $msg + ' | at: ' + $_.InvocationInfo.Line.Trim()
		}
		if ($script:MUTATED) {
			$code = 'applied_unverified'
			$msg = 'Part or all of this write was applied; do not retry, read the page again. ' + $msg
		}
		Write-Response @{ ok = $false; error = ($msg -replace "`r?`n", ' '); code = $code }
	}
	finally {
		$script:RESP = $null
	}
}

# --- worker loop ------------------------------------------------------------

$script:WORK = $d
$__dhUtf8 = New-Object System.Text.UTF8Encoding($false)
# The raw stdin stream avoids Console.In, which decodes with the OEM code page.
$__dhIn = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $__dhUtf8, $false)
$__dhOut = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $__dhUtf8)
$__dhOut.AutoFlush = $true

# Pay the JSON module's auto-load now instead of inside the first call.
$null = '{}' | ConvertFrom-Json
$null = @{ a = 1 } | ConvertTo-Json -Compress

$__dhOut.WriteLine('deskhand-ready 1 ' + $PID)
while ($null -ne ($__dhLine = $__dhIn.ReadLine())) {
	$__dhStem = $__dhLine.Trim()
	if ($__dhStem -notmatch '^[0-9A-Za-z-]+$') { continue }
	try {
		# Redirecting every stream keeps host output (warnings and the like) from splicing into a protocol line.
		$null = Invoke-Request $__dhStem *> $null
	}
	catch {
		[Console]::Error.WriteLine('deskhand worker: ' + $_.Exception.Message)
	}
	$Error.Clear()
	$__dhOut.WriteLine('deskhand-done ' + $__dhStem)
}
Reset-OneNote
