# Creates the probe section in an existing notebook if it is missing, and prints its id.
# The caller prepends: $Notebook = '<name>'; $Section = '<name>'
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject OneNote.Application
$xml = ''
$app.GetHierarchy('', 3, [ref]$xml)
$doc = [xml]$xml
$ns = New-Object Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('one', $doc.DocumentElement.NamespaceURI)
$book = $doc.SelectNodes('//one:Notebook', $ns) | Where-Object { $_.name -eq $Notebook } | Select-Object -First 1
if (-not $book) { throw "Notebook not found: $Notebook" }
$existing = $book.SelectNodes('.//one:Section', $ns) | Where-Object { $_.name -eq $Section } | Select-Object -First 1
if ($existing) { Write-Output "exists $($existing.ID)"; exit 0 }
$newId = ''
# CreateFileType 3 = cftSection
$app.OpenHierarchy("$Section.one", $book.ID, [ref]$newId, 3)
Write-Output "created $newId"
