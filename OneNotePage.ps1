# Reads and compares OneNote page XML (the 2013 schema GetPageContent returns), without
# OneNote. Used by OneNoteReview.ps1 and by OneNotePage.Tests.ps1.
#
# A page becomes a summary: its title, its page-level Meta tags (hidden record IDs such
# as spm.key), and each outline with its position, size, own Meta tags (spm.block) and
# its content as lines. A line is one paragraph, table row or image, with its indent,
# quick style, colours and tags, so two versions of a page compare line by line.

$ErrorActionPreference = 'Stop'

$script:OneNoteNs = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

function New-OneNoteXml([string]$Xml) {
    $doc = New-Object Xml.XmlDocument
    $doc.XmlResolver = $null   # never fetch a DTD or entity from anywhere
    $doc.LoadXml($Xml)
    $ns = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('one', $script:OneNoteNs)
    [pscustomobject]@{ Doc = $doc; Ns = $ns }
}

# The visible text of a T element: OneNote keeps it as HTML inside CDATA.
function ConvertFrom-OneNoteText([string]$Html) {
    $t = $Html -replace '(?i)<br\s*/?>', ' ' -replace '<[^>]+>', ''
    $t = [Net.WebUtility]::HtmlDecode($t) -replace '\s+', ' '
    $t.Trim()
}

# Colours named in a style attribute or in the span styles of a T element's HTML, as
# lower-case #rrggbb, in order of first use.
function Get-OneNoteColors([string[]]$Styles) {
    $seen = New-Object Collections.Generic.List[string]
    foreach ($s in $Styles) {
        if (-not $s) { continue }
        foreach ($m in [regex]::Matches($s, '(?i)(?:color|background)\s*:\s*(#[0-9a-f]{6}|#[0-9a-f]{3})\b')) {
            $c = $m.Groups[1].Value.ToLower()
            if ($c.Length -eq 4) { $c = '#' + $c[1] + $c[1] + $c[2] + $c[2] + $c[3] + $c[3] }
            if (-not $seen.Contains($c)) { $seen.Add($c) }
        }
    }
    , $seen.ToArray()
}

function Get-OneNoteMeta($Node, $Ns) {
    $m = [ordered]@{}
    foreach ($n in $Node.SelectNodes('one:Meta', $Ns)) { $m[$n.GetAttribute('name')] = $n.GetAttribute('content') }
    $m
}

function Format-OneNoteLine([int]$Depth, [string]$Text, [string]$Style, [string[]]$Colors, [string[]]$Tags) {
    $s = ('  ' * $Depth) + $Text
    if ($Style -and $Style -ne 'p') { $s += " [$Style]" }
    if ($Colors -and $Colors.Count) { $s += ' (' + ($Colors -join ' ') + ')' }
    if ($Tags -and $Tags.Count) { $s += ' {' + ($Tags -join ', ') + '}' }
    $s
}

# Lines for the OE elements under an OEChildren node, depth first.
function Get-OneNoteLines($Children, $Ns, [hashtable]$Styles, [hashtable]$TagNames, [int]$Depth) {
    $out = New-Object Collections.Generic.List[string]
    if (-not $Children) { return }
    foreach ($oe in $Children.SelectNodes('one:OE', $Ns)) {
        $style = $Styles[$oe.GetAttribute('quickStyleIndex')]
        $tags = @(foreach ($t in $oe.SelectNodes('one:Tag', $Ns)) {
            $n = $TagNames[$t.GetAttribute('index')]; if (-not $n) { $n = 'tag ' + $t.GetAttribute('index') }
            if ($t.GetAttribute('completed') -eq 'true') { "$n done" } else { $n }
        })
        $tNode = $oe.SelectSingleNode('one:T', $Ns)
        $table = $oe.SelectSingleNode('one:Table', $Ns)
        $image = $oe.SelectSingleNode('one:Image', $Ns)
        if ($tNode) {
            $colors = Get-OneNoteColors @($oe.GetAttribute('style'), $tNode.GetAttribute('style'), $tNode.InnerText)
            $out.Add((Format-OneNoteLine $Depth (ConvertFrom-OneNoteText $tNode.InnerText) $style $colors $tags))
        }
        if ($table) {
            $widths = @(foreach ($c in $table.SelectNodes('one:Columns/one:Column', $Ns)) { [string][Math]::Round([double]$c.GetAttribute('width')) })
            $out.Add((('  ' * $Depth) + "[table $($widths.Count) columns: $($widths -join ' ')]"))
            foreach ($row in $table.SelectNodes('one:Row', $Ns)) {
                $cells = @(); $shades = @()
                foreach ($cell in $row.SelectNodes('one:Cell', $Ns)) {
                    $texts = @(Get-OneNoteLines $cell.SelectSingleNode('one:OEChildren', $Ns) $Ns $Styles $TagNames 0 | ForEach-Object { $_.Trim() })
                    $cells += ($texts -join ' / ')
                    $sh = $cell.GetAttribute('shadingColor')
                    if ($sh) { $shades += $sh.ToLower() }
                }
                $line = ('  ' * $Depth) + '| ' + ($cells -join ' | ') + ' |'
                if ($shades.Count) { $line += ' (shading ' + (($shades | Select-Object -Unique) -join ' ') + ')' }
                $out.Add($line)
            }
        }
        if ($image) {
            $sz = $image.SelectSingleNode('one:Size', $Ns)
            $dim = if ($sz) { "$([Math]::Round([double]$sz.GetAttribute('width')))x$([Math]::Round([double]$sz.GetAttribute('height')))" } else { '?' }
            $alt = $image.GetAttribute('alt')
            $out.Add((('  ' * $Depth) + "[image $dim$(if ($alt) { " '$alt'" })]"))
        }
        foreach ($l in (Get-OneNoteLines $oe.SelectSingleNode('one:OEChildren', $Ns) $Ns $Styles $TagNames ($Depth + 1))) { $out.Add($l) }
    }
    $out.ToArray()
}

# The summary of one page's XML (GetPageContent, piBasic, xs2013).
function ConvertFrom-OneNotePage([string]$Xml) {
    $x = New-OneNoteXml $Xml
    $ns = $x.Ns
    $page = $x.Doc.SelectSingleNode('/one:Page', $ns)
    if (-not $page) { throw 'Not OneNote page XML (no one:Page element in the 2013 namespace).' }
    $styles = @{}; $styleDefs = @()
    foreach ($q in $page.SelectNodes('one:QuickStyleDef', $ns)) {
        $styles[$q.GetAttribute('index')] = $q.GetAttribute('name')
        $styleDefs += [pscustomobject][ordered]@{
            Name = $q.GetAttribute('name'); Font = $q.GetAttribute('font'); Size = $q.GetAttribute('fontSize')
            Color = $q.GetAttribute('fontColor').ToLower(); Highlight = $q.GetAttribute('highlightColor').ToLower()
        }
    }
    $tagNames = @{}
    foreach ($t in $page.SelectNodes('one:TagDef', $ns)) { $tagNames[$t.GetAttribute('index')] = $t.GetAttribute('name') }
    $titleNode = $page.SelectSingleNode('one:Title/one:OE/one:T', $ns)
    $title = if ($titleNode) { ConvertFrom-OneNoteText $titleNode.InnerText } else { $page.GetAttribute('name') }
    $i = 0
    $outlines = foreach ($o in $page.SelectNodes('one:Outline', $ns)) {
        $i++
        $pos = $o.SelectSingleNode('one:Position', $ns); $size = $o.SelectSingleNode('one:Size', $ns)
        $meta = Get-OneNoteMeta $o $ns
        [pscustomobject][ordered]@{
            Index = $i; ObjectId = $o.GetAttribute('objectID'); Meta = $meta
            X = $(if ($pos) { [Math]::Round([double]$pos.GetAttribute('x'), 1) }); Y = $(if ($pos) { [Math]::Round([double]$pos.GetAttribute('y'), 1) })
            Width = $(if ($size) { [Math]::Round([double]$size.GetAttribute('width'), 1) }); Height = $(if ($size) { [Math]::Round([double]$size.GetAttribute('height'), 1) })
            Lines = @(Get-OneNoteLines $o.SelectSingleNode('one:OEChildren', $ns) $ns $styles $tagNames 0)
        }
    }
    $images = @($page.SelectNodes('one:Image', $ns)).Count   # floating images, outside outlines
    $settings = $page.SelectSingleNode('one:PageSettings', $ns)
    [pscustomobject][ordered]@{
        Id = $page.GetAttribute('ID'); Title = $title; Meta = (Get-OneNoteMeta $page $ns)
        Level = $page.GetAttribute('pageLevel'); PageColor = $(if ($settings) { $settings.GetAttribute('color') })
        Styles = @($styleDefs); Outlines = @($outlines); FloatingImages = $images
    }
}

# Outlines matched across two versions by a Meta tag they carry (spm.block), else by order.
function Get-OneNoteOutlineKey($Outline) {
    if ($Outline.Meta.Count) { $k = @($Outline.Meta.Keys)[0]; return "outline $k=$($Outline.Meta[$k])" }
    "outline $($Outline.Index)"
}

# Line differences as '- old' and '+ new', from a longest common subsequence.
function Compare-OneNoteLines([string[]]$Before, [string[]]$After) {
    $n = $Before.Count; $m = $After.Count
    $L = New-Object 'int[,]' ($n + 1), ($m + 1)
    for ($i = $n - 1; $i -ge 0; $i--) {
        for ($j = $m - 1; $j -ge 0; $j--) {
            # Windows PowerShell 5.1 can't parse a 2-D index with expressions inside a
            # method call's arguments, so the neighbours are read into variables first.
            $i1 = $i + 1; $j1 = $j + 1
            if ($Before[$i] -ceq $After[$j]) { $L[$i, $j] = $L[$i1, $j1] + 1 }
            else { $down = $L[$i1, $j]; $right = $L[$i, $j1]; $L[$i, $j] = [Math]::Max($down, $right) }
        }
    }
    $out = New-Object Collections.Generic.List[string]
    $i = 0; $j = 0
    while ($i -lt $n -and $j -lt $m) {
        if ($Before[$i] -ceq $After[$j]) { $i++; $j++ }
        else {
            $i1 = $i + 1; $j1 = $j + 1; $down = $L[$i1, $j]; $right = $L[$i, $j1]
            if ($down -ge $right) { $out.Add("- $($Before[$i])"); $i++ } else { $out.Add("+ $($After[$j])"); $j++ }
        }
    }
    while ($i -lt $n) { $out.Add("- $($Before[$i])"); $i++ }
    while ($j -lt $m) { $out.Add("+ $($After[$j])"); $j++ }
    $out.ToArray()
}

function Compare-OneNoteMeta([string]$Where, $Before, $After) {
    foreach ($k in $Before.Keys) {
        if (-not $After.Contains($k)) { "$Where meta $k removed (was '$($Before[$k])')" }
        elseif ($After[$k] -cne $Before[$k]) { "$Where meta $k '$($Before[$k])' -> '$($After[$k])'" }
    }
    foreach ($k in $After.Keys) { if (-not $Before.Contains($k)) { "$Where meta $k added '$($After[$k])'" } }
}

# Every difference between two summaries of the same page, as text lines.
function Compare-OneNotePage($Before, $After) {
    $changes = New-Object Collections.Generic.List[string]
    if ($Before.Title -cne $After.Title) { $changes.Add("title '$($Before.Title)' -> '$($After.Title)'") }
    foreach ($c in (Compare-OneNoteMeta 'page' $Before.Meta $After.Meta)) { $changes.Add($c) }
    if ($Before.PageColor -ne $After.PageColor) { $changes.Add("page colour '$($Before.PageColor)' -> '$($After.PageColor)'") }
    $bs = @{}; foreach ($s in $Before.Styles) { $bs[$s.Name] = "$($s.Font) $($s.Size) $($s.Color) $($s.Highlight)" }
    foreach ($s in $After.Styles) {
        $v = "$($s.Font) $($s.Size) $($s.Color) $($s.Highlight)"
        if ($bs.ContainsKey($s.Name) -and $bs[$s.Name] -ne $v) { $changes.Add("style $($s.Name): $($bs[$s.Name]) -> $v") }
    }
    $bo = [ordered]@{}; foreach ($o in $Before.Outlines) { $bo[(Get-OneNoteOutlineKey $o)] = $o }
    $ao = [ordered]@{}; foreach ($o in $After.Outlines) { $ao[(Get-OneNoteOutlineKey $o)] = $o }
    foreach ($k in $bo.Keys) { if (-not $ao.Contains($k)) { $changes.Add("$k removed ($($bo[$k].Lines.Count) lines)") } }
    foreach ($k in $ao.Keys) {
        $a = $ao[$k]
        if (-not $bo.Contains($k)) {
            $changes.Add("$k added at $($a.X),$($a.Y) size $($a.Width)x$($a.Height)")
            foreach ($l in $a.Lines) { $changes.Add("  + $l") }
            continue
        }
        $b = $bo[$k]
        if ($b.X -ne $a.X -or $b.Y -ne $a.Y) { $changes.Add("$k moved $($b.X),$($b.Y) -> $($a.X),$($a.Y)") }
        if ($b.Width -ne $a.Width -or $b.Height -ne $a.Height) { $changes.Add("$k size $($b.Width)x$($b.Height) -> $($a.Width)x$($a.Height)") }
        foreach ($c in (Compare-OneNoteMeta $k $b.Meta $a.Meta)) { $changes.Add($c) }
        $diff = @(Compare-OneNoteLines ([string[]]$b.Lines) ([string[]]$a.Lines))
        if ($diff.Count) { $changes.Add("$k text:"); foreach ($d in $diff) { $changes.Add("  $d") } }
    }
    if ($Before.FloatingImages -ne $After.FloatingImages) { $changes.Add("floating images $($Before.FloatingImages) -> $($After.FloatingImages)") }
    $changes.ToArray()
}

# Record-ID check over pages: each item has Key (the meta value, or $null), Title and
# Section. Reports pages without the meta tag, keys on more than one page, and, given
# the keys that should exist, which are missing and which are not expected.
function Test-OneNoteMetaKeys([object[]]$Pages, [string]$MetaName, [string[]]$Expected) {
    $missing = @($Pages | Where-Object { -not $_.Key } | ForEach-Object { "$($_.Section) / $($_.Title)" })
    $groups = @($Pages | Where-Object { $_.Key } | Group-Object Key -CaseSensitive | Where-Object { $_.Count -gt 1 })
    $dupes = @($groups | ForEach-Object { "$($_.Name): " + (($_.Group | ForEach-Object { "$($_.Section) / $($_.Title)" }) -join '; ') })
    $keys = @($Pages | Where-Object { $_.Key } | ForEach-Object { $_.Key })
    $notFound = @(); $unexpected = @()
    if ($Expected) {
        $notFound = @($Expected | Where-Object { $keys -cnotcontains $_ })
        $unexpected = @($keys | Where-Object { $Expected -cnotcontains $_ } | Select-Object -Unique)
    }
    [pscustomobject][ordered]@{
        MetaName = $MetaName; Pages = $Pages.Count; WithKey = $keys.Count
        WithoutKey = $missing; Duplicates = $dupes; ExpectedNotFound = $notFound; NotExpected = $unexpected
        Ok = (-not $dupes.Count -and -not $notFound.Count -and -not $unexpected.Count)
    }
}

# A snapshot is a folder of page XML files plus index.json ([{Key, Title, Section, File}]).
# Pages are matched across snapshots by Key: the record-ID meta when the page has one,
# else "<section>/<title>".
function Compare-OneNoteSnapshotFolder([string]$BeforeDir, [string]$AfterDir) {
    $read = {
        param($dir)
        $idx = [IO.File]::ReadAllText((Join-Path $dir 'index.json')) | ConvertFrom-Json
        $map = [ordered]@{}
        foreach ($e in @($idx)) { $map[[string]$e.Key] = $e }
        $map
    }
    $b = & $read $BeforeDir; $a = & $read $AfterDir
    $added = @(); $removed = @(); $changed = [ordered]@{}; $same = 0
    foreach ($k in $b.Keys) { if (-not $a.Contains($k)) { $removed += "$k ($($b[$k].Section) / $($b[$k].Title))" } }
    foreach ($k in $a.Keys) {
        if (-not $b.Contains($k)) { $added += "$k ($($a[$k].Section) / $($a[$k].Title))"; continue }
        $pb = ConvertFrom-OneNotePage ([IO.File]::ReadAllText((Join-Path $BeforeDir $b[$k].File)))
        $pa = ConvertFrom-OneNotePage ([IO.File]::ReadAllText((Join-Path $AfterDir $a[$k].File)))
        $c = @(Compare-OneNotePage $pb $pa)
        if ($b[$k].Section -ne $a[$k].Section) { $c = @("section '$($b[$k].Section)' -> '$($a[$k].Section)'") + $c }
        if ($c.Count) { $changed[$k] = $c } else { $same++ }
    }
    [pscustomobject][ordered]@{ PagesAdded = $added; PagesRemoved = $removed; PagesChanged = $changed; PagesSame = $same }
}

function Format-OneNoteSnapshotReport($Comparison, [string]$Title) {
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add($Title)
    $lines.Add("$(@($Comparison.PagesChanged.Keys).Count) changed, $($Comparison.PagesAdded.Count) added, $($Comparison.PagesRemoved.Count) removed, $($Comparison.PagesSame) unchanged")
    foreach ($p in $Comparison.PagesAdded) { $lines.Add("+ page $p") }
    foreach ($p in $Comparison.PagesRemoved) { $lines.Add("- page $p") }
    foreach ($k in $Comparison.PagesChanged.Keys) {
        $lines.Add("~ page $k")
        foreach ($c in $Comparison.PagesChanged[$k]) { $lines.Add("    $c") }
    }
    ($lines -join "`r`n")
}
