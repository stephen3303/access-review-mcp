# Checks OneNotePage.ps1 against small OneNote page XML, without OneNote:
#   powershell -NoProfile -File tools\review-mcp\OneNotePage.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OneNotePage.ps1')

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}

function PageXml([string]$Title, [string]$Key, [string]$Body, [string]$X = '36', [string]$Color = '#1F3A5F', [string]$Extra = '') {
    @"
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="{P1}" name="$Title" pageLevel="1">
  <one:TagDef index="0" type="0" symbol="3" name="To Do"/>
  <one:QuickStyleDef index="0" name="PageTitle" fontColor="automatic" highlightColor="automatic" font="Calibri Light" fontSize="20.0"/>
  <one:QuickStyleDef index="1" name="p" fontColor="automatic" highlightColor="automatic" font="Calibri" fontSize="11.0"/>
  <one:QuickStyleDef index="2" name="h2" fontColor="$Color" highlightColor="automatic" font="Calibri" fontSize="14.0"/>
  <one:PageSettings RTL="false" color="automatic"/>
  <one:Meta name="spm.key" content="$Key"/>
  <one:Title><one:OE quickStyleIndex="0"><one:T><![CDATA[$Title]]></one:T></one:OE></one:Title>
  <one:Outline objectID="{O1}">
    <one:Position x="$X" y="86" z="0"/>
    <one:Size width="720" height="40" isSetByUser="true"/>
    <one:Meta name="spm.block" content="kit"/>
    <one:OEChildren>$Body</one:OEChildren>
  </one:Outline>$Extra
</one:Page>
"@
}

$body = @'
<one:OE quickStyleIndex="2"><one:T><![CDATA[Programs]]></one:T></one:OE>
<one:OE quickStyleIndex="1"><one:T><![CDATA[<span style='color:#C00000'>3 overdue</span> &amp; 2 open]]></one:T>
  <one:OEChildren><one:OE quickStyleIndex="1"><one:Tag index="0" completed="true"/><one:T><![CDATA[Close POA&amp;M 7]]></one:T></one:OE></one:OEChildren>
</one:OE>
<one:OE><one:Table bordersVisible="true"><one:Columns><one:Column index="0" width="120.5"/><one:Column index="1" width="300"/></one:Columns>
  <one:Row><one:Cell shadingColor="#E8EEF5"><one:OEChildren><one:OE><one:T><![CDATA[Program]]></one:T></one:OE></one:OEChildren></one:Cell>
           <one:Cell shadingColor="#E8EEF5"><one:OEChildren><one:OE><one:T><![CDATA[Manager]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
  <one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Fire safety]]></one:T></one:OE></one:OEChildren></one:Cell>
           <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Vacant]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
</one:Table></one:OE>
<one:OE><one:Image alt="logo"><one:Size width="64" height="32"/></one:Image></one:OE>
'@

# --- reading a page ---
$s = ConvertFrom-OneNotePage (PageXml 'Command status' 'command' $body)
Check 'title' ($s.Title -eq 'Command status')
Check 'page meta' ($s.Meta['spm.key'] -eq 'command')
Check 'one outline with its block meta' ($s.Outlines.Count -eq 1 -and $s.Outlines[0].Meta['spm.block'] -eq 'kit')
Check 'outline position and size' ($s.Outlines[0].X -eq 36 -and $s.Outlines[0].Y -eq 86 -and $s.Outlines[0].Width -eq 720)
$lines = $s.Outlines[0].Lines
Check 'heading with style' ($lines[0] -eq 'Programs [h2]')
Check 'span colour lower-cased, entity decoded' ($lines[1] -eq '3 overdue & 2 open (#c00000)')
Check 'nested line indented with done tag' ($lines[2] -eq '  Close POA&M 7 {To Do done}')
Check 'table columns' ($lines[3] -eq '[table 2 columns: 120 300]')
Check 'table header row with shading' ($lines[4] -eq '| Program | Manager | (shading #e8eef5)')
Check 'table body row' ($lines[5] -eq '| Fire safety | Vacant |')
Check 'image with size and alt' ($lines[6] -eq "[image 64x32 'logo']")
Check 'quick style colour' ((@($s.Styles | Where-Object Name -eq 'h2')[0].Color) -eq '#1f3a5f')
Check 'short colour expanded' ((Get-OneNoteColors @('color:#abc')) -join ',' -eq '#aabbcc')

$bad = $false
try { $null = ConvertFrom-OneNotePage '<root/>' } catch { $bad = $_.Exception.Message -match 'Not OneNote page XML' }
Check 'non-page XML refused' $bad

# --- comparing two versions ---
$body2 = $body -replace '3 overdue', '5 overdue' -replace 'Vacant', 'LT Smith'
$after = ConvertFrom-OneNotePage (PageXml 'Command status' 'command' $body2 '40' '#0F4C81')
$c = @(Compare-OneNotePage $s $after)
Check 'outline move reported' (@($c | Where-Object { $_ -eq 'outline spm.block=kit moved 36,86 -> 40,86' }).Count -eq 1)
Check 'style colour change reported' (@($c | Where-Object { $_ -like 'style h2:*#1f3a5f*->*#0f4c81*' }).Count -eq 1)
Check 'changed line as - and +' ((@($c) -contains '  - 3 overdue & 2 open (#c00000)') -and (@($c) -contains '  + 5 overdue & 2 open (#c00000)'))
Check 'changed table row' (@($c) -contains '  + | Fire safety | LT Smith |')
Check 'unchanged lines not reported' (-not (@($c) -match 'Programs \[h2\]'))
$same = @(Compare-OneNotePage $s (ConvertFrom-OneNotePage (PageXml 'Command status' 'command' $body)))
Check 'identical pages: no changes' ($same.Count -eq 0)

$extra = '<one:Outline objectID="{O2}"><one:Position x="400" y="300" z="1"/><one:Size width="200" height="20"/><one:OEChildren><one:OE><one:T><![CDATA[My own note]]></one:T></one:OE></one:OEChildren></one:Outline>'
$withNote = ConvertFrom-OneNotePage (PageXml 'Command status' 'command' $body '36' '#1F3A5F' $extra)
$c = @(Compare-OneNotePage $s $withNote)
Check 'added unkeyed outline reported with its lines' (($c -contains 'outline 2 added at 400,300 size 200x20') -and ($c -contains '  + My own note'))
$c = @(Compare-OneNotePage $s (ConvertFrom-OneNotePage (PageXml 'Command status' 'cmd' $body)))
Check 'record ID change reported' ($c -contains "page meta spm.key 'command' -> 'cmd'")

# --- record IDs ---
$pages = @(
    [pscustomobject]@{ Key = 'command'; Title = 'Command status'; Section = 'PM / Binder' }
    [pscustomobject]@{ Key = 'program:1'; Title = 'Fire safety'; Section = 'PM / Programs' }
    [pscustomobject]@{ Key = 'program:1'; Title = 'Fire safety (copy)'; Section = 'PM / Programs' }
    [pscustomobject]@{ Key = $null; Title = 'Notes'; Section = 'PM / Binder' }
)
$r = Test-OneNoteMetaKeys $pages 'spm.key' @('command', 'people', 'program:1')
Check 'page without ID listed' (($r.WithoutKey -join ',') -eq 'PM / Binder / Notes')
Check 'duplicate ID listed with both pages' ($r.Duplicates.Count -eq 1 -and $r.Duplicates[0] -like 'program:1: *Fire safety; *Fire safety (copy)')
Check 'expected ID not found' (($r.ExpectedNotFound -join ',') -eq 'people')
Check 'not ok' (-not $r.Ok)
$r = Test-OneNoteMetaKeys @($pages[0], $pages[1]) 'spm.key' @('command', 'program:1')
Check 'clean set ok' ($r.Ok -and $r.WithKey -eq 2)
$r = Test-OneNoteMetaKeys @($pages[0], $pages[1]) 'spm.key' @('command')
Check 'unexpected ID listed' (($r.NotExpected -join ',') -eq 'program:1' -and -not $r.Ok)

# --- snapshot folders ---
$root = Join-Path ([IO.Path]::GetTempPath()) ('onenote-tests-' + [guid]::NewGuid())
$b = Join-Path $root 'before'; $a = Join-Path $root 'after'
New-Item -ItemType Directory -Force -Path $b, $a | Out-Null
try {
    $utf8 = New-Object Text.UTF8Encoding($false)
    $write = {
        param($dir, $entries)
        foreach ($e in $entries) { [IO.File]::WriteAllText((Join-Path $dir $e.File), $e.Xml, $utf8) }
        $idx = @($entries | ForEach-Object { [pscustomobject][ordered]@{ Key = $_.Key; Title = $_.Title; Section = $_.Section; File = $_.File } })
        [IO.File]::WriteAllText((Join-Path $dir 'index.json'), (ConvertTo-Json -InputObject $idx -Depth 3), $utf8)
    }
    & $write $b @(
        @{ Key = 'command'; Title = 'Command status'; Section = 'Binder'; File = '001-command.xml'; Xml = (PageXml 'Command status' 'command' $body) }
        @{ Key = 'people'; Title = 'Gains and losses'; Section = 'Binder'; File = '002-people.xml'; Xml = (PageXml 'Gains and losses' 'people' $body) }
        @{ Key = 'program:1'; Title = 'Fire safety'; Section = 'Binder'; File = '003-program-1.xml'; Xml = (PageXml 'Fire safety' 'program:1' $body) })
    & $write $a @(
        @{ Key = 'command'; Title = 'Command status'; Section = 'Binder'; File = '001-command.xml'; Xml = (PageXml 'Command status' 'command' $body2) }
        @{ Key = 'people'; Title = 'Gains and losses'; Section = 'Binder'; File = '002-people.xml'; Xml = (PageXml 'Gains and losses' 'people' $body) }
        @{ Key = 'program:2'; Title = 'Damage control'; Section = 'Binder'; File = '003-program-2.xml'; Xml = (PageXml 'Damage control' 'program:2' $body) })
    # A single-page index is written by ConvertTo-Json as one object, not an array.
    $cmp = Compare-OneNoteSnapshotFolder $b $a
    Check 'snapshot: added page' (($cmp.PagesAdded -join ',') -eq 'program:2 (Binder / Damage control)')
    Check 'snapshot: removed page' (($cmp.PagesRemoved -join ',') -eq 'program:1 (Binder / Fire safety)')
    Check 'snapshot: changed page' ((@($cmp.PagesChanged.Keys) -join ',') -eq 'command' -and $cmp.PagesSame -eq 1)
    $text = Format-OneNoteSnapshotReport $cmp 'OneNote snapshots: before -> after'
    Check 'report summary line' (($text -split "`r`n")[1] -eq '1 changed, 1 added, 1 removed, 1 unchanged')
    Check 'report shows line change' ($text -match '(?m)^      \+ 5 overdue')

    $one = Join-Path $root 'one'; New-Item -ItemType Directory -Force -Path $one | Out-Null
    & $write $one @(@{ Key = 'people'; Title = 'Gains and losses'; Section = 'Binder'; File = '001-people.xml'; Xml = (PageXml 'Gains and losses' 'people' $body) })
    $cmp = Compare-OneNoteSnapshotFolder $one $one
    Check 'single-page snapshot reads' ($cmp.PagesSame -eq 1 -and -not $cmp.PagesAdded.Count)
} finally { Remove-Item -Recurse -Force $root }

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
