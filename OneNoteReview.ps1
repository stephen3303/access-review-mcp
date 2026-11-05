# OneNote review helpers: read, compare and capture pages of local OneNote notebooks
# through the desktop OneNote COM API (OneNote.Application), the same API an Access
# app uses to publish pages (SPM's Publish to OneNote). No app registration, no Graph,
# no sign-in: everything stays on this PC.
#
# Dot-sourced by AccessReviewMcp.ps1 after AccessReview.ps1 (it uses KitReviewWin and
# the output, full-path and allowed-folder helpers there), and usable directly:
#
#   . "$PSScriptRoot\AccessReview.ps1"; . "$PSScriptRoot\OneNoteReview.ps1"
#   Open-ReviewNotebook -Path 'C:\Review\Notebooks\Binder'
#   Get-ReviewNotebookPages -Notebook Binder -MetaName spm.key
#   Get-ReviewOneNotePage -Page command -MetaName spm.key
#   Save-ReviewOneNoteCapture -Page command -MetaName spm.key -Label navy
#   Save-ReviewOneNoteSnapshot -Label before -Notebook Binder -MetaName spm.key
#   Compare-ReviewOneNoteSnapshot -Before before -After after
#
# Boundary: only notebooks stored on this PC are touched. A notebook whose path is a
# web address (OneDrive, SharePoint, a work tenant) is never listed, read or captured,
# and when ACCESS_REVIEW_ALLOWED_FOLDERS is set a local notebook must be inside one of
# those folders too. Nothing here edits a page.
#
# OneNote is one shared process per user: this attaches to the OneNote already running
# (or starts it) and never quits it. Capture opens the page in a new OneNote window,
# parks that window past the desktop edge, draws it with PrintWindow and closes it.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'OneNotePage.ps1')

$script:OneNote = $null
$script:OneNoteOut = $null
$script:HsSelf = 0; $script:HsNotebooks = 2; $script:HsSections = 3; $script:HsPages = 4
$script:Xs2013 = 2

function Get-ReviewOneNote {
    if ($script:OneNote) {
        try { $null = $script:OneNote.Windows.Count; return $script:OneNote } catch { $script:OneNote = $null }
    }
    try { $script:OneNote = New-Object -ComObject OneNote.Application }
    catch { throw "OneNote desktop is not available here ($($_.Exception.Message)). It needs the OneNote desktop app (not OneNote for Windows 10), run at least once by this user." }
    $script:OneNote
}

function Get-ReviewOneNoteOut {
    if (-not $script:OneNoteOut) {
        $script:OneNoteOut = Join-Path (Get-ReviewOutputBase) ('onenote-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType Directory -Force -Path $script:OneNoteOut | Out-Null
    }
    $script:OneNoteOut
}

function Get-ReviewOneNoteXml([string]$Id, [int]$Scope) {
    $on = Get-ReviewOneNote
    [string]$xml = ''
    $on.GetHierarchy($Id, $Scope, [ref]$xml, $script:Xs2013)
    New-OneNoteXml $xml
}

function Get-ReviewOneNotePageXml([string]$PageId) {
    $on = Get-ReviewOneNote
    [string]$xml = ''
    $on.GetPageContent($PageId, [ref]$xml, 0, $script:Xs2013)   # piBasic: no image data
    $xml
}

# Why a notebook is out of bounds, or $null when this tool may use it.
function Get-ReviewNotebookRefusal([string]$Path) {
    if (-not $Path -or $Path -match '^[a-z][a-z0-9+.-]*://' -or -not (Test-FullPath $Path)) {
        return 'it is not stored on this PC (cloud and work notebooks are never read)'
    }
    $allowed = @(((Get-ReviewSetting 'ACCESS_REVIEW_ALLOWED_FOLDERS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($allowed.Count -and -not @($allowed | Where-Object { (Test-FullPath $_) -and (Test-PathInFolder ($Path.TrimEnd('\') + '\') $_) }).Count) {
        return "it is outside ACCESS_REVIEW_ALLOWED_FOLDERS ($($allowed -join '; '))"
    }
    $null
}

# Open notebooks this tool may use, plus how many others are open and left alone.
function Get-ReviewNotebooks {
    $x = Get-ReviewOneNoteXml '' $script:HsNotebooks
    $ok = @(); $other = 0
    foreach ($n in $x.Doc.SelectNodes('/one:Notebooks/one:Notebook', $x.Ns)) {
        if (Get-ReviewNotebookRefusal $n.GetAttribute('path')) { $other++; continue }
        $ok += [pscustomobject][ordered]@{ Name = $n.GetAttribute('nickname'); Id = $n.GetAttribute('ID'); Path = $n.GetAttribute('path') }
    }
    [pscustomobject][ordered]@{ Notebooks = $ok; OtherNotebooksNotShown = $other }
}

# Opens a notebook stored in a local folder (OneNote adds it to its notebook list).
function Open-ReviewNotebook {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-FullPath $Path)) { throw "'$Path' is not a full path to a local notebook folder or .onetoc2 file." }
    $Path = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "'$Path' does not exist. This tool opens existing notebooks only." }
    $why = Get-ReviewNotebookRefusal $Path
    if ($why) { throw "Not opening '$Path': $why." }
    Assert-ReviewFileAllowed $Path   # also refuses a junction or symbolic link out of the allowed folder
    $on = Get-ReviewOneNote
    [string]$id = ''
    $on.OpenHierarchy($Path, '', [ref]$id, 0)   # cftNone: open, never create
    $nb = Get-ReviewNotebooks
    [pscustomobject][ordered]@{ Opened = $Path; Id = $id; Notebooks = $nb.Notebooks; OtherNotebooksNotShown = $nb.OtherNotebooksNotShown }
}

# The allowed notebook named (nickname or ID), or every allowed notebook when Name is empty.
function Resolve-ReviewNotebooks([string]$Name) {
    $all = @((Get-ReviewNotebooks).Notebooks)
    if (-not $Name) { return $all }
    $hit = @($all | Where-Object { $_.Name -eq $Name -or $_.Id -eq $Name })
    if (-not $hit.Count) { throw "No open local notebook '$Name'. Open notebooks this tool may use: $((@($all | ForEach-Object Name)) -join ', ')$(if (-not $all.Count) { '(none; use onenote_open_notebook)' })." }
    $hit
}

# Every page of the allowed notebooks (or one notebook, or one section), skipping the
# recycle bin. With MetaName, each page's value of that Meta tag (its record ID) as Key.
function Get-ReviewOneNotePageList([string]$Notebook, [string]$Section, [string]$MetaName) {
    foreach ($nb in (Resolve-ReviewNotebooks $Notebook)) {
        $x = Get-ReviewOneNoteXml $nb.Id $script:HsPages
        foreach ($s in $x.Doc.SelectNodes('//one:Section', $x.Ns)) {
            if ($s.SelectSingleNode('ancestor::one:SectionGroup[@isRecycleBin="true"]', $x.Ns) -or $s.GetAttribute('isInRecycleBin') -eq 'true') { continue }
            $groups = @($s.SelectNodes('ancestor::one:SectionGroup', $x.Ns) | ForEach-Object { $_.GetAttribute('name') })
            $secPath = (@($nb.Name) + $groups + @($s.GetAttribute('name'))) -join ' / '
            if ($Section -and $s.GetAttribute('name') -ne $Section -and $secPath -ne $Section) { continue }
            foreach ($p in $s.SelectNodes('one:Page', $x.Ns)) {
                if ($p.GetAttribute('isInRecycleBin') -eq 'true') { continue }
                $key = $null
                if ($MetaName) {
                    $px = New-OneNoteXml (Get-ReviewOneNotePageXml $p.GetAttribute('ID'))
                    $m = $px.Doc.SelectSingleNode("/one:Page/one:Meta[@name='$($MetaName -replace "'", '')']", $px.Ns)
                    if ($m) { $key = $m.GetAttribute('content') }
                }
                [pscustomobject][ordered]@{
                    Section = $secPath; Title = $p.GetAttribute('name'); Id = $p.GetAttribute('ID'); Key = $key
                    Level = $p.GetAttribute('pageLevel'); LastModified = $p.GetAttribute('lastModifiedTime')
                }
            }
        }
    }
}

function Get-ReviewNotebookPages {
    param([string]$Notebook, [string]$Section, [string]$MetaName)
    $pages = @(Get-ReviewOneNotePageList $Notebook $Section $MetaName)
    [pscustomobject][ordered]@{ Pages = $pages; Count = $pages.Count }
}

# A page by its ID, its record ID (the MetaName meta value) or its title, in the allowed
# notebooks only. A page ID from somewhere else is refused, not read.
function Resolve-ReviewOneNotePage([string]$Page, [string]$Notebook, [string]$MetaName) {
    $list = @(Get-ReviewOneNotePageList $Notebook '' $MetaName)
    $hit = @($list | Where-Object { $_.Id -eq $Page })
    if (-not $hit.Count -and $MetaName) { $hit = @($list | Where-Object { $_.Key -ceq $Page }) }
    if (-not $hit.Count) { $hit = @($list | Where-Object { $_.Title -eq $Page }) }
    if (-not $hit.Count) { throw "No page '$Page' in the open local notebooks$(if ($MetaName) { " (by ID, $MetaName or title)" } else { ' (by ID or title; pass meta_name to find it by record ID)' })." }
    if ($hit.Count -gt 1) { throw "'$Page' matches $($hit.Count) pages: $((@($hit | ForEach-Object { "$($_.Section) / $($_.Title) [$($_.Id)]" })) -join '; '). Use the page ID." }
    $hit[0]
}

function Get-ReviewSafeName([string]$Text) {
    $s = ($Text -replace '[^A-Za-z0-9_.-]+', '-').Trim('-', '.')
    if (-not $s) { $s = 'page' }
    if ($s.Length -gt 60) { $s = $s.Substring(0, 60) }
    $s
}

# The page as a summary (title, record IDs, outlines with position, size and lines),
# and its XML saved in the session folder.
function Get-ReviewOneNotePage {
    param([Parameter(Mandatory)][string]$Page, [string]$Notebook, [string]$MetaName)
    $p = Resolve-ReviewOneNotePage $Page $Notebook $MetaName
    $xml = Get-ReviewOneNotePageXml $p.Id
    $file = Join-Path (Get-ReviewOneNoteOut) ((Get-ReviewSafeName $(if ($p.Key) { $p.Key } else { $p.Title })) + '.xml')
    [IO.File]::WriteAllText($file, $xml, (New-Object Text.UTF8Encoding($false)))
    $s = ConvertFrom-OneNotePage $xml
    $s | Add-Member -NotePropertyName Section -NotePropertyValue $p.Section
    $s | Add-Member -NotePropertyName XmlPath -NotePropertyValue $file
    $s
}

# Opens the page in a new OneNote window, keeps that window off the desktop (offscreen)
# or at the top left (onscreen), draws it with PrintWindow until two frames match, and
# closes the window. The user's own OneNote windows are not moved.
function Save-ReviewOneNoteCapture {
    param([Parameter(Mandatory)][string]$Page, [string]$Notebook, [string]$MetaName, [string]$Label,
          [ValidateSet('offscreen', 'onscreen')][string]$Mode = 'offscreen', [int]$WidthPx = 1280, [int]$HeightPx = 900)
    if ($Label -and $Label -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') { throw "Capture label '$Label' must be letters, digits, '.', '_' or '-'." }
    $p = Resolve-ReviewOneNotePage $Page $Notebook $MetaName
    $on = Get-ReviewOneNote
    $proc = @(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)
    if (-not $proc.Count) { throw 'OneNote is not running.' }
    $pidOn = [uint32]$proc[0].Id
    $before = @([KitReviewWin]::TopWindows($pidOn))
    $on.NavigateTo($p.Id, '', $true)   # a new window, so the user's window keeps its page
    $win = [IntPtr]::Zero
    $v = [KitReviewWin]::VirtualScreen()
    $w = [Math]::Min($WidthPx, $v[2]); $h = [Math]::Min($HeightPx, $v[3])
    for ($k = 0; $k -lt 100 -and $win -eq [IntPtr]::Zero; $k++) {
        foreach ($cand in [KitReviewWin]::TopWindows($pidOn)) {
            if ($before -contains $cand) { continue }
            # OneNote's page window is Framework::CFrame; any other new window of a
            # usable size is taken too, in case a later OneNote renames the class.
            $r = [KitReviewWin]::Rect($cand)
            if ([KitReviewWin]::Cls($cand) -like 'Framework::CFrame*' -or ($k -ge 30 -and $r[2] -ge 400 -and $r[3] -ge 300)) { $win = $cand; break }
        }
        if ($win -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 100 }
    }
    if ($win -eq [IntPtr]::Zero) { throw 'OneNote did not open a new window for the page within 10 s.' }
    try {
        if ($Mode -eq 'offscreen') { [void][KitReviewWin]::Park($win) } else { [void][KitReviewWin]::Show($win, 0, 0) }
        [void][KitReviewWin]::Size($win, $w, $h)
        if ($Mode -eq 'offscreen') { [void][KitReviewWin]::Park($win) }
        Start-Sleep -Milliseconds 800
        $name = ((Get-ReviewSafeName $(if ($p.Key) { $p.Key } else { $p.Title })), $Label | Where-Object { $_ }) -join '-'
        $path = Join-Path (Get-ReviewOneNoteOut) ($name + '.png')
        $sha = [Security.Cryptography.SHA256]::Create()
        $prev = $null; $stable = $false; $res = $null
        for ($k = 0; $k -lt 12; $k++) {
            $res = [KitReviewWin]::Save($win, $path, 0, 0, 0, 0, 1)
            $hash = [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($path)))
            if ($hash -eq $prev) { $stable = $true; break }
            $prev = $hash
            Start-Sleep -Milliseconds 400
        }
    } finally {
        [void][KitReviewWin]::Close($win)
    }
    [pscustomobject][ordered]@{
        Path = $path; Page = "$($p.Section) / $($p.Title)"; Key = $p.Key; ImagePx = "$($res[0])x$($res[1])"; Mode = $Mode; Stable = $stable
        Warning = $(if ($res[2]) { 'The capture is one solid colour: the session may be locked, or OneNote does not draw off-screen here; try mode onscreen on a review machine.' }
                    elseif (-not $stable) { 'The page was still changing after 5 s of captures; this frame may be mid-draw.' } else { $null })
    }
}

function Get-ReviewOneNoteSnapshotDir([string]$Label) {
    if ($Label -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') { throw "Snapshot label '$Label' must be letters, digits, '.', '_' or '-'." }
    Join-Path (Join-Path (Get-ReviewOutputBase) 'onenote-snapshots') $Label
}

# Saves every page's XML of the notebook (or section) under onenote-snapshots\<Label>,
# replacing that label's earlier snapshot. Pages are keyed by record ID (MetaName) when
# they have one, else by section and title, so compare can match them.
function Save-ReviewOneNoteSnapshot {
    param([Parameter(Mandatory)][string]$Label, [string]$Notebook, [string]$Section, [string]$MetaName)
    $dir = Get-ReviewOneNoteSnapshotDir $Label
    $pages = @(Get-ReviewOneNotePageList $Notebook $Section $MetaName)
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $index = @(); $seen = @{}; $i = 0
    $utf8 = New-Object Text.UTF8Encoding($false)
    foreach ($p in $pages) {
        $i++
        $key = if ($p.Key) { $p.Key } else { "$($p.Section)/$($p.Title)" }
        if ($seen.ContainsKey($key)) { $seen[$key]++; $key = "$key #$($seen[$key])" } else { $seen[$key] = 1 }
        $file = ('{0:000}-{1}.xml' -f $i, (Get-ReviewSafeName $key))
        [IO.File]::WriteAllText((Join-Path $dir $file), (Get-ReviewOneNotePageXml $p.Id), $utf8)
        $index += [pscustomobject][ordered]@{ Key = $key; Title = [string]$p.Title; Section = [string]$p.Section; File = $file }
    }
    [IO.File]::WriteAllText((Join-Path $dir 'index.json'), (ConvertTo-Json -InputObject @($index) -Depth 3), $utf8)
    [pscustomobject][ordered]@{ Label = $Label; Folder = $dir; Pages = $index.Count; WithRecordId = @($pages | Where-Object { $_.Key }).Count }
}

function Compare-ReviewOneNoteSnapshot {
    param([Parameter(Mandatory)][string]$Before, [Parameter(Mandatory)][string]$After)
    $bd = Get-ReviewOneNoteSnapshotDir $Before; $ad = Get-ReviewOneNoteSnapshotDir $After
    foreach ($d in $bd, $ad) { if (-not (Test-Path -LiteralPath (Join-Path $d 'index.json'))) { throw "No OneNote snapshot at $d. Take it with onenote_snapshot first." } }
    $cmp = Compare-OneNoteSnapshotFolder $bd $ad
    $text = Format-OneNoteSnapshotReport $cmp "OneNote snapshots: $Before -> $After"
    $report = Join-Path (Split-Path $bd) "compare $Before to $After.txt"
    [IO.File]::WriteAllText($report, $text, (New-Object Text.UTF8Encoding($false)))
    $lines = @($text -split "`r`n")
    $changes = @($lines | Select-Object -Skip 2)
    if ($changes.Count -gt 400) { $changes = @($changes[0..399]) + "... $($changes.Count - 400) more lines in the report" }
    [pscustomobject][ordered]@{ Report = $report; Summary = $lines[1]; Changes = $changes }
}

function Test-ReviewOneNoteRecordIds {
    param([string]$Notebook, [string]$Section, [string]$MetaName = 'spm.key', [string[]]$Expected)
    $pages = @(Get-ReviewOneNotePageList $Notebook $Section $MetaName)
    Test-OneNoteMetaKeys $pages $MetaName $Expected
}

function Close-ReviewOneNote {
    if ($script:OneNote) { try { Release-Com $script:OneNote } catch {} ; $script:OneNote = $null }
}
