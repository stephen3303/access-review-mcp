# Form snapshots: a form's saved definition (the text Application.SaveAsText writes)
# read into objects and properties, so two builds of the same form can be compared
# property by property. A screenshot shows that something moved; a snapshot diff says
# which control, which property and by how many pixels, including changes a capture
# can't show (a hidden control, a changed event handler, a record source).
#
# Comparing needs no Access and no Windows API, so this file is separate from
# AccessReview.ps1 (which takes the snapshots) and can be tested anywhere PowerShell
# runs: see FormSnapshot.Tests.ps1.

# Properties Access rewrites on every save or derives from other properties. They'd
# show as changes when nothing anyone sees has changed.
$script:SnapshotIgnored = @(
    'Checksum', 'GUID', 'NameMap', 'RecSrcDt', 'Version', 'VersionRequired',
    'PrtMip', 'PrtDevMode', 'PrtDevModeW', 'PrtDevNames', 'PrtDevNamesW',
    'LayoutCachedLeft', 'LayoutCachedTop', 'LayoutCachedWidth', 'LayoutCachedHeight',
    'OverlapFlags'
)

# On the form itself (not its controls) these are the design window's rectangle,
# where the form last sat in design view, which differs from build to build.
$script:SnapshotRootIgnored = @('Left', 'Top', 'Right', 'Bottom')

# Properties stored in twips, reported in 96 dpi pixels (twips / 15) like the kit's px().
$script:SnapshotTwips = @(
    'Left', 'Top', 'Width', 'Height', 'LeftMargin', 'TopMargin', 'RightMargin', 'BottomMargin',
    'LeftPadding', 'TopPadding', 'RightPadding', 'BottomPadding'
)

# Section objects, keyed by their Name ("Section Detail") rather than as controls.
$script:SnapshotSections = @(
    'Section', 'FormHeader', 'FormFooter', 'PageHeader', 'PageFooter', 'BreakHeader', 'BreakFooter'
)

# Reads SaveAsText output into one object per Begin <Type> ... End block, in file
# order. Each has Key (the control's Name, "Section <name>", "Form"/"Report" for the
# root, or "Default <Type>" for the control-type defaults that have no name), Type and
# Props (property name -> text value as saved). Binary values (pictures, conditional
# formats) become a short hash so a change still shows. Code behind the form is not
# read: it is in source control and its diff is in the PR.
function ConvertFrom-FormText {
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Lines)
    $all = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.List[object]   # $null marks a plain Begin (a list of children)
    $blob = $null
    $last = $null      # property a continuation line ("...") adds to
    foreach ($raw in $Lines) {
        $line = $raw.Replace([string][char]0, '').Trim()
        if ($blob) {
            if ($line -eq 'End') {
                $blob.Depth--
                if ($blob.Depth -eq 0) {
                    $sha = [Security.Cryptography.SHA256]::Create()
                    $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($blob.Text.ToString()))).Replace('-', '').Substring(0, 12).ToLower()
                    if ($blob.Owner) { $blob.Owner.Props[$blob.Name] = "<binary, $($blob.Count) lines, $hash>" }
                    $blob = $null
                    continue
                }
            } elseif ($line -match '(^|\s)Begin$') { $blob.Depth++ }
            [void]$blob.Text.AppendLine($line)
            $blob.Count++
            continue
        }
        if (-not $line) { continue }
        if ($line -eq 'CodeBehindForm') { break }
        $owner = $null
        for ($i = $stack.Count - 1; $i -ge 0; $i--) { if ($stack[$i]) { $owner = $stack[$i]; break } }
        if ($line -match '^Begin\s+(\w+)$') {
            $o = [pscustomobject]@{ Key = $null; Type = $Matches[1]; Props = [ordered]@{}; Parent = $owner; Root = ($all.Count -eq 0) }
            $all.Add($o); $stack.Add($o); $last = $null
            continue
        }
        if ($line -eq 'Begin') { $stack.Add($null); $last = $null; continue }
        if ($line -eq 'End') {
            if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) }
            $last = $null
            continue
        }
        if ($line -match '^(\w+)\s*=\s*Begin$') {
            $name = $Matches[1]
            $blob = @{ Name = $name; Depth = 1; Count = 0; Text = (New-Object Text.StringBuilder)
                       Owner = $(if ($owner -and $SnapshotIgnored -notcontains $name) { $owner } else { $null }) }
            $last = $null
            continue
        }
        if ($line -match '^(\w+)\s*=\s?(.*)$') {
            $name = $Matches[1]
            if ($owner -and $SnapshotIgnored -notcontains $name -and -not ($owner.Root -and $SnapshotRootIgnored -contains $name)) {
                $owner.Props[$name] = $Matches[2]
                $last = @{ Owner = $owner; Name = $name }
            } else { $last = $null }
            continue
        }
        if ($line.StartsWith('"') -and $last) {
            # A long string continues on the next line as another quoted piece; join the
            # pieces into one quoted value.
            $prev = [string]$last.Owner.Props[$last.Name]
            if ($prev.EndsWith('"') -and $line.Length -gt 1) { $prev = $prev.Substring(0, $prev.Length - 1); $line = $line.Substring(1) }
            $last.Owner.Props[$last.Name] = $prev + $line
        }
    }

    $seen = @{}
    foreach ($o in $all) {
        $nm = if ($o.Props.Contains('Name')) { $o.Props['Name'].Trim('"') } else { '' }
        $key = if ($o.Root) { $o.Type }
               elseif ($SnapshotSections -contains $o.Type) { 'Section ' + $(if ($nm) { $nm } else { $o.Type }) }
               elseif ($nm) { $nm }
               else { "Default $($o.Type)" }
        if ($seen.ContainsKey($key)) { $seen[$key]++; $key = "$key #$($seen[$key])" } else { $seen[$key] = 1 }
        $o.Key = $key
    }
    , $all.ToArray()
}

# A saved value as a reviewer reads it: twips as px, colours as #RRGGBB (Access stores
# them as BGR numbers; negative ones are system colours and stay as saved).
function Format-SnapshotValue([string]$Name, $Value) {
    if ($null -eq $Value) { return '(not set)' }
    $n = 0L
    if ([long]::TryParse([string]$Value, [ref]$n)) {
        if ($SnapshotTwips -contains $Name) {
            $px = [Math]::Round($n / 15.0, 2)
            return "$px px"
        }
        if ($Name -match 'Color$' -and $n -ge 0 -and $n -le 0xFFFFFF) {
            return ('#{0:X2}{1:X2}{2:X2}' -f ($n -band 0xFF), (($n -shr 8) -band 0xFF), (($n -shr 16) -band 0xFF))
        }
    }
    [string]$Value
}

# Compares two parsed forms. Returns the objects only one side has (with type and
# position) and every property whose saved value differs, before and after.
function Compare-FormText {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Before, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$After)
    $b = [ordered]@{}; foreach ($o in $Before) { $b[$o.Key] = $o }
    $a = [ordered]@{}; foreach ($o in $After) { $a[$o.Key] = $o }
    $where = {
        param($o)
        $parts = foreach ($p in 'Left', 'Top', 'Width', 'Height') { if ($o.Props.Contains($p)) { "$p $(Format-SnapshotValue $p $o.Props[$p])" } }
        ($parts -join ', ')
    }
    $added = @(foreach ($k in $a.Keys) { if (-not $b.Contains($k)) { [pscustomobject]@{ Object = $k; Type = $a[$k].Type; Position = (& $where $a[$k]) } } })
    $removed = @(foreach ($k in $b.Keys) { if (-not $a.Contains($k)) { [pscustomobject]@{ Object = $k; Type = $b[$k].Type; Position = (& $where $b[$k]) } } })
    $changed = @(foreach ($k in $a.Keys) {
        if (-not $b.Contains($k)) { continue }
        $x = $b[$k]; $y = $a[$k]
        if ($x.Type -ne $y.Type) { [pscustomobject]@{ Object = $k; Property = '(control type)'; Before = $x.Type; After = $y.Type } }
        # Which section or tab page holds it is the Begin nesting, not a property.
        $px = if ($x.Parent) { $x.Parent.Key } else { '' }; $py = if ($y.Parent) { $y.Parent.Key } else { '' }
        if ($px -ne $py) { [pscustomobject]@{ Object = $k; Property = '(inside)'; Before = $px; After = $py } }
        $names = @($x.Props.Keys) + @($y.Props.Keys | Where-Object { -not $x.Props.Contains($_) })
        foreach ($p in $names) {
            $v1 = if ($x.Props.Contains($p)) { $x.Props[$p] } else { $null }
            $v2 = if ($y.Props.Contains($p)) { $y.Props[$p] } else { $null }
            if ($v1 -cne $v2) {
                [pscustomobject]@{ Object = $k; Property = $p; Before = (Format-SnapshotValue $p $v1); After = (Format-SnapshotValue $p $v2) }
            }
        }
    })
    [pscustomobject]@{ Added = $added; Removed = $removed; Changed = $changed }
}

# Reads a SaveAsText file (Access writes forms as UTF-16 with a byte order mark;
# ReadAllLines detects it) and parses it.
function Read-FormSnapshot([string]$Path) {
    ConvertFrom-FormText -Lines ([IO.File]::ReadAllLines($Path))
}

# Compares every <form>.txt in two snapshot folders. Returns the forms only one side
# has and, per form in both, the Compare-FormText result. With -Forms, only those.
function Compare-FormSnapshotFolder {
    param([Parameter(Mandatory)][string]$BeforeDir, [Parameter(Mandatory)][string]$AfterDir, [string[]]$Forms)
    foreach ($d in $BeforeDir, $AfterDir) { if (-not (Test-Path -LiteralPath $d -PathType Container)) { throw "No snapshot at $d." } }
    $list = { param($dir) @(Get-ChildItem -LiteralPath $dir -Filter *.txt | ForEach-Object { $_.BaseName } | Sort-Object) }
    # Forms whose save failed: a <form>.failed marker, or a line in unsaved-names.list
    # for a name no file can carry.
    $failedIn = { param($dir)
        $l = Join-Path $dir 'unsaved-names.list'
        @(@(Get-ChildItem -LiteralPath $dir -Filter *.failed | ForEach-Object { $_.BaseName }) +
          @(if (Test-Path -LiteralPath $l) { Get-Content -LiteralPath $l -Encoding UTF8 | Where-Object { $_ } })) }
    $bn = & $list $BeforeDir; $an = & $list $AfterDir
    # A form whose snapshot failed on either side can't be compared; it is listed as such
    # instead of looking added or removed.
    $bf = & $failedIn $BeforeDir; $af = & $failedIn $AfterDir
    $notCompared = @(@($bf | ForEach-Object { "$_ (snapshot failed in before)" }) + @($af | ForEach-Object { "$_ (snapshot failed in after)" }))
    $skip = @($bf) + @($af)
    if ($Forms) {
        $bn = @($bn | Where-Object { $Forms -contains $_ }); $an = @($an | Where-Object { $Forms -contains $_ })
        $notCompared = @($notCompared | Where-Object { $Forms -contains ($_ -replace ' \(snapshot failed.*$', '') })
    }
    $bn = @($bn | Where-Object { $skip -notcontains $_ }); $an = @($an | Where-Object { $skip -notcontains $_ })
    $results = foreach ($f in $an) {
        if ($bn -notcontains $f) { continue }
        $r = Compare-FormText -Before (Read-FormSnapshot (Join-Path $BeforeDir "$f.txt")) -After (Read-FormSnapshot (Join-Path $AfterDir "$f.txt"))
        [pscustomobject]@{ Form = $f; Added = $r.Added; Removed = $r.Removed; Changed = $r.Changed
                           Differences = $r.Added.Count + $r.Removed.Count + $r.Changed.Count }
    }
    [pscustomobject]@{
        FormsAdded = @($an | Where-Object { $bn -notcontains $_ })
        FormsRemoved = @($bn | Where-Object { $an -notcontains $_ })
        NotCompared = $notCompared
        Forms = @($results)
    }
}

# The comparison as plain text, one line per difference, for a review comment or a file.
function Format-FormSnapshotReport($Comparison, [string]$Title) {
    $sb = New-Object Text.StringBuilder
    if ($Title) { [void]$sb.AppendLine($Title) }
    foreach ($f in $Comparison.FormsAdded) { [void]$sb.AppendLine("+ form $f") }
    foreach ($f in $Comparison.FormsRemoved) { [void]$sb.AppendLine("- form $f") }
    foreach ($f in @($Comparison.NotCompared)) { [void]$sb.AppendLine("? form $f, not compared") }
    $same = @($Comparison.Forms | Where-Object { $_.Differences -eq 0 } | ForEach-Object { $_.Form })
    foreach ($f in @($Comparison.Forms | Where-Object { $_.Differences -gt 0 })) {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("$($f.Form): $($f.Differences) difference(s)")
        foreach ($x in $f.Added) { [void]$sb.AppendLine("  + $($x.Object) ($($x.Type)) $($x.Position)") }
        foreach ($x in $f.Removed) { [void]$sb.AppendLine("  - $($x.Object) ($($x.Type)) $($x.Position)") }
        foreach ($x in $f.Changed) { [void]$sb.AppendLine("  ~ $($x.Object).$($x.Property): $($x.Before) -> $($x.After)") }
    }
    if ($same.Count) { [void]$sb.AppendLine(); [void]$sb.AppendLine("Unchanged: $($same -join ', ')") }
    $sb.ToString()
}
