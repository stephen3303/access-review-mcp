# Pure helpers for the Office COM tools (OfficeReview.ps1): member paths, the safety
# rules every call passes, and conversion between JSON values and COM values. Nothing
# here touches COM, so OfficeCom.Tests.ps1 checks it on any machine.
#
# A member path is VBA-like: dotted members, each optionally with arguments in brackets.
#   Worksheets('Summary').Range("A1:C4").Value2
#   Slides(2).Shapes('Title 1').TextFrame.TextRange.Text
#   SaveAs2(FileName:='C:\Review\out.docx', FileFormat:=16)
# Arguments are literals: 'text' or "text" (a doubled quote inside is one quote),
# numbers, true, false, null, missing (or an empty slot, as in Foo(, 2)), and @h3 for an
# object handle an earlier call returned. Name:=value passes a named argument.

$ErrorActionPreference = 'Stop'

# Members no call may use, on any app: they run code the caller supplies (formula and
# macro engines, DDE, the VBA editor), start programs, send input or mail, schedule
# work, load add-ins, print to paper or turn macros on for every file. Application.Run has its own tool
# (office_run_macro), quitting has office_quit, printing has office_export_pdf.
$script:OfficeDeniedMembers = @(
    'ExecuteExcel4Macro', 'Evaluate', '_Evaluate', 'Eval', 'Run', 'RunAutoMacro', 'RunAutoMacros', 'RunMacro', 'MacroScript', 'WordBasic',
    'DDEInitiate', 'DDEExecute', 'DDEPoke', 'DDERequest', 'DDETerminate', 'DDETerminateAll',
    'Shell', 'SendKeys', 'OnKey', 'OnTime', 'OnRepeat', 'OnUndo', 'OnData', 'OnDoubleClick', 'OnEntry', 'OnSheetActivate', 'OnSheetDeactivate', 'OnWindow',
    'VBE', 'VBProject', 'RegisterXLL', 'AddIns', 'AddIns2', 'COMAddIns', 'ExecuteMso', 'CommandBars',
    'FollowHyperlink', 'Follow', 'Navigate',
    'Send', 'SendMail', 'SendForReview', 'SendFax', 'SendFaxOverInternet', 'Route', 'RouteDocument', 'RoutingSlip', 'MailEnvelope', 'Dial',
    'PrintOut', 'PrintPreview', 'Quit', 'CreateObject', 'GetObject', 'Kill',
    'AutomationSecurity'   # macros are turned on per file by office_open enable_macros
)

# Further members refused per app: Outlook works on a real mailbox, so items are read,
# never deleted, moved or sent.
$script:OfficeDeniedByApp = @{
    outlook = @('Delete', 'PermanentDelete', 'Move', 'MoveTo', 'Forward', 'Reply', 'ReplyAll', 'Respond', 'ClearConversationIndex', 'AddStore', 'AddStoreEx', 'RemoveStore', 'Logon', 'Logoff')
    access  = @('RunSQL', 'RunCode', 'RunCommand', 'RunSavedImportExport', 'TransferSharePointList', 'CloseCurrentDatabase', 'OpenCurrentDatabase', 'NewCurrentDatabase', 'CompactRepair', 'ConvertAccessProject', 'CreateAccessProject')
}

# Methods that read or write files. A string argument of theirs must be a full path
# (a bare name would land in the app's default folder) inside an allowed folder.
$script:OfficeFileMembers = @(
    'Open', 'OpenEx', 'OpenText', 'OpenXML', 'OpenDatabase', 'OpenNoRepairDialog', 'SaveAs', 'SaveAs2', 'SaveAsAXE', 'SaveCopyAs', 'SaveAsFile', 'SaveToFile', 'SaveAsText', 'LoadFromText',
    'ExportAsFixedFormat', 'ExportAsFixedFormat2', 'ExportAsFixedFormat3', 'Export', 'ExportXML', 'ExportXml', 'ImportXML', 'ImportXml', 'Import',
    'InsertFile', 'InsertFromFile', 'AddFromFile', 'AddPicture', 'AddPicture2', 'AddOLEObject', 'AddMediaObject', 'AddMediaObject2', 'AddMediaObjectFromEmbedTag', 'UserPicture',
    'TransferText', 'TransferSpreadsheet', 'TransferDatabase', 'OutputTo', 'CopyFile', 'SaveAsPicture', 'SaveAsFixedFormat', 'MergeDocuments', 'CompareDocuments', 'Compare', 'Merge', 'ApplyTemplate', 'ApplyTheme', 'LoadThemeColors', 'LoadThemeFonts', 'LoadTheme'
)

# One member of a path: Name, its argument values, the names of the named ones (named
# arguments come first in Args, as Type.InvokeMember wants) and whether brackets were given.
function New-OfficeSegment([string]$Name, [object[]]$Positional, [object[]]$NamedValues, [string[]]$Names, [bool]$Call) {
    [pscustomobject]@{ Name = $Name; Args = @(@($NamedValues) + @($Positional)); Names = @($Names); HasArgs = $Call }
}

# Stand-ins the parser uses before the COM layer swaps in real values.
function New-OfficeHandleRef([string]$Id) { [pscustomobject]@{ OfficeHandle = $Id } }
$script:OfficeMissing = [pscustomobject]@{ OfficeMissing = $true }

function ConvertFrom-OfficePath {
    param([AllowEmptyString()][string]$Path)
    $segs = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Path)) { return @() }
    $s = $Path.Trim(); $i = 0; $n = $s.Length
    $fail = { param($why) throw "Member path '$Path': $why (at character $($i + 1))." }
    while ($true) {
        while ($i -lt $n -and $s[$i] -eq ' ') { $i++ }
        $start = $i
        while ($i -lt $n -and ($s[$i] -match '[A-Za-z0-9_]')) { $i++ }
        $name = $s.Substring($start, $i - $start)
        if ($name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { & $fail 'expected a member name' }
        $pos = @(); $namedV = @(); $names = @(); $call = $false
        while ($i -lt $n -and $s[$i] -eq ' ') { $i++ }
        if ($i -lt $n -and $s[$i] -eq '(') {
            $call = $true; $i++
            $slot = 0
            while ($true) {
                while ($i -lt $n -and $s[$i] -eq ' ') { $i++ }
                if ($i -ge $n) { & $fail 'missing )' }
                if ($s[$i] -eq ')' -and $slot -eq 0 -and $pos.Count -eq 0 -and $names.Count -eq 0) { $i++; break }
                # optional Name:=
                $argName = $null
                if ($s.Substring($i) -match '^([A-Za-z_][A-Za-z0-9_]*)\s*:=') { $argName = $Matches[1]; $i += $Matches[0].Length; while ($i -lt $n -and $s[$i] -eq ' ') { $i++ } }
                $val = $null
                if ($i -lt $n -and ($s[$i] -eq ',' -or $s[$i] -eq ')')) {
                    if ($argName) { & $fail "no value for $argName" }
                    $val = $script:OfficeMissing
                } elseif ($s[$i] -eq "'" -or $s[$i] -eq '"') {
                    $q = $s[$i]; $i++; $sb = New-Object Text.StringBuilder
                    while ($true) {
                        if ($i -ge $n) { & $fail 'unclosed text' }
                        if ($s[$i] -eq $q) { if ($i + 1 -lt $n -and $s[$i + 1] -eq $q) { [void]$sb.Append($q); $i += 2; continue }; $i++; break }
                        [void]$sb.Append($s[$i]); $i++
                    }
                    $val = $sb.ToString()
                } elseif ($s.Substring($i) -match '^-?\d+(\.\d+)?([eE][+-]?\d+)?') {
                    $t = $Matches[0]; $i += $t.Length
                    if ($t -match '^-?\d+$' -and [long]$t -ge [int]::MinValue -and [long]$t -le [int]::MaxValue) { $val = [int]$t } else { $val = [double]::Parse($t, [Globalization.CultureInfo]::InvariantCulture) }
                } elseif ($s.Substring($i) -match '^@(h\d+)\b') {
                    $val = New-OfficeHandleRef $Matches[1]; $i += $Matches[0].Length
                } elseif ($s.Substring($i) -match '^(true|false|null|missing)\b') {
                    $w = $Matches[1].ToLowerInvariant(); $i += $w.Length
                    $val = switch ($w) { 'true' { $true } 'false' { $false } 'null' { $null } 'missing' { $script:OfficeMissing } }
                } else { & $fail 'expected a value: text in quotes, a number, true, false, null, missing or @h<n>' }
                if ($argName) {
                    if ($names -contains $argName) { & $fail "$argName given twice" }
                    $names += $argName; $namedV += , $val
                } else {
                    if ($names.Count) { & $fail 'a positional argument after a named one' }
                    $pos += , $val
                }
                $slot++
                while ($i -lt $n -and $s[$i] -eq ' ') { $i++ }
                if ($i -lt $n -and $s[$i] -eq ',') { $i++; continue }
                if ($i -lt $n -and $s[$i] -eq ')') { $i++; break }
                & $fail 'expected , or )'
            }
        }
        $segs.Add((New-OfficeSegment $name $pos $namedV $names $call))
        while ($i -lt $n -and $s[$i] -eq ' ') { $i++ }
        if ($i -ge $n) { break }
        if ($s[$i] -ne '.') { & $fail 'expected .' }
        $i++
    }
    $segs.ToArray()   # callers wrap the result in @()
}

# Why a member may not be used, or $null.
function Get-OfficeMemberRefusal([string]$App, [string]$Member) {
    if ($script:OfficeDeniedMembers -contains $Member) { return "'$Member' is not allowed: it runs caller-supplied code, starts programs, sends input or mail, prints, or quits. Use the dedicated tool (office_run_macro, office_quit, office_export_pdf) where there is one." }
    $extra = $script:OfficeDeniedByApp[$App]
    if ($extra -and $extra -contains $Member) { return "'$Member' is not allowed on $App$(if ($App -eq 'outlook') { ': Outlook items are read only' } else { ': use the Access review tools' })." }
    $null
}

function Test-OfficeFullPath([string]$Path) { $Path -match '^[A-Za-z]:[\\/]' -or $Path -match '^[\\/]{2}[^\\/]' }

# True for text that names a file or a web address rather than ordinary text.
function Test-OfficePathLike([string]$Text) {
    if (Test-OfficeFullPath $Text) { return $true }
    if ($Text -match '^\s*(file|https?|ftps?|mhtml|ms-[a-z-]+|onenote|outlook|smb|webdav|res|javascript|vbscript|data|mailto|tel|news|nntp|ldap)\s*:') { return $true }
    $Text -match '^\s*www\.'
}

# Text a cell, field or shape must not be given: a DDE link (=cmd|' /c ...'!A0) or a
# formula function that calls code, the registry, a COM server or the web.
function Get-OfficeValueRefusal([string]$Text) {
    if ($Text -notmatch '^\s*[=+\-@]') { return $null }
    if ($Text -match '\|') { return 'a formula with | is a DDE link, which can start programs.' }
    if ($Text -match '(?i)\b(CALL|REGISTER|REGISTER\.ID|EXEC|RTD|WEBSERVICE|FILTERXML|HYPERLINK|IMPORTDATA|DDE|EXECUTE|RUN|SHELL)\s*\(') { return "the formula calls $($Matches[1]), which reaches code, other programs or the web." }
    $null
}

# Every string inside a value (arrays and nested arrays included).
function Get-OfficeStrings($Value) {
    if ($null -eq $Value) { return }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [Array]) { foreach ($v in $Value) { Get-OfficeStrings $v } }
}

# Checks one call before it is made. $InAllowed tells whether a full path may be used
# (inside an allowed folder or the output folder). Throws the reason when refused.
function Assert-OfficeCallAllowed {
    param([string]$App, [string]$Member, [object[]]$Arguments = @(), [scriptblock]$InAllowed, [switch]$IsSet)
    $why = Get-OfficeMemberRefusal $App $Member
    if ($why) { throw $why }
    $strings = @(foreach ($a in $Arguments) { Get-OfficeStrings $a })
    foreach ($t in $strings) {
        $bad = Get-OfficeValueRefusal $t
        if ($bad) { throw "Refused value for ${Member}: $bad" }
        if (Test-OfficePathLike $t) {
            if (-not (Test-OfficeFullPath $t)) { throw "Refused value for ${Member}: '$t' is a web or other non-file address; only files in the allowed folders are used." }
            if ($t -match '\.\.[\\/]|[\\/]\.\.$') { throw "Refused value for ${Member}: '$t' contains '..'; give the real full path." }
            if (-not (& $InAllowed $t)) { throw "Refused value for ${Member}: '$t' is outside the allowed folders and the output folder." }
        }
    }
    if ($script:OfficeFileMembers -contains $Member -and $strings.Count -and -not @($strings | Where-Object { Test-OfficeFullPath $_ }).Count) {
        throw "$Member reads or writes a file: give it a full path inside an allowed folder (a bare name would use the app's default folder)."
    }
}

# The plain .NET object behind a PowerShell wrapper (COM can't take the wrapper).
function Get-OfficeBase($v) { if ($null -ne $v -and $v.PSObject) { $v = $v.PSObject.BaseObject }; if ($v -is [Array]) { return , $v }; $v }

# A JSON value (from ConvertFrom-Json) as a COM argument: whole numbers as Int32 (a VBA
# Long rejects Int64) or Double, decimals as Double, an array of equal-length arrays
# as a 2-D array (a block of cells), {"handle":"h3"} as that object. $Resolve maps a
# handle ID to its object.
function ConvertTo-OfficeComValue($Value, [scriptblock]$Resolve) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [long]) { if ($Value -ge [int]::MinValue -and $Value -le [int]::MaxValue) { return [int]$Value } else { return [double]$Value } }
    if ($Value -is [decimal]) { return [double]$Value }
    if ($Value -is [string] -or $Value -is [bool] -or $Value -is [int] -or $Value -is [double]) { return $Value }
    if ($Value -is [Array]) {
        $rows = @($Value)
        if ($rows.Count -and @($rows | Where-Object { $_ -is [Array] }).Count -eq $rows.Count) {
            $w = @($rows[0]).Count
            if (@($rows | Where-Object { @($_).Count -ne $w }).Count) { throw 'A block of values needs rows of the same length.' }
            $grid = New-Object 'object[,]' $rows.Count, $w
            for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt $w; $c++) { $grid[$r, $c] = Get-OfficeBase (ConvertTo-OfficeComValue @($rows[$r])[$c] $Resolve) } }
            return , $grid
        }
        $list = New-Object object[] $rows.Count
        for ($k = 0; $k -lt $rows.Count; $k++) { $list[$k] = Get-OfficeBase (ConvertTo-OfficeComValue $rows[$k] $Resolve) }
        return , $list
    }
    if ($Value.PSObject.Properties['handle'] -and @($Value.PSObject.Properties).Count -eq 1) { return (& $Resolve ([string]$Value.handle)) }
    if ($Value.PSObject.Properties['OfficeHandle']) { return (& $Resolve ([string]$Value.OfficeHandle)) }
    throw "Unsupported argument value: $($Value | ConvertTo-Json -Compress -Depth 3)"
}

# A COM result as a JSON-friendly value. $OnObject turns a COM object into a handle
# record. Blocks of cells (2-D arrays, often 1-based from Excel) become rows; at most
# MaxCells values are returned, and Truncated says when there were more.
function ConvertFrom-OfficeComValue($Value, [scriptblock]$OnObject, [int]$MaxCells = 10000) {
    if ($null -eq $Value -or $Value -is [DBNull] -or $Value -is [Reflection.Missing]) { return $null }
    if ($Value -is [DateTime]) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss') }
    if ($Value -is [string] -or $Value -is [ValueType]) { return $Value }
    if ($Value -is [Array]) {
        if ($Value.Rank -eq 2) {
            $r0 = $Value.GetLowerBound(0); $r1 = $Value.GetUpperBound(0); $c0 = $Value.GetLowerBound(1); $c1 = $Value.GetUpperBound(1)
            $cols = $c1 - $c0 + 1; $count = 0; $cut = $false
            $rows = @(for ($r = $r0; $r -le $r1; $r++) {
                if ($count + $cols -gt $MaxCells) { $cut = $true; break }
                $count += $cols
                , @(for ($c = $c0; $c -le $c1; $c++) { ConvertFrom-OfficeComValue $Value[$r, $c] $OnObject $MaxCells })
            })
            return [pscustomobject]@{ Rows = $rows; RowCount = $r1 - $r0 + 1; ColumnCount = $cols; Truncated = $cut }
        }
        if ($Value.Rank -eq 1) {
            $list = @($Value)
            $cut = $list.Count -gt $MaxCells
            $vals = @($list | Select-Object -First $MaxCells | ForEach-Object { ConvertFrom-OfficeComValue $_ $OnObject $MaxCells })
            if ($cut) { return [pscustomobject]@{ Items = $vals; Count = $list.Count; Truncated = $true } }
            return , $vals
        }
        throw "A $($Value.Rank)-dimensional array is not supported."
    }
    if ($OnObject) { return (& $OnObject $Value) }
    [string]$Value
}

# One audit line: what was called, never the text it was given beyond a short preview.
function Format-OfficeAudit {
    param([string]$Tool, [string]$App, [string]$Target, [string]$Path, $Arguments, [bool]$Ok, [string]$Failure)
    $preview = $null
    if ($null -ne $Arguments) {
        $preview = ($Arguments | ConvertTo-Json -Compress -Depth 4)
        if ($preview.Length -gt 300) { $preview = $preview.Substring(0, 300) + '...' }
    }
    [ordered]@{ time = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssK'); tool = $Tool; app = $App; target = $Target; path = $Path; args = $preview; ok = $Ok; error = $Failure } |
        ConvertTo-Json -Compress
}
