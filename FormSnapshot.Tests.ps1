# Checks FormSnapshot.ps1 against a small SaveAsText form, without Access:
#   powershell -NoProfile -File tools\review-mcp\FormSnapshot.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'FormSnapshot.ps1')

$base = @'
Version =20
VersionRequired =20
Checksum =-123456
Begin Form
    RecordSelectors = NotDefault
    Width =7200
    Left =-1590
    Top =120
    Right =11490
    Bottom =9180
    PrtMip = Begin
        0x6801000068010000
    End
    RecordSource ="SELECT id, title FROM items "
        "WHERE active = True"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Section
            Height =3000
            Name ="Detail"
            GUID = Begin
                0x0123456789abcdef
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =240
                    Width =1800
                    Height =300
                    ForeColor =8355711
                    Name ="lblTitle"
                    Caption ="Title"
                    LayoutCachedLeft =240
                End
                Begin TextBox
                    Left =360
                    Top =600
                    Width =3000
                    Height =330
                    Name ="txtTitle"
                    OnClick ="[Event Procedure]"
                End
                Begin Image
                    Left =4000
                    Top =240
                    Width =480
                    Height =480
                    Name ="imgPill"
                    PictureData = Begin
                        0x89504e470d0a1a0a
                        0x0000000d49484452
                    End
                End
            End
        End
    End
End
CodeBehindForm
Option Compare Database
'@ -split "`r?`n"

# The same form saved again after a build that moved the label 4 px, recoloured it,
# redrew the pill, removed the text box's click handler, added a button and
# removed nothing else. Checksum, GUID, LayoutCached* and the form's own design
# window rectangle (Left/Top/Right/Bottom on the form) change on every build.
$pr = $base -replace 'Checksum =-123456', 'Checksum =999' `
            -replace 'Right =11490', 'Right =11310' -replace 'Bottom =9180', 'Bottom =7800' `
            -replace '0x0123456789abcdef', '0xfedcba9876543210' `
            -replace '                    Left =240\r?$', '                    Left =300' `
            -replace 'LayoutCachedLeft =240', 'LayoutCachedLeft =300' `
            -replace 'ForeColor =8355711', 'ForeColor =16711680' `
            -replace '0x0000000d49484452', '0x0000000d494844ff' `
            -replace '"WHERE active = True"', '"WHERE active = False"'
$pr = @($pr | Where-Object { $_ -notmatch 'OnClick' })
$i = [array]::IndexOf($pr, '                Begin Image')
$pr = $pr[0..($i - 1)] + @(
    '                Begin CommandButton',
    '                    Left =240',
    '                    Top =1200',
    '                    Width =1500',
    '                    Height =420',
    '                    Name ="btnSave"',
    '                End') + $pr[$i..($pr.Count - 1)]

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}

$before = ConvertFrom-FormText -Lines $base
Check 'objects keyed by name, section and default' (($before.Key -join ',') -eq 'Form,Default Label,Section Detail,lblTitle,txtTitle,imgPill')
Check 'continued string joined' ($before[0].Props['RecordSource'] -eq '"SELECT id, title FROM items WHERE active = True"')
Check 'volatile properties dropped' (-not $before[0].Props.Contains('PrtMip') -and -not $before[2].Props.Contains('GUID') -and -not $before[3].Props.Contains('LayoutCachedLeft') -and -not $before[3].Props.Contains('OverlapFlags'))
Check 'form design window ignored, form width kept' (-not $before[0].Props.Contains('Right') -and -not $before[0].Props.Contains('Left') -and $before[0].Props['Width'] -eq '7200')
Check 'control Left and Top kept' ($before[3].Props['Left'] -eq '240' -and $before[3].Props['Top'] -eq '240')
Check 'picture kept as a hash' ($before[5].Props['PictureData'] -match '^<binary, 2 lines, [0-9a-f]{12}>$')
Check 'code behind not read' (-not ($before | Where-Object { $_.Props.Contains('Option') }))

$diff = Compare-FormText -Before $before -After (ConvertFrom-FormText -Lines $pr)
$changes = @($diff.Changed | ForEach-Object { "$($_.Object).$($_.Property): $($_.Before) -> $($_.After)" })
Check 'button added with position' ($diff.Added.Count -eq 1 -and $diff.Added[0].Object -eq 'btnSave' -and $diff.Added[0].Position -eq 'Left 16 px, Top 80 px, Width 100 px, Height 28 px')
Check 'nothing removed' ($diff.Removed.Count -eq 0)
Check 'move reported in px' ($changes -contains 'lblTitle.Left: 16 px -> 20 px')
Check 'colour reported as hex' ($changes -contains 'lblTitle.ForeColor: #7F7F7F -> #0000FF')
Check 'removed handler reported' ($changes -contains 'txtTitle.OnClick: "[Event Procedure]" -> (not set)')
Check 'record source change reported' (@($diff.Changed | Where-Object { $_.Property -eq 'RecordSource' }).Count -eq 1)
Check 'redrawn picture reported' (@($diff.Changed | Where-Object { $_.Object -eq 'imgPill' -and $_.Property -eq 'PictureData' }).Count -eq 1)
Check 'only real changes' ($diff.Changed.Count -eq 5)
$moved = @(
    'Begin Form', 'Begin',
    'Begin FormHeader', 'Name ="FormHeader"', 'Begin', 'Begin Label', 'Name ="lblA"', 'Left =0', 'End', 'End', 'End',
    'Begin Section', 'Name ="Detail"', 'Begin', 'End', 'End',
    'End', 'End')
$movedTo = $moved[0..1] + $moved[2..4] + @('End', 'End') + $moved[11..13] + $moved[5..8] + $moved[14..17]
$m = Compare-FormText -Before (ConvertFrom-FormText -Lines $moved) -After (ConvertFrom-FormText -Lines $movedTo)
Check 'move to another section reported' ($m.Changed.Count -eq 1 -and $m.Changed[0].Property -eq '(inside)' -and $m.Changed[0].Before -eq 'Section FormHeader' -and $m.Changed[0].After -eq 'Section Detail')
$same = Compare-FormText -Before $before -After (ConvertFrom-FormText -Lines $base)
Check 'identical forms compare equal' (($same.Added.Count + $same.Removed.Count + $same.Changed.Count) -eq 0)

# Folders as AccessReview.ps1 writes them: UTF-16 files named <form>.txt.
$root = Join-Path ([IO.Path]::GetTempPath()) ('snaptest-' + [guid]::NewGuid().ToString('N'))
$b = Join-Path $root 'base'; $a = Join-Path $root 'pr'
New-Item -ItemType Directory -Force $b, $a | Out-Null
try {
    [IO.File]::WriteAllLines((Join-Path $b 'frmItem.txt'), [string[]]$base, [Text.Encoding]::Unicode)
    [IO.File]::WriteAllLines((Join-Path $b 'frmOld.txt'), [string[]]$base, [Text.Encoding]::Unicode)
    [IO.File]::WriteAllLines((Join-Path $a 'frmItem.txt'), [string[]]$pr, [Text.Encoding]::Unicode)
    [IO.File]::WriteAllLines((Join-Path $a 'frmNew.txt'), [string[]]$base, [Text.Encoding]::Unicode)
    $cmp = Compare-FormSnapshotFolder -BeforeDir $b -AfterDir $a
    Check 'forms added and removed' (($cmp.FormsAdded -join ',') -eq 'frmNew' -and ($cmp.FormsRemoved -join ',') -eq 'frmOld')
    Check 'UTF-16 files compared' ($cmp.Forms.Count -eq 1 -and $cmp.Forms[0].Differences -eq 6)
    $report = Format-FormSnapshotReport $cmp 'base -> pr'
    Check 'report lists each difference' ($report -match '\+ btnSave \(CommandButton\)' -and $report -match '~ lblTitle\.Left: 16 px -> 20 px' -and $report -match '\+ form frmNew')

    # A form whose snapshot failed on one side (a <form>.failed marker, no .txt) is
    # reported as not compared, not as removed.
    [IO.File]::WriteAllLines((Join-Path $b 'frmBroken.txt'), [string[]]$base, [Text.Encoding]::Unicode)
    [IO.File]::WriteAllText((Join-Path $a 'frmBroken.failed'), 'SaveAsText failed')
    $cmp = Compare-FormSnapshotFolder -BeforeDir $b -AfterDir $a
    Check 'failed snapshot not compared' (($cmp.FormsRemoved -join ',') -eq 'frmOld' -and ($cmp.NotCompared -join ',') -eq 'frmBroken (snapshot failed in after)')
    [IO.File]::WriteAllText((Join-Path $b 'unsaved-names.list'), "frm:Odd`r`n")
    $cmp = Compare-FormSnapshotFolder -BeforeDir $b -AfterDir $a
    Check 'unsaveable name not compared' (@($cmp.NotCompared) -contains 'frm:Odd (snapshot failed in before)')
    Check 'report lists not compared' ((Format-FormSnapshotReport $cmp 'x') -match '\? form frmBroken \(snapshot failed in after\), not compared')
} finally { Remove-Item -Recurse -Force $root }

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
