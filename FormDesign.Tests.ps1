# Checks FormDesign.ps1 against a small SaveAsText form, without Access:
#   powershell -NoProfile -File tools\review-mcp\FormDesign.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'FormDesign.ps1')

$form = @'
Version =20
VersionRequired =20
Checksum =-123456
Begin Form
    PopUp = NotDefault
    Width =7200
    Left =-1590
    Top =120
    Right =11490
    Bottom =9180
    Caption ="Edit ""item"""
    RecordSource ="SELECT id, title FROM items "
        "WHERE active = True"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Segoe UI"
            ForeColor =2758415
        End
        Begin FormHeader
            Height =900
            BackColor =6044187
            Name ="FormHeader"
            Begin
                Begin Label
                    Left =240
                    Top =240
                    Width =3000
                    Height =420
                    ForeColor =16777215
                    Name ="lblHeading"
                    Caption ="Items"
                End
            End
        End
        Begin Section
            Height =3000
            BackColor =16711679
            Name ="Detail"
            GUID = Begin
                0x0123456789abcdef
            End
            Begin
                Begin TextBox
                    Left =360
                    Top =600
                    Width =3000
                    Height =330
                    BorderColor =-2147483632
                    Name ="txtTitle"
                    ControlSource ="title"
                    OnClick ="[Event Procedure]"
                    AfterUpdate ="=SaveTitle()"
                End
                Begin Tab
                    Left =240
                    Top =1200
                    Width =6000
                    Height =1500
                    TabFixedHeight =480
                    Name ="tabMain"
                    Begin
                        Begin Page
                            Left =300
                            Top =1680
                            Width =5880
                            Height =960
                            Name ="pgOne"
                            Caption ="One"
                            Begin
                                Begin Label
                                    Left =420
                                    Top =1800
                                    Width =1800
                                    Height =300
                                    Name ="lblOnPage"
                                    Caption ="On the page"
                                End
                            End
                        End
                    End
                End
                Begin Image
                    Left =4000
                    Top =240
                    Width =480
                    Height =480
                    Name ="imgPill"
                    PictureData = Begin
                        0x89504e470d0a1a0a
                    End
                End
            End
        End
    End
End
CodeBehindForm
Option Compare Database
'@ -split "`r?`n"

$theme = @'
Option Compare Database
Option Explicit
Public Const THEME_DARK As Boolean = False
Public Const CLR_BG As Long = 16711679        ' RGB(255, 255, 254)
Public Const CLR_FG As Long = 2758415         ' RGB(15, 23, 42)
Public Const CLR_WHITE As Long = 16777215     ' RGB(255, 255, 255)
Public Const CLR_SIDEBAR As Long = 6044187    ' RGB(27, 58, 92)
'@ -split "`r?`n"

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('designtest-' + [guid]::NewGuid().ToString('N'))
$snap = Join-Path $root 'main'; $out = Join-Path $root 'designs'
New-Item -ItemType Directory -Force $snap | Out-Null
try {
    $themeFile = Join-Path $root 'modTheme.bas'
    [IO.File]::WriteAllLines($themeFile, [string[]]$theme)
    $tokens = Read-DesignTokens $themeFile
    Check 'theme tokens read' ($tokens.Tokens.Count -eq 4 -and $tokens.Tokens['CLR_SIDEBAR'] -eq '#1B3A5C' -and $tokens.ByColour[[long]16711679] -eq 'CLR_BG')

    $d = ConvertTo-FormDesign -Objects (ConvertFrom-FormText -Lines $form) -Name 'frmItem' -Tokens $tokens
    Check 'form fields' ($d.schema -eq 'access-form-design/1' -and $d.name -eq 'frmItem' -and $d.kind -eq 'Form' -and $d.width -eq 480)
    Check 'doubled quotes undone' ($d.properties['Caption'] -eq 'Edit "item"')
    Check 'continued string joined' ($d.properties['RecordSource'] -eq 'SELECT id, title FROM items WHERE active = True')
    Check 'NotDefault kept as saved' ($d.properties['PopUp'] -eq 'NotDefault')
    Check 'form events apart' ($d.events['OnLoad'] -eq '[Event Procedure]' -and -not $d.properties.Contains('OnLoad'))
    Check 'design window dropped' (-not $d.properties.Contains('Right') -and -not $d.properties.Contains('Left'))
    Check 'sections in order' ((@($d.sections | ForEach-Object { $_.name }) -join ',') -eq 'FormHeader,Detail')
    Check 'section height and token colour' ($d.sections[0].height -eq 60 -and $d.sections[0].properties['BackColor'].hex -eq '#1B3A5C' -and $d.sections[0].properties['BackColor'].token -eq 'CLR_SIDEBAR')
    Check 'section has no left or width' (-not $d.sections[1].Contains('left') -and -not $d.sections[1].Contains('width'))

    $h = $d.sections[0].controls[0]
    Check 'label position in px' ($h.name -eq 'lblHeading' -and $h.type -eq 'Label' -and $h.left -eq 16 -and $h.top -eq 16 -and $h.width -eq 200 -and $h.height -eq 28)
    Check 'own colour beats default' ($h.properties['ForeColor'].token -eq 'CLR_WHITE')
    Check 'type defaults merged' ($h.properties['BackStyle'] -eq 0 -and $h.properties['FontName'] -eq 'Segoe UI')
    Check 'position not repeated in properties' (-not $h.properties.Contains('Left'))

    $detail = $d.sections[1].controls
    Check 'detail controls in order' ((@($detail | ForEach-Object { $_.name }) -join ',') -eq 'txtTitle,tabMain,imgPill')
    $t = $detail[0]
    Check 'control events apart' ($t.events['OnClick'] -eq '[Event Procedure]' -and $t.events['AfterUpdate'] -eq '=SaveTitle()' -and -not $t.properties.Contains('OnClick'))
    Check 'system colour kept' ($t.properties['BorderColor'].system -eq -2147483632)
    Check 'defaults only for their type' (-not $t.properties.Contains('FontName'))
    $tab = $detail[1]
    Check 'tab size in px' ($tab.properties['TabFixedHeight'] -eq 32)
    Check 'page under tab, label under page' ($tab.children[0].name -eq 'pgOne' -and $tab.children[0].children[0].name -eq 'lblOnPage' -and $tab.children[0].children[0].top -eq 120)
    Check 'picture as hash' ($detail[2].properties['PictureData'] -match '^<binary, 1 lines, [0-9a-f]{12}>$')
    Check 'volatile properties dropped' (-not $d.sections[1].properties.Contains('GUID'))

    $plain = ConvertTo-FormDesign -Objects (ConvertFrom-FormText -Lines $form) -Name 'frmItem'
    Check 'no theme: hex without token' ($plain.sections[0].properties['BackColor'].hex -eq '#1B3A5C' -and -not $plain.sections[0].properties['BackColor'].Contains('token'))

    # A snapshot folder as AccessReview.ps1 writes it: UTF-16 <form>.txt files and
    # markers for forms it couldn't save.
    [IO.File]::WriteAllLines((Join-Path $snap 'frmItem.txt'), [string[]]$form, [Text.Encoding]::Unicode)
    [IO.File]::WriteAllText((Join-Path $snap 'frmBroken.failed'), 'SaveAsText failed')
    [IO.File]::WriteAllText((Join-Path $snap 'unsaved-names.list'), "frm:Odd`r`n")
    New-Item -ItemType Directory -Force $out | Out-Null
    [IO.File]::WriteAllText((Join-Path $out 'frmGone.json'), '{}')
    $r = Export-FormDesignFolder -SnapshotDir $snap -OutDir $out -ThemeFile $themeFile
    Check 'forms written' (($r.Written -join ',') -eq 'frmItem' -and (Test-Path (Join-Path $out 'frmItem.json')))
    Check 'old design removed' (-not (Test-Path (Join-Path $out 'frmGone.json')))
    $j = Get-Content -Raw (Join-Path $out 'frmItem.json') | ConvertFrom-Json
    Check 'json reads back' ($j.sections[1].controls[1].children[0].children[0].name -eq 'lblOnPage' -and $j.sections[0].properties.BackColor.token -eq 'CLR_SIDEBAR')
    Check 'json keeps one-item lists as lists' ($j.sections[0].controls -is [array] -and $j.sections[0].controls.Count -eq 1)
    $ix = Get-Content -Raw (Join-Path $out 'index.json') | ConvertFrom-Json
    Check 'index counts controls' ($ix.forms[0].form -eq 'frmItem' -and $ix.forms[0].sections -eq 2 -and $ix.forms[0].controls -eq 6 -and $ix.snapshot -eq 'main' -and $ix.theme -eq 'modTheme.bas')
    Check 'index lists failed snapshots' ((@($ix.failed | ForEach-Object { $_.form }) -join ',') -eq 'frmBroken,frm:Odd')
    $tk = Get-Content -Raw (Join-Path $out 'tokens.json') | ConvertFrom-Json
    Check 'tokens file' ($tk.tokens.CLR_BG -eq '#FFFFFE' -and $tk.tokens.CLR_FG -eq '#0F172A')
} finally { Remove-Item -Recurse -Force $root }

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
