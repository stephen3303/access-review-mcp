# Checks OfficeCom.ps1 (member paths, safety rules, value conversion) and that the
# Office tools' C# helper compiles, without Office:
#   powershell -NoProfile -File tools\review-mcp\OfficeCom.Tests.ps1
# (pwsh works too, on any OS). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeCom.ps1')

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}
function Throws([scriptblock]$b, [string]$like = '*') {
    try { & $b; $false } catch { $_.Exception.Message -like $like }
}

# --- member paths ---
$p = ConvertFrom-OfficePath "Worksheets('Summary').Range(`"A1:C4`").Value2"
Check 'three segments' ($p.Count -eq 3 -and $p[0].Name -eq 'Worksheets' -and $p[2].Name -eq 'Value2')
Check 'single-quoted argument' ($p[0].Args[0] -eq 'Summary' -and $p[0].HasArgs)
Check 'double-quoted argument' ($p[1].Args[0] -eq 'A1:C4')
Check 'member without brackets has no arguments' (-not $p[2].HasArgs -and $p[2].Args.Count -eq 0)

$p = ConvertFrom-OfficePath "Cells(3, -2).Offset(1.5, 2e3)"
Check 'integers are Int32' ($p[0].Args[0] -is [int] -and $p[0].Args[0] -eq 3 -and $p[0].Args[1] -eq -2)
Check 'decimals and exponents are Double' ($p[1].Args[0] -is [double] -and $p[1].Args[0] -eq 1.5 -and $p[1].Args[1] -eq 2000)

Check 'comma after a member is refused' (Throws { ConvertFrom-OfficePath "Range('a'), Foo" } '*expected .*')
$p = ConvertFrom-OfficePath "Range('It''s ""ok""')"
Check 'doubled quote is one quote' ($p[0].Args[0] -eq "It's `"ok`"")

$p = ConvertFrom-OfficePath "SaveAs2('C:\Review\a.docx', FileFormat:=16, AddToRecentFiles:=false)"
Check 'named arguments come first, then positional' ($p[0].Names.Count -eq 2 -and $p[0].Names[0] -eq 'FileFormat' -and $p[0].Args[0] -eq 16 -and $p[0].Args[1] -eq $false -and $p[0].Args[2] -eq 'C:\Review\a.docx')
Check 'positional after named is refused' (Throws { ConvertFrom-OfficePath "Foo(A:=1, 2)" } '*positional*')
Check 'named argument twice is refused' (Throws { ConvertFrom-OfficePath "Foo(A:=1, A:=2)" } '*twice*')

$p = ConvertFrom-OfficePath "Foo(, 2, missing, null, TRUE, @h12)"
Check 'empty slot is missing' ($p[0].Args[0].OfficeMissing)
Check 'missing keyword' ($p[0].Args[2].OfficeMissing)
Check 'null' ($null -eq $p[0].Args[3])
Check 'true is case-insensitive' ($p[0].Args[4] -eq $true)
Check 'handle reference' ($p[0].Args[5].OfficeHandle -eq 'h12')
Check 'empty brackets' ((ConvertFrom-OfficePath 'Add()')[0].HasArgs -and (ConvertFrom-OfficePath 'Add()')[0].Args.Count -eq 0)
Check 'empty path is no segments' (@(ConvertFrom-OfficePath '').Count -eq 0)
Check 'unclosed text refused' (Throws { ConvertFrom-OfficePath "Range('A1" } '*unclosed*')
Check 'missing ) refused' (Throws { ConvertFrom-OfficePath "Range('A1'" } '*')
Check 'expression in an argument refused' (Throws { ConvertFrom-OfficePath "Range(1+2)" } '*')
Check 'member name with odd characters refused' (Throws { ConvertFrom-OfficePath "Range;Shell" } '*')
Check 'leading dot refused' (Throws { ConvertFrom-OfficePath ".Value" } '*member name*')

# --- refused members ---
Check 'Run refused everywhere' ([bool](Get-OfficeMemberRefusal 'excel' 'Run'))
Check 'refusal ignores case' ([bool](Get-OfficeMemberRefusal 'excel' 'executeexcel4macro'))
foreach ($m in 'Evaluate', 'DDEInitiate', 'Shell', 'SendKeys', 'OnTime', 'VBE', 'VBProject', 'Send', 'SendMail', 'PrintOut', 'Quit', 'FollowHyperlink', 'ExecuteMso', 'WordBasic', 'RegisterXLL', 'Eval', 'AutomationSecurity') {
    Check "$m refused" ([bool](Get-OfficeMemberRefusal 'word' $m))
}
Check 'Outlook Delete refused' ([bool](Get-OfficeMemberRefusal 'outlook' 'Delete'))
Check 'Excel Delete allowed' (-not (Get-OfficeMemberRefusal 'excel' 'Delete'))
Check 'Access RunSQL refused' ([bool](Get-OfficeMemberRefusal 'access' 'RunSQL'))
Check 'ordinary members allowed' (-not (Get-OfficeMemberRefusal 'excel' 'Range') -and -not (Get-OfficeMemberRefusal 'word' 'Execute'))

# --- paths and values in arguments ---
$inAllowed = { param($p) $p -like 'C:\Review\*' }
Check 'full path inside allowed passes' (-not (Throws { Assert-OfficeCallAllowed -App excel -Member SaveAs -Arguments @('C:\Review\out.xlsx', 51) -InAllowed $inAllowed }))
Check 'full path outside allowed refused' (Throws { Assert-OfficeCallAllowed -App excel -Member SaveAs -Arguments @('C:\Windows\x.xlsx') -InAllowed $inAllowed } '*outside*')
Check 'UNC path outside allowed refused' (Throws { Assert-OfficeCallAllowed -App excel -Member Open -Arguments @('\\server\share\x.xlsx') -InAllowed $inAllowed } '*outside*')
Check 'web address refused' (Throws { Assert-OfficeCallAllowed -App word -Member Open -Arguments @('https://example.com/a.docx') -InAllowed $inAllowed } '*web*')
Check 'file: address refused' (Throws { Assert-OfficeCallAllowed -App word -Member InsertFile -Arguments @('file:///C:/x.docx') -InAllowed $inAllowed } '*')
Check '.. refused' (Throws { Assert-OfficeCallAllowed -App excel -Member SaveAs -Arguments @('C:\Review\..\Windows\x.xlsx') -InAllowed { $true } } "*'..'*")
Check 'file method with a bare name refused' (Throws { Assert-OfficeCallAllowed -App excel -Member SaveAs -Arguments @('report.xlsx') -InAllowed $inAllowed } '*full path*')
Check 'path inside a nested array checked' (Throws { Assert-OfficeCallAllowed -App excel -Member Foo -Arguments @(, @(@('a', 'D:\x'))) -InAllowed $inAllowed } '*outside*')
Check 'path on a property set checked' (Throws { Assert-OfficeCallAllowed -App excel -Member DefaultFilePath -Arguments @('C:\Windows') -InAllowed $inAllowed } '*outside*')
Check 'plain text passes' (-not (Throws { Assert-OfficeCallAllowed -App word -Member Text -Arguments @('Status: 3 overdue (see C drive)') -InAllowed $inAllowed }))
Check 'ordinary formula passes' (-not (Throws { Assert-OfficeCallAllowed -App excel -Member Formula -Arguments @('=SUM(A1:A3)') -InAllowed $inAllowed }))
Check 'negative number text passes' (-not (Get-OfficeValueRefusal '-5'))
Check 'DDE formula refused' (Throws { Assert-OfficeCallAllowed -App excel -Member Formula -Arguments @("=cmd|' /c calc'!A0") -InAllowed $inAllowed } '*DDE*')
Check 'DDE with + refused' ([bool](Get-OfficeValueRefusal "+cmd|'/c calc'!A0"))
Check 'CALL refused' ([bool](Get-OfficeValueRefusal '=CALL("kernel32","WinExec","JCJ","calc",0)'))
Check 'WEBSERVICE refused' ([bool](Get-OfficeValueRefusal '=WEBSERVICE("http://x")'))
Check 'RTD refused, any case' ([bool](Get-OfficeValueRefusal '=rtd("x",,"y")'))
Check 'refused member stops before values' (Throws { Assert-OfficeCallAllowed -App excel -Member Run -Arguments @() -InAllowed $inAllowed } '*not allowed*')

# --- JSON to COM values ---
$v = ConvertTo-OfficeComValue ([long]5) { }
Check 'Int64 becomes Int32' ($v -is [int] -and $v -eq 5)
Check 'large Int64 becomes Double' ((ConvertTo-OfficeComValue ([long]5000000000) { }) -is [double])
Check 'decimal becomes Double' ((ConvertTo-OfficeComValue ([decimal]1.25) { }) -is [double])
$json = '{"v":[[1,"a",true],[2,"b",null]]}' | ConvertFrom-Json
$g = ConvertTo-OfficeComValue $json.v { }
Check 'rows become a 2-D array' ($g -is [object[,]] -and $g.GetLength(0) -eq 2 -and $g.GetLength(1) -eq 3)
Check '2-D cells keep values' ($g[0, 1] -eq 'a' -and $g[1, 0] -eq 2 -and $g[1, 0] -is [int] -and $null -eq $g[1, 2])
Check 'ragged rows refused' (Throws { ConvertTo-OfficeComValue (('[[1,2],[3]]' | ConvertFrom-Json)) { } } '*same length*')
$one = ConvertTo-OfficeComValue (('[1,2,3]' | ConvertFrom-Json)) { }
Check 'flat array stays 1-D' ($one -is [object[]] -and $one.Count -eq 3 -and $one[2] -eq 3)
$h = ConvertTo-OfficeComValue (('{"handle":"h4"}' | ConvertFrom-Json)) { param($id) "object $id" }
Check 'handle object resolved' ($h -eq 'object h4')
Check 'other objects refused' (Throws { ConvertTo-OfficeComValue (('{"a":1}' | ConvertFrom-Json)) { } } '*Unsupported*')

# --- COM to JSON values ---
$a = [Array]::CreateInstance([object], [int[]]@(2, 2), [int[]]@(1, 1))   # Excel's 1-based block
$a.SetValue(1, 1, 1); $a.SetValue('x', 1, 2); $a.SetValue([DBNull]::Value, 2, 1); $a.SetValue([datetime]'2026-11-06T08:30:00', 2, 2)
$r = ConvertFrom-OfficeComValue $a $null
Check '1-based block becomes rows' ($r.RowCount -eq 2 -and $r.ColumnCount -eq 2 -and $r.Rows[0][1] -eq 'x' -and $r.Rows[0][0] -eq 1)
Check 'DBNull becomes null, dates ISO' ($null -eq $r.Rows[1][0] -and $r.Rows[1][1] -eq '2026-11-06T08:30:00')
$big = New-Object 'object[,]' 100, 10
$r = ConvertFrom-OfficeComValue $big $null 250
Check 'large block truncated by whole rows' ($r.Truncated -and $r.Rows.Count -eq 25 -and $r.RowCount -eq 100)
Check 'scalars pass through' ((ConvertFrom-OfficeComValue 3.5 $null) -eq 3.5 -and (ConvertFrom-OfficeComValue 'a' $null) -eq 'a' -and $null -eq (ConvertFrom-OfficeComValue ([Reflection.Missing]::Value) $null))
$l = ConvertFrom-OfficeComValue ([object[]]@(1, 'b')) $null
Check '1-D array becomes a list' ($l.Count -eq 2 -and $l[1] -eq 'b')

# --- audit line ---
$line = Format-OfficeAudit -Tool office_invoke -App excel -Target h2 -Path 'Range("A1").Value2' -Arguments ([pscustomobject]@{ value = ('x' * 1000) }) -Ok $false -Failure 'nope'
$o = $line | ConvertFrom-Json
Check 'audit line is one JSON object' ($o.tool -eq 'office_invoke' -and $o.app -eq 'excel' -and $o.ok -eq $false -and $o.error -eq 'nope')
Check 'audit arguments preview cut to 300' ($o.args.Length -le 303)

# --- the C# type-information helper compiles (PowerShell 5.1 builds it with the C# 5 compiler) ---
$compiled = $false
try {
    $src = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'OfficeReview.ps1') -Raw
    $cs = [regex]::Match($src, "(?s)Add-Type @'\r?\n(.*?)\r?\n'@").Groups[1].Value
    if (-not ('KitOfficeCom' -as [type])) { Add-Type -TypeDefinition $cs }
    $compiled = [bool]('KitOfficeCom' -as [type])
} catch { Write-Host $_.Exception.Message }
Check 'KitOfficeCom compiles' $compiled
if ($compiled) { Check 'TypeName of a plain object is null' ($null -eq [KitOfficeCom]::TypeName('text')) }

Write-Host ''
if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'all checks passed'
