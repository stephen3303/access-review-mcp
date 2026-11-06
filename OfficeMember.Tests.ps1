# Checks the member-path walk of OfficeReview.ps1 (office_invoke) without Office: plain
# .NET objects stand in for an object model and a stand-in dispatcher records each call.
#   powershell -NoProfile -File tools\review-mcp\OfficeMember.Tests.ps1
# (pwsh works too, on any OS). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'

# The few AccessReview.ps1 helpers OfficeReview.ps1 calls, when it isn't loaded.
if (-not (Get-Command Get-ReviewSetting -ErrorAction SilentlyContinue)) {
    function Get-ReviewSetting($Name) { $null }
    function Get-ReviewOutputBase { Join-Path ([IO.Path]::GetTempPath()) 'office-member-tests' }
    function Test-FullPath($Path) { $Path -match '^[A-Za-z]:[\\/]' }
    function Test-PathInFolder($Path, $Folder) { $Path.StartsWith($Folder.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) }
    function Release-Com($o) {}
}
. (Join-Path $PSScriptRoot 'OfficeReview.ps1')

$failures = 0
function Check([string]$what, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what $detail"; $script:failures++ }
}

if (-not ('FakeOfficeBook' -as [type])) {
    Add-Type @'
public class FakeOfficeRange : System.Collections.IEnumerable {
    public object Value2 = 7;
    // A real Range enumerates its cells; PowerShell unrolls such output unless kept whole.
    public System.Collections.IEnumerator GetEnumerator() { return new object[] { 1, 2, 3 }.GetEnumerator(); }
    public string Address(bool rowAbs, bool colAbs) { return "A1:C3"; }
}
public class FakeOfficeSheet {
    public string Name;
    public FakeOfficeRange UsedRange = new FakeOfficeRange();
    public string LastRange;
    public FakeOfficeRange Range(string a) { LastRange = a; return new FakeOfficeRange(); }
}
public class FakeOfficeSheets {
    public int Count = 2;
    public FakeOfficeSheet Item(int i) { return new FakeOfficeSheet { Name = "S" + i }; }
}
public class FakeOfficeDoc {
    public string Text = "Review heading";
    public string Span(ref int start, ref int end) { return Text.Substring(start, end - start); }
}
public class FakeOfficeBook {
    public FakeOfficeSheets Worksheets = new FakeOfficeSheets();
    public string Name = "Book1";
    public string Received;
    public void Take(object v) {
        var g = v as object[,];
        Received = v == null ? "null" : g != null ? "grid " + g.GetLength(0) + "x" + g.GetLength(1) + " " + g[1, 0].GetType().Name : v.GetType().Name;
    }
}
'@
}

# Stand-ins: plain objects count as COM objects, and the dispatcher reads fields and
# calls methods by reflection, sending a field given arguments to its Item as the
# adapter path does.
function Test-OfficeObject($v) { $null -ne $v -and $v.GetType().Name -like 'FakeOffice*' }
$script:Calls = New-Object Collections.Generic.List[string]
function Invoke-OfficeDispatch($Obj, [string]$Name, [Reflection.BindingFlags]$Flags, [object[]]$Argv, [string[]]$Names) {
    $o = Unwrap $Obj
    $plain = New-Object object[] $Argv.Count
    for ($k = 0; $k -lt $Argv.Count; $k++) { $plain[$k] = Unwrap $Argv[$k] }
    $script:Calls.Add("$Name/$($plain.Count)")
    $f = $o.GetType().GetField($Name)
    if ($f) {
        if ($Flags -band [Reflection.BindingFlags]::SetProperty) { $f.SetValue($o, $plain[$plain.Count - 1]); return }
        $v = $f.GetValue($o)
        if ($plain.Count) { return , (Invoke-OfficeDispatch $v 'Item' $Flags $plain $null) }
        return , $v
    }
    $m = $o.GetType().GetMethod($Name)
    if (-not $m) { throw "Unknown name $Name" }
    , $m.Invoke($o, $plain)
}
$script:OfficeSessions['excel'] = [pscustomobject]@{ App = 'excel'; Label = 'Excel'; Com = $null; Pid = 0; Process = $null; Started = $true; Mode = 'hidden'; Dialogs = @() }

$book = New-Object FakeOfficeBook
$id = Add-OfficeHandle $book 'excel'
Check 'handle for an object' ($id -eq 'h1')

$script:Calls.Clear()
$r = Invoke-OfficeMember -Target $id -Path "Worksheets(1).Range('A1').Value2"
Check 'path walks one member at a time' (($script:Calls -join ' ') -eq 'Worksheets/1 Item/1 Range/1 Value2/0') ($script:Calls -join ' ')
Check 'value at the end of the path' ($r.Value -eq 7)

$r = Invoke-OfficeMember -Target $id -Path "Worksheets(1).Range('A3:B3')"
Check 'enumerable result stays one object (a handle)' ($r.Value.Handle -match '^h\d+$') ($r | ConvertTo-Json -Compress -Depth 4)
$r = Invoke-OfficeMember -Target $id -Path "Worksheets(1).UsedRange.Address(false, false)"
Check 'member read through an enumerable object' ($r.Value -eq 'A1:C3') ($r | ConvertTo-Json -Compress -Depth 4)

$r = Invoke-OfficeMember -Target $id -Path 'Worksheets(2)' -Fields @('Name', 'UsedRange.Address(false, false)')
Check 'object result comes back as a handle' ($r.Value.Handle -match '^h\d+$')
Check 'fields read from the result' ($r.Fields.Name -eq 'S2' -and $r.Fields.'UsedRange.Address(false, false)' -eq 'A1:C3') ($r.Fields | ConvertTo-Json -Compress)
$sheet = $r.Value.Handle
Check 'handle works as a target' ((Invoke-OfficeMember -Target $sheet -Path 'Name').Value -eq 'S2')
Check 'same object, same handle' ((Invoke-OfficeMember -Target $id -Path '').Value.Handle -eq $id)

$r = Invoke-OfficeMember -Target $id -Path 'Worksheets' -Enumerate
Check 'collection enumerated with names' ($r.Count -eq 2 -and (@($r.Items | ForEach-Object { $_.Fields.Name }) -join ',') -eq 'S1,S2')

$rows = ('{"a":[[1,2],[3,4]]}' | ConvertFrom-Json).a
$null = Invoke-OfficeMember -Target $id -Path 'Take' -Arguments @(, $rows)
Check 'JSON rows reach the call as one 2-D array' ($book.Received -eq 'grid 2x2 Int32') $book.Received
$null = Invoke-OfficeMember -Target $id -Path 'Name' -Set -Value 'Renamed'
Check 'set a property' ($book.Name -eq 'Renamed')

$script:Calls.Clear()
$refused = $false
try { $null = Invoke-OfficeMember -Target $id -Path "Worksheets(1).Range('A1').ExecuteExcel4Macro('x')" } catch { $refused = $_.Exception.Message -like '*not allowed*' }
Check 'refused member stops the walk before any call' ($refused -and $script:Calls.Count -eq 0) ($script:Calls -join ' ')
$refused = $false
try { $null = Invoke-OfficeMember -Target $id -Path 'Take' -Arguments @('=cmd|''/c calc''!A0') } catch { $refused = $_.Exception.Message -like '*DDE*' }
Check 'refused argument stops the call' ($refused -and $book.Received -eq 'grid 2x2 Int32')

# By-reference parameters (Word's Range(Start, End)): the adapter refuses plain values,
# so the call is retried through InvokeMember.
$doc = New-Object FakeOfficeDoc
$span = $doc.PSObject.Members.Match('Span') | Select-Object -First 1
$adapterRefuses = $false
try { $null = $span.Invoke([object[]]@(0, 6)) } catch { $adapterRefuses = $_.Exception.Message -like '*PSReference*' }
Check 'adapter refuses by-reference arguments (the case being covered)' $adapterRefuses
$got = $null
try { $got = Invoke-OfficeAdapterCall $span $doc 'Span' ([Reflection.BindingFlags]::InvokeMethod) ([object[]]@(0, 6)) } catch { $got = $_.Exception.Message }
Check 'by-reference arguments retried through InvokeMember' ($got -eq 'Review') "$got"

Write-Host ''
if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'all checks passed'
