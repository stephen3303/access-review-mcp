# Office COM tools: drive Excel, Word, PowerPoint, Outlook, Visio, Project, Publisher
# (and the Access instance open_app started) through their object models, without the
# desktop's mouse, keyboard or screen.
#
# Two layers:
#   - a generic layer that reaches any member of any object model: office_invoke walks a
#     VBA-like member path (Worksheets('Summary').Range("A1").Value2) and gets, sets or
#     calls it; office_members lists what an object offers, read from its type library,
#     so nothing has to be written by hand for each method;
#   - typed helpers for the common jobs: start and quit an app, open, close, save and
#     export documents, capture a window, run a document's macro, read or write cells,
#     read a document's paragraphs, list slides and shapes, list mail items.
#
# Dot-sourced by AccessReviewMcp.ps1 after AccessReview.ps1 (it uses KitReviewWin,
# KitReviewKeeper, KitReviewGuard and the setting, output and allowed-folder helpers
# there), and usable directly:
#
#   . "$PSScriptRoot\AccessReview.ps1"; . "$PSScriptRoot\OfficeReview.ps1"
#   Start-OfficeApp excel
#   $wb = Open-OfficeDocument excel 'C:\Review\Budget.xlsx'
#   Invoke-OfficeMember -Target $wb.Handle -Path "Worksheets(1).Range('A1:C3').Value2"
#   Save-OfficeCapture -Target excel
#   Stop-OfficeApp excel
#
# Safety rules every call keeps (see OfficeCom.ps1 for the lists):
#   - only the apps in ACCESS_REVIEW_OFFICE_APPS (Outlook is off unless listed);
#   - an app this tool starts is its own instance, parked off the desktop; a program
#     that keeps one instance per user (PowerPoint, Outlook, Project, Publisher) is not
#     attached to while the user has it open, unless ACCESS_REVIEW_OFFICE_ATTACH lists it;
#   - documents open from the allowed folders only, never from the internet (Mark of the
#     Web), read-only by default and with macros disabled unless asked for;
#   - members that run caller-supplied code, start programs, send mail or input, print
#     or quit are refused; Outlook items are read only;
#   - any path or address in an argument must be a full path inside an allowed folder
#     or the output folder; formulas that reach code or the web (DDE, CALL, RTD,
#     WEBSERVICE...) are refused;
#   - every call is written to office-audit.log in the output folder;
#   - message boxes during a call are captured and closed, and a call that hangs ends the
#     app (never an instance this tool attached to).

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'OfficeCom.ps1')

if (-not ('KitOfficeCom' -as [type])) {
    Add-Type @'
using System; using System.Collections.Generic; using System.Reflection; using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

[ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IKitDispatchInfo {
    [PreserveSig] int GetTypeInfoCount(out uint count);
    [PreserveSig] int GetTypeInfo(uint index, int lcid, out ComTypes.ITypeInfo info);
}

// Reads what a COM object offers from its type information (IDispatch.GetTypeInfo), so
// the tools can describe any object model without a hand-written list.
public static class KitOfficeCom {
    // The plain value behind a PowerShell wrapper, inside arrays too: COM would get a
    // wrapper as an IDispatch of its own, and the app would fail to read it.
    static object Base(object o) {
        var p = o as System.Management.Automation.PSObject;
        if (p != null) o = p.BaseObject;
        var a = o as Array;
        if (a != null && a.Rank == 2) {
            for (int i = a.GetLowerBound(0); i <= a.GetUpperBound(0); i++)
                for (int j = a.GetLowerBound(1); j <= a.GetUpperBound(1); j++) a.SetValue(Base(a.GetValue(i, j)), i, j);
        } else if (a != null && a.Rank == 1 && a.GetType().GetElementType() == typeof(object)) {
            for (int i = a.GetLowerBound(0); i <= a.GetUpperBound(0); i++) a.SetValue(Base(a.GetValue(i)), i);
        }
        return o;
    }

    // IDispatch call made here rather than from PowerShell, so PowerShell's method binder
    // never rewrites the arguments. Named arguments come first in args.
    public static object Invoke(object target, string name, int flags, object[] args, string[] names) {
        target = Base(target);
        var a = new object[args == null ? 0 : args.Length];
        for (int i = 0; i < a.Length; i++) a[i] = Base(args[i]);
        try {
            return target.GetType().InvokeMember(name, (BindingFlags)flags, null, target, a, null, null, (names != null && names.Length > 0) ? names : null);
        } catch (TargetInvocationException e) {
            if (e.InnerException != null) ExceptionDispatchInfo.Capture(e.InnerException).Throw();
            throw;
        }
    }

    static ComTypes.ITypeInfo Info(object o) {
        var d = o as IKitDispatchInfo;
        if (d == null) return null;
        uint c;
        if (d.GetTypeInfoCount(out c) != 0 || c == 0) return null;
        ComTypes.ITypeInfo ti;
        if (d.GetTypeInfo(0, 0, out ti) != 0) return null;
        return ti;
    }

    static string Doc(ComTypes.ITypeInfo ti, int memid) {
        string name, doc, file; int ctx;
        ti.GetDocumentation(memid, out name, out doc, out ctx, out file);
        return memid == -1 ? name : doc;
    }

    // The object's type name, e.g. Workbook (Office names the interface _Workbook).
    public static string TypeName(object o) {
        try {
            var ti = Info(o);
            if (ti == null) return null;
            var n = Doc(ti, -1);
            return n != null && n.StartsWith("_") ? n.Substring(1) : n;
        } catch { return null; }
    }

    // One row per member: name, kind (method, get, put, putref, field), parameters
    // (optional ones end in ?), description. Restricted and hidden members are left out.
    public static List<string[]> Members(object o) {
        var rows = new List<string[]>();
        var ti = Info(o);
        if (ti == null) return rows;
        IntPtr pa; ti.GetTypeAttr(out pa);
        var ta = (ComTypes.TYPEATTR)Marshal.PtrToStructure(pa, typeof(ComTypes.TYPEATTR));
        ti.ReleaseTypeAttr(pa);
        int elemSize = Marshal.SizeOf(typeof(ComTypes.ELEMDESC));
        for (int i = 0; i < ta.cFuncs; i++) {
            IntPtr pf; ti.GetFuncDesc(i, out pf);
            try {
                var fd = (ComTypes.FUNCDESC)Marshal.PtrToStructure(pf, typeof(ComTypes.FUNCDESC));
                if ((fd.wFuncFlags & 0x41) != 0) continue;   // FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN
                var names = new string[fd.cParams + 1]; int cn;
                ti.GetNames(fd.memid, names, names.Length, out cn);
                var ps = new List<string>();
                for (int k = 1; k < cn; k++) {
                    var ed = (ComTypes.ELEMDESC)Marshal.PtrToStructure(new IntPtr(fd.lprgelemdescParam.ToInt64() + (k - 1) * elemSize), typeof(ComTypes.ELEMDESC));
                    bool opt = (((int)ed.desc.paramdesc.wParamFlags) & 0x10) != 0 || k > fd.cParams - fd.cParamsOpt;
                    ps.Add(names[k] + (opt ? "?" : ""));
                }
                string kind = fd.invkind == ComTypes.INVOKEKIND.INVOKE_FUNC ? "method"
                    : fd.invkind == ComTypes.INVOKEKIND.INVOKE_PROPERTYGET ? "get"
                    : fd.invkind == ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT ? "put" : "putref";
                string d = ""; try { d = Doc(ti, fd.memid) ?? ""; } catch { }
                rows.Add(new string[] { names[0], kind, string.Join(", ", ps), d });
            } finally { ti.ReleaseFuncDesc(pf); }
        }
        for (int i = 0; i < ta.cVars; i++) {
            IntPtr pv; ti.GetVarDesc(i, out pv);
            try {
                var vd = (ComTypes.VARDESC)Marshal.PtrToStructure(pv, typeof(ComTypes.VARDESC));
                var names = new string[1]; int cn;
                ti.GetNames(vd.memid, names, 1, out cn);
                string d = ""; try { d = Doc(ti, vd.memid) ?? ""; } catch { }
                rows.Add(new string[] { names[0], "field", "", d });
            } finally { ti.ReleaseVarDesc(pv); }
        }
        return rows;
    }
}
'@
}

# The apps these tools know. Single: one instance per user, so a running copy is the
# user's own. Access is driven through open_app's instance, OneNote through onenote_*.
$script:OfficeAppDefs = [ordered]@{
    excel      = @{ ProgId = 'Excel.Application';      Process = 'EXCEL';    Single = $false; Label = 'Excel' }
    word       = @{ ProgId = 'Word.Application';       Process = 'WINWORD';  Single = $false; Label = 'Word' }
    powerpoint = @{ ProgId = 'PowerPoint.Application'; Process = 'POWERPNT'; Single = $true;  Label = 'PowerPoint' }
    outlook    = @{ ProgId = 'Outlook.Application';    Process = 'OUTLOOK';  Single = $true;  Label = 'Outlook' }
    visio      = @{ ProgId = 'Visio.Application';      Process = 'VISIO';    Single = $false; Label = 'Visio' }
    project    = @{ ProgId = 'MSProject.Application';  Process = 'WINPROJ';  Single = $true;  Label = 'Project' }
    publisher  = @{ ProgId = 'Publisher.Application';  Process = 'MSPUB';    Single = $true;  Label = 'Publisher' }
    access     = @{ ProgId = 'Access.Application';     Process = 'MSACCESS'; Single = $false; Label = 'Access' }
}
$script:OfficeDefaultApps = 'excel;word;powerpoint;visio;project;publisher;access'

$script:OfficeSessions = @{}
$script:OfficeHandles = @{}
$script:OfficeNextHandle = 0
$script:OfficeOut = $null

function Get-OfficeSettingList([string]$Name, [string]$Default) {
    $v = Get-ReviewSetting $Name
    if (-not $v) { $v = $Default }
    @(($v -split ';') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
}

function Get-OfficeAllowedApps { Get-OfficeSettingList 'ACCESS_REVIEW_OFFICE_APPS' $script:OfficeDefaultApps }

function Get-OfficeOut {
    if (-not $script:OfficeOut) {
        $script:OfficeOut = Join-Path (Get-ReviewOutputBase) ('office-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType Directory -Force -Path $script:OfficeOut | Out-Null
    }
    $script:OfficeOut
}

function Write-OfficeAudit([string]$Tool, $Arguments, [bool]$Ok, [string]$Failure) {
    try {
        $a = $Arguments
        $app = $null; $target = $null; $path = $null
        if ($a) {
            if ($a.PSObject.Properties['app']) { $app = [string]$a.app }
            foreach ($k in 'target', 'document') { if ($a.PSObject.Properties[$k]) { $target = [string]$a.$k } }
            if ($a.PSObject.Properties['path']) { $path = [string]$a.path }
        }
        $base = Get-ReviewOutputBase
        New-Item -ItemType Directory -Force -Path $base | Out-Null
        $line = Format-OfficeAudit -Tool $Tool -App $app -Target $target -Path $path -Arguments $a -Ok $Ok -Failure $Failure
        [IO.File]::AppendAllText((Join-Path $base 'office-audit.log'), $line + "`r`n", (New-Object Text.UTF8Encoding($false)))
    } catch {}
}

# True when a full path is inside an allowed folder (ACCESS_REVIEW_ALLOWED_FOLDERS) or
# the output folder. With no allowed folders set, only the output folder.
function Test-OfficePathAllowed([string]$Path) {
    if (-not (Test-OfficeFullPath $Path)) { return $false }
    $full = [IO.Path]::GetFullPath($Path)
    $roots = @(((Get-ReviewSetting 'ACCESS_REVIEW_ALLOWED_FOLDERS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and (Test-FullPath $_) })
    $roots += Get-ReviewOutputBase
    [bool]@($roots | Where-Object { $full.TrimEnd('\') -ieq ([IO.Path]::GetFullPath($_).TrimEnd('\')) -or (Test-PathInFolder $full $_) }).Count
}

# Refuses a path that goes through a junction or symbolic link below the allowed or
# output folder it is in (the name check alone could then lead outside), and a file
# downloaded from the internet (Mark of the Web, zone 3 or 4).
function Assert-OfficeFileSafe([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $roots = @(((Get-ReviewSetting 'ACCESS_REVIEW_ALLOWED_FOLDERS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and (Test-FullPath $_) })
    $roots += Get-ReviewOutputBase
    foreach ($r in $roots) {
        $root = [IO.Path]::GetFullPath($r).TrimEnd('\')
        if (-not ($full -ieq $root -or (Test-PathInFolder $full $root))) { continue }
        for ($p = $full; $p -and $p.Length -gt $root.Length; $p = [IO.Path]::GetDirectoryName($p)) {
            $link = (Get-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue).LinkType
            if ($link -in 'Junction', 'SymbolicLink') { throw "'$Path' goes through the $link '$p', which may lead outside $root. Use the real location." }
        }
    }
    if (Test-Path -LiteralPath $full -PathType Leaf) {
        $zone = $null
        try { $zone = Get-Content -LiteralPath $full -Stream Zone.Identifier -ErrorAction Stop } catch {}
        if ($zone -and (($zone -join "`n") -match '(?m)^ZoneId=([34])\s*$')) {
            throw "'$Path' was downloaded from the internet (Zone.Identifier ZoneId=$($Matches[1])), so it is not opened. If you trust it, unblock it first (file Properties, Unblock)."
        }
    }
}

function Get-OfficeRoots {
    $r = @(((Get-ReviewSetting 'ACCESS_REVIEW_ALLOWED_FOLDERS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    ($r + @(Get-ReviewOutputBase)) -join '; '
}

function Get-OfficeDef([string]$App) {
    $App = ([string]$App).ToLowerInvariant()
    if (-not $script:OfficeAppDefs.Contains($App)) {
        if ($App -eq 'onenote') { throw 'OneNote is reached through the onenote_* tools, which keep to local notebooks.' }
        throw "Unknown app '$App'. Known: $(@($script:OfficeAppDefs.Keys) -join ', ')."
    }
    if ((Get-OfficeAllowedApps) -notcontains $App) { throw "$($script:OfficeAppDefs[$App].Label) is not enabled for these tools (ACCESS_REVIEW_OFFICE_APPS: $((Get-OfficeAllowedApps) -join ';'))." }
    $script:OfficeAppDefs[$App]
}

function Test-OfficeRegistered([string]$ProgId) {
    Test-Path -LiteralPath "Registry::HKEY_CLASSES_ROOT\$ProgId\CLSID"
}

# The plain .NET object behind a PowerShell wrapper: a wrapped value inside an argument
# array would reach COM as a PowerShell object (type mismatch). Arrays come back whole.
# True for a COM object (OfficeMember.Tests.ps1 replaces it to drive the walk with plain .NET objects).
function Test-OfficeObject($v) { $v -is [__ComObject] }

# Always returned with a leading comma: a function's output is unrolled, and a COM
# collection (a Range of many cells, a Sheets object) would come back as its items, or
# hang while Excel lists every cell of Cells.
function Unwrap($v) {
    if ($null -ne $v -and $v.PSObject) { $v = $v.PSObject.BaseObject }
    , $v
}

function Set-OfficeProp($obj, [string]$Name, $Value) {
    try { [void](Invoke-OfficeDispatch $obj $Name ([Reflection.BindingFlags]::SetProperty) @(, (Unwrap $Value)) $null); $true } catch { $false }
}

function Get-OfficeProp($obj, [string]$Name) {
    try { , (Invoke-OfficeDispatch $obj $Name ([Reflection.BindingFlags]::GetProperty) @() $null) } catch { $null }
}

# The innermost error message, with its HRESULT.
function Get-OfficeError($err) {
    $e = $err; if ($e -is [Management.Automation.ErrorRecord]) { $e = $e.Exception }
    while ($e.InnerException) { $e = $e.InnerException }
    $msg = $e.Message
    if ($e -is [Runtime.InteropServices.COMException]) { $msg += ' (0x{0:X8})' -f $e.HResult }
    elseif ($err -is [Management.Automation.ErrorRecord] -and $err.InvocationInfo -and $err.InvocationInfo.ScriptLineNumber) {
        $msg += " ($(Split-Path -Leaf $err.InvocationInfo.ScriptName) line $($err.InvocationInfo.ScriptLineNumber))"
    }
    $msg
}

function Get-OfficeSession([string]$App) {
    $s = $script:OfficeSessions[$App]
    if (-not $s) { return $null }
    if ($s.Process -and $s.Process.HasExited) {
        Clear-OfficeHandles $App
        $script:OfficeSessions.Remove($App)
        return $null
    }
    $s
}

function Assert-OfficeSession([string]$App) {
    $App = $App.ToLowerInvariant()
    if ($App -eq 'access') { $null = Get-OfficeDef 'access'; Assert-ReviewOpen; return $null }
    $null = Get-OfficeDef $App
    $s = Get-OfficeSession $App
    if (-not $s) { throw "$($script:OfficeAppDefs[$App].Label) is not started. Call office_start first." }
    $s
}

# Starts the app in its own instance. offscreen (default) shows it past the desktop
# edge so windows can be captured; hidden never shows it; onscreen is for a dedicated
# review machine.
function Start-OfficeApp {
    param([Parameter(Mandatory)][string]$App, [ValidateSet('offscreen', 'hidden', 'onscreen')][string]$Mode = 'offscreen')
    $App = $App.ToLowerInvariant()
    if ($App -eq 'access') { throw 'Access is started with open_app (it opens a front end); office_invoke can then use app access.' }
    $def = Get-OfficeDef $App
    $existing = Get-OfficeSession $App
    if ($existing) { return (Get-OfficeSessionInfo $existing) }
    if (-not (Test-OfficeRegistered $def.ProgId)) { throw "$($def.Label) is not installed here (no $($def.ProgId) COM class)." }
    $attachOk = (Get-OfficeSettingList 'ACCESS_REVIEW_OFFICE_ATTACH' '') -contains $App
    $before = @(Get-Process -Name $def.Process -ErrorAction SilentlyContinue | ForEach-Object Id)
    if ($def.Single -and $before.Count -and -not $attachOk) {
        throw "$($def.Label) is already running for this user and keeps one instance, so starting it would take over the user's own copy. Close it first, or list $App in ACCESS_REVIEW_OFFICE_ATTACH to allow attaching."
    }
    $keeper = $null
    $com = New-Object -ComObject $def.ProgId
    $procId = 0
    if ($App -eq 'excel') { try { $procId = [int][KitReviewWin]::PidOf([IntPtr][long]$com.Hwnd) } catch {} }
    if (-not $procId) {
        for ($k = 0; $k -lt 20 -and -not $procId; $k++) {
            $new = @(Get-Process -Name $def.Process -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
            if ($new.Count) { $procId = $new[0] } else { Start-Sleep -Milliseconds 250 }
        }
    }
    $started = [bool]$procId -and ($before -notcontains $procId)
    if (-not $started) {
        if (-not $attachOk) {
            Release-Com $com; [GC]::Collect()
            throw "Windows handed back the user's running $($def.Label) instead of a new one. Close it first, or list $App in ACCESS_REVIEW_OFFICE_ATTACH."
        }
        if ($before.Count) { $procId = $before[0] }
        $Mode = 'attached'
    }
    $proc = $null; if ($procId) { $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue }
    if ($Mode -eq 'offscreen' -and $procId) { $keeper = New-Object KitReviewKeeper -ArgumentList ([uint32]$procId); $keeper.Start() }
    if ($started) {
        # Alerts off, macros off for every file opened (office_open can turn them on for one file).
        switch ($App) {
            'excel'      { $null = Set-OfficeProp $com 'DisplayAlerts' $false; $null = Set-OfficeProp $com 'AskToUpdateLinks' $false }
            'word'       { $null = Set-OfficeProp $com 'DisplayAlerts' 0 }
            'powerpoint' { $null = Set-OfficeProp $com 'DisplayAlerts' 1 }
            'visio'      { $null = Set-OfficeProp $com 'AlertResponse' 7 }
        }
        $null = Set-OfficeProp $com 'AutomationSecurity' 3   # msoAutomationSecurityForceDisable
        if ($Mode -ne 'hidden' -and $App -ne 'outlook' -and $App -ne 'powerpoint') { $null = Set-OfficeProp $com 'Visible' $true }
        if ($App -eq 'powerpoint' -and $Mode -ne 'hidden') { $null = Set-OfficeProp $com 'Visible' -1 }
        if ($Mode -eq 'onscreen' -and $procId) { foreach ($h in [KitReviewWin]::TopWindows([uint32]$procId)) { [void][KitReviewWin]::Show($h, 0, 0) } }
        $null = Set-OfficeProp $com 'UserControl' $false
    }
    $sec = Get-OfficeProp $com 'AutomationSecurity'
    $s = [pscustomobject]@{
        App = $App; Label = $def.Label; Com = $com; Pid = $procId; Process = $proc; Started = $started; Mode = $Mode
        Keeper = $keeper; Version = [string](Get-OfficeProp $com 'Version'); Dialogs = @()
        MacrosForcedOff = $(if ($null -ne $sec) { [int]$sec -eq 3 } else { $null })
    }
    $script:OfficeSessions[$App] = $s
    Get-OfficeSessionInfo $s
}

function Get-OfficeSessionInfo($s) {
    [pscustomobject][ordered]@{
        App = $s.App; Version = $s.Version; Pid = $s.Pid; Mode = $s.Mode; StartedByTool = $s.Started; MacrosForcedOff = $s.MacrosForcedOff
        Documents = @($script:OfficeHandles.GetEnumerator() | Where-Object { $_.Value.App -eq $s.App -and $_.Value.Doc } | ForEach-Object { [pscustomobject]@{ Handle = $_.Key; Name = $_.Value.Name; Path = $_.Value.Path; MacrosEnabled = $_.Value.Macros } })
        Output = (Get-OfficeOut)
    }
}

# Runs a block against an app under a KitReviewGuard: message boxes are captured and
# closed, and a call that hangs ends the app (only an instance this tool started).
function Invoke-OfficeGuarded {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][scriptblock]$Body, [int]$TimeoutSec = 120, [int]$ModalSec = 4)
    # ModalSec 0 leaves new windows open (opening a document makes one, and a large file
    # can take longer than the few seconds after which a blocking window is closed).
    # The body runs in this function's child scope and reads its caller's variables, so
    # the locals here have names no caller uses.
    if ($App -eq 'access') { return (Invoke-ReviewGuarded -Body $Body -TimeoutSec $TimeoutSec -ModalSec $ModalSec) }
    $ogSession = Assert-OfficeSession $App
    $ogGuard = $null
    if ($ogSession.Pid) {
        $ogLimit = 0; if ($ogSession.Started) { $ogLimit = $TimeoutSec }
        $ogGuard = New-Object KitReviewGuard -ArgumentList ([uint32]$ogSession.Pid), (Get-OfficeOut), ($ogSession.Mode -eq 'offscreen'), $ogLimit, $ModalSec, $ogSession.Label
        $ogGuard.Start()
    }
    $ogResult = $null; $ogFailure = $null
    try { $ogResult = & $Body } catch { $ogFailure = Get-OfficeError $_ }
    $ogEvents = @(); if ($ogGuard) { $ogEvents = @($ogGuard.Stop()) }
    $ogSession.Dialogs += $ogEvents
    if ($ogGuard -and $ogGuard.Killed) { Clear-OfficeHandles $App; $script:OfficeSessions.Remove($App) }
    if ($ogFailure) {
        $ogMsg = "$($ogSession.Label) call failed: $ogFailure"
        if ($ogEvents.Count) { $ogMsg += "`nDialogs: " + ($ogEvents -join "`n") }
        throw $ogMsg
    }
    [pscustomobject]@{ Value = $ogResult; Dialogs = $ogEvents }
}

# --- handles ---

function Add-OfficeHandle($Obj, [string]$App, [hashtable]$Extra = @{}) {
    $Obj = Unwrap $Obj
    foreach ($kv in $script:OfficeHandles.GetEnumerator()) { if ([object]::ReferenceEquals($kv.Value.Obj, $Obj)) { return $kv.Key } }
    $script:OfficeNextHandle++
    $id = "h$($script:OfficeNextHandle)"
    # An app object is shared with its session; releasing its handle must not release it.
    $shared = [bool]@($script:OfficeSessions.Values | Where-Object { [object]::ReferenceEquals((Unwrap $_.Com), $Obj) }).Count -or
              ($script:Review -and [object]::ReferenceEquals((Unwrap $script:Review.Acc), $Obj))
    $rec = @{ Obj = $Obj; App = $App; Type = [KitOfficeCom]::TypeName($Obj); Doc = $false; Macros = $false; Name = $null; Path = $null; Shared = $shared }
    foreach ($k in $Extra.Keys) { $rec[$k] = $Extra[$k] }
    $script:OfficeHandles[$id] = $rec
    $id
}

function Get-OfficeHandle([string]$Id) {
    $h = $script:OfficeHandles[$Id]
    if (-not $h) { throw "No handle $Id (it was released, or its app was closed)." }
    $h
}

function Clear-OfficeHandles([string]$App) {
    foreach ($k in @($script:OfficeHandles.Keys)) {
        $h = $script:OfficeHandles[$k]
        if ($h.App -eq $App) { if (-not $h.Shared) { try { Release-Com $h.Obj } catch {} }; $script:OfficeHandles.Remove($k) }
    }
}

function Remove-OfficeHandle {
    param([Parameter(Mandatory)][string[]]$Handles)
    $done = @()
    foreach ($id in $Handles) {
        $h = $script:OfficeHandles[$id]
        if (-not $h) { continue }
        if ($h.Doc) { throw "$id is an open document; close it with office_close." }
        if (-not $h.Shared) { try { Release-Com $h.Obj } catch {} }
        $script:OfficeHandles.Remove($id); $done += $id
    }
    [pscustomobject]@{ Released = $done; Remaining = $script:OfficeHandles.Count }
}

# The app and object a target names: an app (excel, word, access...) or a handle (h3).
function Resolve-OfficeTarget([string]$Target) {
    if ($Target -match '^h\d+$') { $h = Get-OfficeHandle $Target; return [pscustomobject]@{ App = $h.App; Obj = $h.Obj } }
    $app = ([string]$Target).ToLowerInvariant()
    $s = Assert-OfficeSession $app
    if ($app -eq 'access') { return [pscustomobject]@{ App = 'access'; Obj = $script:Review.Acc } }
    [pscustomobject]@{ App = $app; Obj = $s.Com }
}

function ConvertTo-OfficeResult($Value, [string]$App, [int]$MaxCells = 10000) {
    # No GetNewClosure: a closure runs in a module of its own, where $script: is not this
    # script's scope. Called from ConvertFrom-OfficeComValue, the block still sees $App.
    $on = { param($o) $id = Add-OfficeHandle $o $App; [pscustomobject]@{ Handle = $id; Type = (Get-OfficeHandle $id).Type } }
    ConvertFrom-OfficeComValue $Value $on $MaxCells
}

# --- the generic layer ---

function Resolve-OfficeArgs([object[]]$Values) {
    $resolve = { param($id) , (Get-OfficeHandle $id).Obj }
    $out = New-Object object[] @($Values).Count
    $i = 0
    foreach ($v in @($Values)) {
        if ($null -ne $v -and $v.PSObject -and $v.PSObject.Properties['OfficeMissing']) { $out[$i] = [Type]::Missing }
        else { $c = ConvertTo-OfficeComValue $v $resolve; $out[$i] = Unwrap $c }
        $i++
    }
    , $out
}

$script:FlagGet = [Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::GetProperty

# A call through the adapter. The adapter wants [ref] for by-reference parameters (Word's
# Range(Start, End)), which JSON arguments can't carry, so that one refusal is retried
# through InvokeMember, which passes by-reference arguments as they are.
function Invoke-OfficeAdapterCall($Member, $Obj, [string]$Name, [Reflection.BindingFlags]$Flags, [object[]]$Plain) {
    try { return , $Member.Invoke($Plain) }
    catch {
        if ($_.Exception.Message -notlike '*PSReference*') { throw }
        return , [KitOfficeCom]::Invoke($Obj, $Name, [int]$Flags, $Plain, [string[]]@())
    }
}

# One IDispatch call. Without named arguments it goes through PowerShell's own COM
# adapter, which reads the type library and calls each member the way it is declared
# (method, property, or property with parameters such as Range("A1")); Office apps
# refuse some calls made with the combined get-or-call flags that reflection sends.
# Objects the adapter can't describe (Access forms) and calls with named arguments go
# through KitOfficeCom.Invoke (Type.InvokeMember).
function Invoke-OfficeDispatch($Obj, [string]$Name, [Reflection.BindingFlags]$Flags, [object[]]$Argv, [string[]]$Names) {
    $named = [string[]]@(); if ($Names) { $named = [string[]]@($Names | Where-Object { $_ }) }
    if ($null -eq $Argv) { $Argv = @() }
    $o = Unwrap $Obj
    $isSet = [bool]($Flags -band [Reflection.BindingFlags]::SetProperty)
    if (-not $named.Count -and (Test-OfficeObject $o)) {
        $m = $null
        try { $m = @($o.PSObject.Members.Match($Name) | Where-Object { $_.MemberType -in 'Method', 'ParameterizedProperty', 'Property' }) | Select-Object -First 1 } catch {}
        if ($m) {
            $plain = New-Object object[] $Argv.Count
            for ($k = 0; $k -lt $Argv.Count; $k++) { $plain[$k] = Unwrap $Argv[$k] }
            switch ($m.MemberType) {
                'Method' {
                    if ($isSet) { break }
                    return , (Invoke-OfficeAdapterCall $m $o $Name ([Reflection.BindingFlags]::InvokeMethod) $plain)
                }
                'ParameterizedProperty' {
                    if ($isSet) {
                        $idx = New-Object object[] ($plain.Count - 1)
                        for ($k = 0; $k -lt $idx.Count; $k++) { $idx[$k] = $plain[$k] }
                        $m.InvokeSet($plain[$plain.Count - 1], $idx); return
                    }
                    return , (Invoke-OfficeAdapterCall $m $o $Name ([Reflection.BindingFlags]::GetProperty) $plain)
                }
                'Property' {
                    if ($isSet -and $plain.Count -eq 1) { $m.Value = $plain[0]; return }
                    if (-not $isSet -and -not $plain.Count) { return , $m.Value }
                    # A property given arguments, as VBA reads it: the arguments go to the
                    # value's default member, Item (Worksheets(1), Cells(2, 3)).
                    $inner = Unwrap $m.Value
                    if ((Test-OfficeObject $inner) -and @($inner.PSObject.Members.Match('Item')).Count) {
                        return , (Invoke-OfficeDispatch $inner 'Item' $Flags $plain $null)
                    }
                    break
                }
            }
        }
    }
    , [KitOfficeCom]::Invoke($o, $Name, [int]$Flags, $Argv, $named)
}

# One step of a path. When a member given arguments takes none itself (Workbooks(1) in
# Word's Documents(1) style), reads the member and applies the arguments to its Item.
function Invoke-OfficeStep($Obj, $Seg, [object[]]$Argv) {
    try { return , (Invoke-OfficeDispatch $Obj $Seg.Name $script:FlagGet $Argv $Seg.Names) }
    catch {
        $hr = 0; $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }; $hr = $e.HResult
        if ($Argv.Count -and -not $Seg.Names.Count -and ($hr -eq 0x8002000E -or $hr -eq -2147352562)) {
            $coll = Invoke-OfficeDispatch $Obj $Seg.Name $script:FlagGet @() $null
            return , (Invoke-OfficeDispatch $coll 'Item' $script:FlagGet $Argv $null)
        }
        throw
    }
}

# Reading a range's cells copies all of them over COM before any output limit applies,
# so a range bigger than the limit is refused before it is read.
$script:OfficeCellReads = @('Value', 'Value2', 'Formula', 'Formula2', 'FormulaR1C1', 'FormulaLocal', 'FormulaR1C1Local', 'Text')
function Assert-OfficeCellCount($Range, [int]$MaxCells) {
    $n = $null
    try { $n = [long](Invoke-OfficeDispatch $Range 'CountLarge' $script:FlagGet @() $null) } catch { return }
    if ($n -gt $MaxCells) { throw "That range has $n cells; at most $MaxCells are read at once. Ask for a smaller range." }
}

function Assert-OfficeArgs([string]$App, $Seg, [object[]]$Raw) {
    $plain = @(foreach ($r in $Raw) { if ($null -ne $r -and $r.PSObject -and ($r.PSObject.Properties['OfficeMissing'] -or $r.PSObject.Properties['OfficeHandle'])) { $null } else { $r } })
    # A path argument must sit in an allowed folder (lexically) and pass the same checks
    # office_open makes: no junction or link hop out of the folder, no Mark of the Web.
    Assert-OfficeCallAllowed -App $App -Member $Seg.Name -Arguments $plain -InAllowed { param($p) if (-not (Test-OfficePathAllowed $p)) { return $false }; Assert-OfficeFileSafe $p; $true }
}

# Walks a member path from a target, and reads the end (get), sets it (Value given) or
# calls it with extra JSON arguments (Arguments, appended to the last member's own).
# A COM object result comes back as a handle; with Fields, those members are read
# from it, and with Enumerate, a collection's items (Item(1..Count)) are listed.
function Invoke-OfficeMember {
    param([Parameter(Mandatory)][string]$Target, [AllowEmptyString()][string]$Path = '', $Value, [switch]$Set, [object[]]$Arguments,
          [string[]]$Fields, [switch]$Enumerate, [int]$MaxItems = 50, [int]$MaxCells = 10000)
    $t = Resolve-OfficeTarget $Target
    $segs = @(ConvertFrom-OfficePath $Path)
    if ($Set -and -not $segs.Count) { throw 'Give the property to set as the path.' }
    if ($Set -and $t.App -eq 'outlook') { throw 'Outlook is read only for these tools: properties can be read, not set.' }
    $fieldSegs = @(foreach ($f in @($Fields)) { if ($f) { , @(ConvertFrom-OfficePath $f) } })
    # Check every member and argument before the first call, so a refused path changes nothing.
    for ($i = 0; $i -lt $segs.Count; $i++) {
        $raw = @($segs[$i].Args)
        if ($i -eq $segs.Count - 1) { if ($Arguments) { $raw += @($Arguments) }; if ($Set) { $raw += , $Value } }
        Assert-OfficeArgs $t.App $segs[$i] $raw
    }
    foreach ($fs in $fieldSegs) { foreach ($sg in $fs) { Assert-OfficeArgs $t.App $sg @($sg.Args) } }
    $app = $t.App
    $omValue = $Value   # Invoke-ReviewGuarded has a local $value, which would hide $Value in the body
    $r = Invoke-OfficeGuarded -App $app -Body {
        $cur = $t.Obj
        for ($i = 0; $i -lt $segs.Count; $i++) {
            $seg = $segs[$i]; $last = $i -eq $segs.Count - 1
            if ($null -eq $cur -or -not (Test-OfficeObject (Unwrap $cur))) { throw "'$(($segs[0..([Math]::Max(0, $i - 1))] | ForEach-Object Name) -join '.')' is not an object, so '$($seg.Name)' can't be read from it." }
            if ($last -and -not $Set -and $app -eq 'excel' -and $seg.Name -in $script:OfficeCellReads) { Assert-OfficeCellCount $cur $MaxCells }
            $argv = Resolve-OfficeArgs @($seg.Args)
            if ($last -and $Arguments) { $argv = [object[]](@($argv) + (Resolve-OfficeArgs @($Arguments))) }
            if ($last -and $Set) {
                $v = Resolve-OfficeArgs @(, $omValue)
                [void](Invoke-OfficeDispatch $cur $seg.Name ([Reflection.BindingFlags]::SetProperty) ([object[]](@($argv) + @($v))) $seg.Names)
                $cur = Invoke-OfficeStep $cur $seg $argv
            } else {
                $cur = Invoke-OfficeStep $cur $seg $argv
            }
        }
        $res = [ordered]@{ Value = (ConvertTo-OfficeResult $cur $app $MaxCells) }
        $obj = Unwrap $cur
        if (Test-OfficeObject $obj) {
            if ($fieldSegs.Count) { $res.Fields = Read-OfficeFields $obj $fieldSegs $Fields $app }
            if ($Enumerate) {
                $count = [int](Invoke-OfficeDispatch $obj 'Count' $script:FlagGet @() $null)
                $items = @()
                for ($k = 1; $k -le [Math]::Min($count, $MaxItems); $k++) {
                    $it = Invoke-OfficeDispatch $obj 'Item' $script:FlagGet @($k) $null
                    $row = [ordered]@{ Index = $k; Item = (ConvertTo-OfficeResult $it $app) }
                    $fs = $fieldSegs; $fn = $Fields
                    if (-not $fs.Count) { $fs = @(, @(ConvertFrom-OfficePath 'Name')); $fn = @('Name') }
                    $row.Fields = Read-OfficeFields (Unwrap $it) $fs $fn $app
                    $items += [pscustomobject]$row
                }
                $res.Count = $count; $res.Items = $items; $res.Truncated = $count -gt $MaxItems
            }
        }
        [pscustomobject]$res
    }
    $out = $r.Value
    $out | Add-Member -NotePropertyName Dialogs -NotePropertyValue $r.Dialogs
    $out
}

function Read-OfficeFields($Obj, [object[]]$FieldSegs, [string[]]$Names, [string]$App) {
    $o = [ordered]@{}
    for ($j = 0; $j -lt $FieldSegs.Count; $j++) {
        $cur = $Obj
        try {
            foreach ($sg in @($FieldSegs[$j])) { $cur = Invoke-OfficeStep $cur $sg (Resolve-OfficeArgs @($sg.Args)) }
            $o[$Names[$j]] = ConvertTo-OfficeResult $cur $App 2000
        } catch { $o[$Names[$j]] = "(error: $(Get-OfficeError $_))" }
    }
    [pscustomobject]$o
}

# What an object offers, from its type library: methods, properties (get/put) and
# their parameters, optionally filtered by a name pattern.
function Get-OfficeMembers {
    param([Parameter(Mandatory)][string]$Target, [AllowEmptyString()][string]$Path = '', [string]$Filter)
    $t = Resolve-OfficeTarget $Target
    $segs = @(ConvertFrom-OfficePath $Path)
    foreach ($sg in $segs) { Assert-OfficeArgs $t.App $sg @($sg.Args) }
    $r = Invoke-OfficeGuarded -App $t.App -Body {
        $cur = $t.Obj
        foreach ($sg in $segs) { $cur = Invoke-OfficeStep $cur $sg (Resolve-OfficeArgs @($sg.Args)) }
        $obj = Unwrap $cur
        if (-not (Test-OfficeObject $obj)) { throw "'$Path' is a value ($obj), not an object." }
        $rows = @([KitOfficeCom]::Members($obj))
        $by = [ordered]@{}
        foreach ($row in $rows) {
            $n = $row[0]
            if ($Filter -and $n -notlike $Filter) { continue }
            if (-not $by.Contains($n)) { $by[$n] = [ordered]@{ Name = $n; Kinds = @(); Parameters = $row[2]; Description = $row[3] } }
            if ($by[$n].Kinds -notcontains $row[1]) { $by[$n].Kinds += $row[1] }
            if ($row[1] -eq 'method' -or $row[1] -eq 'get') { $by[$n].Parameters = $row[2] }
            if (-not $by[$n].Description) { $by[$n].Description = $row[3] }
        }
        $list = @(foreach ($m in $by.Values) {
            $m.Refused = [bool](Get-OfficeMemberRefusal $t.App $m.Name)
            $m.Kinds = ($m.Kinds -join '/')
            [pscustomobject]$m
        })
        [pscustomobject][ordered]@{ Type = [KitOfficeCom]::TypeName($obj); Count = $list.Count; Members = $list }
    }
    $r.Value
}

# --- typed helpers ---

function Get-OfficeApps {
    $allowed = Get-OfficeAllowedApps
    $attach = Get-OfficeSettingList 'ACCESS_REVIEW_OFFICE_ATTACH' ''
    @(foreach ($k in $script:OfficeAppDefs.Keys) {
        $d = $script:OfficeAppDefs[$k]
        $s = $null; if ($k -ne 'access') { $s = Get-OfficeSession $k }
        $running = @(Get-Process -Name $d.Process -ErrorAction SilentlyContinue).Count
        [pscustomobject][ordered]@{
            App = $k; Installed = (Test-OfficeRegistered $d.ProgId); Enabled = ($allowed -contains $k)
            Started = $(if ($k -eq 'access') { [bool]$script:Review } else { [bool]$s }); Pid = $(if ($s) { $s.Pid }); Version = $(if ($s) { $s.Version })
            OtherCopiesRunning = [Math]::Max(0, $running - $(if ($s -and $s.Started) { 1 } else { 0 }))
            OneInstancePerUser = $d.Single; AttachAllowed = ($attach -contains $k)
        }
    })
}

function Assert-OfficeDocApp([string]$App) {
    if ($App -notin 'excel', 'word', 'powerpoint', 'visio') { throw "office_open handles Excel, Word, PowerPoint and Visio files. For $App use office_invoke$(if ($App -eq 'access') { ' after open_app' })." }
}

# Opens an existing file from an allowed folder. read_only (default true) still lets
# office_save_as write a copy. Macros stay disabled unless EnableMacros.
function Open-OfficeDocument {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$Path, [bool]$ReadOnly = $true, [switch]$EnableMacros)
    $App = $App.ToLowerInvariant(); Assert-OfficeDocApp $App
    $s = Assert-OfficeSession $App
    if (-not (Test-OfficeFullPath $Path)) { throw "'$Path' is not a full path." }
    $Path = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "'$Path' does not exist. office_open opens existing files; office_new makes a blank one." }
    if (-not (Test-OfficePathAllowed $Path)) { throw "'$Path' is outside the folders these tools may use ($(Get-OfficeRoots))." }
    Assert-OfficeFileSafe $Path
    if (-not $s.Started) { throw "$($s.Label) is the user's own instance; documents are not opened in it." }
    $com = $s.Com
    $win = $s.Mode -ne 'hidden'
    $r = Invoke-OfficeGuarded -App $App -ModalSec 0 -TimeoutSec 300 -Body {
        if ($EnableMacros) { $null = Set-OfficeProp $com 'AutomationSecurity' 1 }
        try {
            switch ($App) {
                'excel'      { $com.Workbooks.Open($Path, 0, $ReadOnly) }
                'word'       { $com.Documents.Open($Path, $false, $ReadOnly, $false) }
                'powerpoint' { $com.Presentations.Open($Path, $(if ($ReadOnly) { -1 } else { 0 }), 0, $(if ($win) { -1 } else { 0 })) }
                'visio'      { $f = 0x80; if ($EnableMacros) { $f = 0 }; if ($ReadOnly) { $f = $f -bor 2 }; $com.Documents.OpenEx($Path, $f) }
            }
        } finally { $null = Set-OfficeProp $com 'AutomationSecurity' 3 }
    }
    Register-OfficeDocument $App $r.Value $Path ([bool]$EnableMacros) $r.Dialogs
}

function New-OfficeDocument {
    param([Parameter(Mandatory)][string]$App)
    $App = $App.ToLowerInvariant(); Assert-OfficeDocApp $App
    $s = Assert-OfficeSession $App
    if (-not $s.Started) { throw "$($s.Label) is the user's own instance; documents are not made in it." }
    $com = $s.Com; $win = $s.Mode -ne 'hidden'
    $r = Invoke-OfficeGuarded -App $App -ModalSec 0 -Body {
        switch ($App) {
            'excel'      { $com.Workbooks.Add() }
            'word'       { $com.Documents.Add() }
            'powerpoint' { $com.Presentations.Add($(if ($win) { -1 } else { 0 })) }
            'visio'      { $com.Documents.Add('') }
        }
    }
    Register-OfficeDocument $App $r.Value $null $false $r.Dialogs
}

function Register-OfficeDocument([string]$App, $Doc, [string]$Path, [bool]$Macros, $Dialogs) {
    $Doc = Unwrap $Doc
    $name = [string](Get-OfficeProp $Doc 'Name')
    $id = Add-OfficeHandle $Doc $App @{ Doc = $true; Macros = $Macros; Name = $name; Path = $Path }
    $ro = Get-OfficeProp $Doc 'ReadOnly'
    [pscustomobject][ordered]@{ Handle = $id; App = $App; Type = $script:OfficeHandles[$id].Type; Name = $name; Path = $Path; ReadOnly = $(if ($null -ne $ro) { [bool]$ro }); MacrosEnabled = $Macros; Dialogs = $Dialogs }
}

function Get-OfficeDocHandle([string]$Handle) {
    $h = Get-OfficeHandle $Handle
    if (-not $h.Doc) { throw "$Handle is not a document handle from office_open or office_new." }
    $h
}

# Closes a document without saving (changes are kept only through office_save_as).
function Close-OfficeDocument {
    param([Parameter(Mandatory)][string]$Handle)
    $h = Get-OfficeDocHandle $Handle
    $d = $h.Obj
    $r = Invoke-OfficeGuarded -App $h.App -Body {
        switch ($h.App) {
            'excel'      { [void]$d.Close($false) }
            'word'       { [void]$d.Close(0) }
            'powerpoint' { $null = Set-OfficeProp $d 'Saved' -1; [void]$d.Close() }
            'visio'      { $null = Set-OfficeProp $d 'Saved' $true; [void]$d.Close() }
        }
    }
    try { Release-Com $d } catch {}
    $script:OfficeHandles.Remove($Handle)
    [pscustomobject]@{ Closed = $Handle; Dialogs = $r.Dialogs }
}

function Assert-OfficeOutPath([string]$Path, [bool]$Overwrite) {
    if (-not (Test-OfficeFullPath $Path)) { throw "'$Path' is not a full path." }
    if ($Path -match '\.\.') { throw "'$Path' contains '..'; give the real full path." }
    $Path = [IO.Path]::GetFullPath($Path)
    if (-not (Test-OfficePathAllowed $Path)) { throw "'$Path' is outside the folders these tools may write ($(Get-OfficeRoots))." }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw "Folder '$dir' does not exist." }
    Assert-OfficeFileSafe $dir
    if ((Test-Path -LiteralPath $Path) -and -not $Overwrite) { throw "'$Path' exists. Pass overwrite=true to replace it." }
    $Path
}

# Saves the document under a new full path (the open document then is that file).
# Format is the app's file-format number (e.g. Excel 51 xlsx, Word 16 docx, PowerPoint
# 24 pptx); default keeps the app's choice for the extension.
function Save-OfficeDocumentAs {
    param([Parameter(Mandatory)][string]$Handle, [Parameter(Mandatory)][string]$Path, [int]$Format = -1, [switch]$Overwrite)
    $h = Get-OfficeDocHandle $Handle
    $Path = Assert-OfficeOutPath $Path $Overwrite
    $d = $h.Obj
    $r = Invoke-OfficeGuarded -App $h.App -Body {
        switch ($h.App) {
            'excel'      { if ($Format -ge 0) { [void]$d.SaveAs($Path, $Format) } else { [void]$d.SaveAs($Path) } }
            'word'       { if ($Format -ge 0) { [void]$d.SaveAs2($Path, $Format) } else { [void]$d.SaveAs2($Path) } }
            'powerpoint' { if ($Format -ge 0) { [void]$d.SaveAs($Path, $Format) } else { [void]$d.SaveAs($Path) } }
            'visio'      { [void]$d.SaveAs($Path) }
        }
    }
    $h.Path = $Path; $h.Name = [string](Get-OfficeProp $d 'Name')
    [pscustomobject]@{ Saved = $Path; Bytes = (Get-Item -LiteralPath $Path).Length; Dialogs = $r.Dialogs }
}

# Writes a PDF of the document (default <output>\<name>.pdf); the document is unchanged.
function Export-OfficePdf {
    param([Parameter(Mandatory)][string]$Handle, [string]$Path, [switch]$Overwrite)
    $h = Get-OfficeDocHandle $Handle
    if (-not $Path) { $Path = Join-Path (Get-OfficeOut) (([IO.Path]::GetFileNameWithoutExtension([string]$h.Name)) + '.pdf') }
    $Path = Assert-OfficeOutPath $Path $Overwrite
    $d = $h.Obj
    $r = Invoke-OfficeGuarded -App $h.App -Body {
        switch ($h.App) {
            'excel'      { [void]$d.ExportAsFixedFormat(0, $Path) }         # xlTypePDF
            'word'       { [void]$d.ExportAsFixedFormat($Path, 17) }        # wdExportFormatPDF
            'powerpoint' { [void]$d.SaveCopyAs($Path, 32) }                 # ppSaveAsPDF
            'visio'      { [void]$d.ExportAsFixedFormat(1, $Path, 1, 0) }   # visFixedFormatPDF, print quality, all pages
        }
    }
    [pscustomobject]@{ Pdf = $Path; Bytes = (Get-Item -LiteralPath $Path).Length; Dialogs = $r.Dialogs }
}

# PNG of an app's window (or a document's window) drawn with PrintWindow until two
# frames match. Only instances this tool started are captured.
function Save-OfficeCapture {
    param([Parameter(Mandatory)][string]$Target, [string]$Label, [int[]]$Crop, [int]$Zoom = 1)
    if ($Label -and $Label -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') { throw "Label '$Label': letters, digits, '.', '_' or '-' only." }
    $t = Resolve-OfficeTarget $Target
    if ($t.App -eq 'access') { throw 'Capture Access forms with capture.' }
    $s = Assert-OfficeSession $t.App
    if (-not $s.Started) { throw "$($s.Label) is the user's own instance; its windows are not captured." }
    if ($s.Mode -eq 'hidden') { throw "$($s.Label) was started hidden; start it offscreen to capture." }
    $hwnd = [IntPtr]::Zero
    $o = Unwrap $t.Obj
    if ($Target -match '^h\d+$') {
        foreach ($p in 'ActiveWindow.Hwnd', 'Windows(1).HWND', 'Windows(1).Hwnd', 'Windows(1).WindowHandle32') {
            try { $cur = $o; foreach ($sg in @(ConvertFrom-OfficePath $p)) { $cur = Invoke-OfficeStep $cur $sg (Resolve-OfficeArgs @($sg.Args)) }; if ($cur) { $hwnd = [IntPtr][long]$cur; break } } catch {}
        }
    } elseif ($t.App -eq 'excel') { try { $hwnd = [IntPtr][long]$s.Com.Hwnd } catch {} }
    if ($hwnd -eq [IntPtr]::Zero) {
        $wins = @([KitReviewWin]::TopWindows([uint32]$s.Pid) | Sort-Object { $r = [KitReviewWin]::Rect($_); - ($r[2] * $r[3]) })
        if (-not $wins.Count) { throw "$($s.Label) has no window to capture (open a document first)." }
        $hwnd = $wins[0]
    }
    $name = $t.App + $(if ($Label) { "-$Label" }) + '-' + (Get-Date -Format 'HHmmss') + '.png'
    $file = Join-Path (Get-OfficeOut) $name
    $c = @(0, 0, 0, 0); if ($Crop -and $Crop.Count -eq 4) { $c = $Crop }
    $prev = $null; $res = $null
    for ($k = 0; $k -lt 6; $k++) {
        Start-Sleep -Milliseconds 300
        $res = [KitReviewWin]::Save($hwnd, $file, $c[0], $c[1], $c[2], $c[3], [Math]::Max(1, $Zoom))
        $hash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
        if ($hash -eq $prev) { break }
        $prev = $hash
    }
    [pscustomobject][ordered]@{ Path = $file; Width = $res[0]; Height = $res[1]; Blank = [bool]$res[2] }
}

# Runs a macro in a document opened with enable_macros (Application.Run).
function Invoke-OfficeMacro {
    param([Parameter(Mandatory)][string]$Handle, [Parameter(Mandatory)][string]$Macro, [object[]]$Arguments = @(), [int]$TimeoutSec = 120)
    $h = Get-OfficeDocHandle $Handle
    if (-not $h.Macros) { throw "$Handle was opened with macros disabled. Close it and open it again with enable_macros=true (only for files you trust)." }
    if ($Macro -notmatch '^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*){0,2}$') { throw "Macro '$Macro': give Module.Macro or Macro." }
    if ($Arguments.Count -gt 10) { throw 'At most 10 arguments.' }
    foreach ($a in $Arguments) { Assert-OfficeCallAllowed -App $h.App -Member 'MacroArgument' -Arguments @(, $a) -InAllowed { param($p) Test-OfficePathAllowed $p } }
    $s = Assert-OfficeSession $h.App
    $d = $h.Obj
    $full = switch ($h.App) {
        'excel'      { "'" + $h.Name.Replace("'", "''") + "'!" + $Macro }
        'powerpoint' { "'" + $h.Name.Replace("'", "''") + "'!" + $Macro }
        default      { $Macro }
    }
    $argv = Resolve-OfficeArgs $Arguments
    $r = Invoke-OfficeGuarded -App $h.App -TimeoutSec $TimeoutSec -Body {
        if ($h.App -eq 'word') { try { [void]$d.Activate() } catch {} }
        [KitOfficeCom]::Invoke($s.Com, 'Run', [int][Reflection.BindingFlags]::InvokeMethod, ([object[]](@($full) + @($argv))), $null)
    }
    [pscustomobject]@{ Macro = $full; Result = (ConvertTo-OfficeResult $r.Value $h.App); Dialogs = $r.Dialogs }
}

# Reads (and optionally writes) a block of cells. Values are rows of cells as Value2
# (numbers, text, true/false; dates as serial numbers); Formulas writes formulas.
function Invoke-ExcelRange {
    param([Parameter(Mandatory)][string]$Handle, [string]$Sheet, [Parameter(Mandatory)][string]$Range, $Values, [switch]$AsFormulas, [switch]$IncludeFormulas)
    $h = Get-OfficeDocHandle $Handle
    if ($h.App -ne 'excel') { throw "$Handle is not an Excel workbook." }
    if ($Range -notmatch '^[A-Za-z0-9$:_,.!'' ]+$') { throw "Range '$Range' is not a cell address or name." }
    if ($null -ne $Values) {
        foreach ($t in @(Get-OfficeStrings $Values)) {
            $bad = Get-OfficeValueRefusal $t
            if ($bad) { throw "Refused value: $bad" }
            if ((Test-OfficePathLike $t) -and -not (Test-OfficePathAllowed $t)) { throw "Refused value '$t': a path or address outside the allowed folders." }
        }
    }
    $wb = $h.Obj
    $r = Invoke-OfficeGuarded -App 'excel' -Body {
        $ws = $(if ($Sheet -match '^\d+$') { $wb.Worksheets.Item([int]$Sheet) } elseif ($Sheet) { $wb.Worksheets.Item($Sheet) } else { $wb.ActiveSheet })
        $rng = $ws.Range($Range)
        $cells = [long]$rng.CountLarge
        if ($cells -gt 10000) { throw "$Range has $cells cells; at most 10000 are read at once. Ask for a smaller range." }
        if ($null -ne $Values) {
            $v = ConvertTo-OfficeComValue $Values { param($id) throw 'No handles in cell values.' }
            $prop = $(if ($AsFormulas) { 'Formula' } else { 'Value2' })
            [void][System.__ComObject].InvokeMember($prop, [Reflection.BindingFlags]::SetProperty, $null, $rng, @(, (Unwrap $v)))
        }
        $out = [ordered]@{ Sheet = [string]$ws.Name; Address = [string]$rng.Address($false, $false); Cells = $cells
                           Values = (ConvertTo-OfficeGrid (ConvertFrom-OfficeComValue $rng.Value2 $null 10000)) }
        if ($IncludeFormulas) { $out.Formulas = ConvertTo-OfficeGrid (ConvertFrom-OfficeComValue $rng.Formula $null 10000) }
        [pscustomobject]$out
    }
    $r.Value | Add-Member -NotePropertyName Dialogs -NotePropertyValue $r.Dialogs -PassThru
}

# Paragraphs of a Word document: text, style and outline level, plus counts.
function Get-WordText {
    param([Parameter(Mandatory)][string]$Handle, [int]$Start = 1, [int]$MaxParagraphs = 200)
    $h = Get-OfficeDocHandle $Handle
    if ($h.App -ne 'word') { throw "$Handle is not a Word document." }
    $d = $h.Obj
    $r = Invoke-OfficeGuarded -App 'word' -Body {
        $ps = $d.Paragraphs; $n = [int]$ps.Count
        $list = @(for ($i = [Math]::Max(1, $Start); $i -le [Math]::Min($n, $Start + $MaxParagraphs - 1); $i++) {
            $p = $ps.Item($i)
            $style = $null; try { $style = [string]$p.Style.NameLocal } catch { try { $style = [string]$p.Style } catch {} }
            $txt = ([string]$p.Range.Text).TrimEnd("`r", "`a", "`n")
            if ($txt.Length -gt 2000) { $txt = $txt.Substring(0, 2000) + '...' }
            [pscustomobject][ordered]@{ Index = $i; Style = $style; OutlineLevel = [int]$p.OutlineLevel; Text = $txt }
        })
        [pscustomobject][ordered]@{
            Name = [string]$d.Name; Paragraphs = $n; Pages = [int]$d.ComputeStatistics(2); Words = [int]$d.ComputeStatistics(0)
            Tables = [int]$d.Tables.Count; Shown = $list; More = ($Start + $MaxParagraphs - 1) -lt $n
        }
    }
    $r.Value
}

# Slides of a presentation with their shapes: position and size in points, text and
# tags (the record IDs a deck builder writes, e.g. spm.key).
function Get-PowerPointSlides {
    param([Parameter(Mandatory)][string]$Handle, [int]$MaxText = 500)
    $h = Get-OfficeDocHandle $Handle
    if ($h.App -ne 'powerpoint') { throw "$Handle is not a PowerPoint presentation." }
    $d = $h.Obj
    $tags = { param($o) $t = [ordered]@{}; try { $tg = $o.Tags; for ($k = 1; $k -le [int]$tg.Count; $k++) { $t[[string]$tg.Name($k)] = [string]$tg.Value($k) } } catch {}; [pscustomobject]$t }
    $r = Invoke-OfficeGuarded -App 'powerpoint' -Body {
        $slides = @(foreach ($sl in $d.Slides) {
            $shapes = @(foreach ($sh in $sl.Shapes) {
                $text = $null
                try { if ([int]$sh.HasTextFrame -ne 0) { $text = [string]$sh.TextFrame.TextRange.Text; if ($text.Length -gt $MaxText) { $text = $text.Substring(0, $MaxText) + '...' } } } catch {}
                [pscustomobject][ordered]@{ Name = [string]$sh.Name; Id = [int]$sh.Id; Type = [int]$sh.Type
                    Left = [double]$sh.Left; Top = [double]$sh.Top; Width = [double]$sh.Width; Height = [double]$sh.Height; Text = $text; Tags = (& $tags $sh) }
            })
            $layout = $null; try { $layout = [string]$sl.CustomLayout.Name } catch {}
            [pscustomobject][ordered]@{ Index = [int]$sl.SlideIndex; SlideId = [int]$sl.SlideID; Name = [string]$sl.Name; Layout = $layout; Hidden = ([int]$sl.SlideShowTransition.Hidden -ne 0); Tags = (& $tags $sl); Shapes = $shapes }
        })
        [pscustomobject][ordered]@{ Name = [string]$d.Name; SlideWidth = [double]$d.PageSetup.SlideWidth; SlideHeight = [double]$d.PageSetup.SlideHeight; Count = $slides.Count; Slides = $slides }
    }
    $r.Value
}

# Mail items in an Outlook folder (Inbox, Inbox/Sub, Sent, Drafts, ...), newest first.
# Read only: nothing is marked read, moved or deleted.
function Get-OutlookItems {
    param([string]$Folder = 'Inbox', [int]$MaxItems = 25, [switch]$UnreadOnly, [switch]$IncludeBody)
    $s = Assert-OfficeSession 'outlook'
    $defaults = @{ inbox = 6; sent = 5; 'sent items' = 5; drafts = 16; deleted = 3; 'deleted items' = 3; outbox = 4; junk = 23; calendar = 9; contacts = 10; tasks = 13 }
    $parts = @($Folder -split '[\\/]' | Where-Object { $_ })
    if (-not $parts.Count) { $parts = @('Inbox') }
    $r = Invoke-OfficeGuarded -App 'outlook' -Body {
        $ns = $s.Com.GetNamespace('MAPI')
        $root = $defaults[$parts[0].ToLowerInvariant()]
        if (-not $root) { throw "Start the folder with one of: $(@($defaults.Keys) -join ', ')." }
        $f = $ns.GetDefaultFolder($root)
        foreach ($p in ($parts | Select-Object -Skip 1)) { $f = $f.Folders.Item($p) }
        $items = $f.Items
        try { $items.Sort('[ReceivedTime]', $true) } catch {}
        if ($UnreadOnly) { $items = $items.Restrict('[UnRead] = True') }
        $n = [int]$items.Count
        $list = @(for ($i = 1; $i -le [Math]::Min($n, $MaxItems); $i++) {
            $it = $items.Item($i)
            $row = [ordered]@{ Index = $i; Class = [int]$it.Class; Subject = [string]$it.Subject }
            foreach ($p in 'SenderName', 'ReceivedTime', 'UnRead', 'Size', 'EntryID') { try { $row[$p] = ConvertFrom-OfficeComValue ($it.$p) $null } catch {} }
            try { $row.Attachments = [int]$it.Attachments.Count } catch {}
            if ($IncludeBody) { try { $b = [string]$it.Body; if ($b.Length -gt 4000) { $b = $b.Substring(0, 4000) + '...' }; $row.Body = $b } catch {} }
            [pscustomobject]$row
        })
        [pscustomobject][ordered]@{ Folder = [string]$f.FolderPath; Count = $n; Items = $list }
    }
    $r.Value
}

# Closes this tool's documents without saving, quits an instance it started (never one
# it attached to) and confirms the process has gone.
function Stop-OfficeApp {
    param([Parameter(Mandatory)][string]$App)
    $App = $App.ToLowerInvariant()
    if ($App -eq 'access') { throw 'Close Access with close_app.' }
    $s = Get-OfficeSession $App
    if (-not $s) { return [pscustomobject]@{ App = $App; Closed = $true; Note = 'It was not started.' } }
    $notes = @()
    foreach ($k in @($script:OfficeHandles.Keys)) {
        $h = $script:OfficeHandles[$k]
        if ($h -and $h.App -eq $App -and $h.Doc) { try { $null = Close-OfficeDocument $k } catch { $notes += "Could not close $($h.Name): $($_.Exception.Message)" } }
    }
    Clear-OfficeHandles $App
    if ($s.Started -and -not ($s.Process -and $s.Process.HasExited)) {
        try { [void][System.__ComObject].InvokeMember('Quit', [Reflection.BindingFlags]::InvokeMethod, $null, $s.Com, $(if ($App -eq 'word') { @(0) } else { @() })) } catch { $notes += "Quit failed: $(Get-OfficeError $_)" }
    }
    try { Release-Com $s.Com } catch {}
    $s.Com = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    if ($s.Started -and $s.Process) {
        if (-not $s.Process.WaitForExit(15000)) { try { $s.Process.Kill(); $notes += "$($s.Label) did not quit within 15s; ended it." } catch {} }
    } elseif (-not $s.Started) { $notes += 'Attached to the user''s own instance: released it and left it running.' }
    if ($s.Keeper) { $s.Keeper.Stop() }
    $script:OfficeSessions.Remove($App)
    [pscustomobject]@{ App = $App; Closed = $true; Note = ($notes -join ' ') }
}

function Stop-OfficeAll { foreach ($a in @($script:OfficeSessions.Keys)) { try { $null = Stop-OfficeApp $a } catch {} } }
