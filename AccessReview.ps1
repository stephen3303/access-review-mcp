# Visual review helpers that drive Access without the screen, mouse or keyboard.
#
# Access runs in its own instance started over COM, with its windows parked beyond the
# edge of the desktop. Forms are captured with PrintWindow, which asks Access to draw
# the window into a bitmap, so nothing has to be in front of anyone and nothing takes
# focus. Every Access call runs under a guard that records (and captures) any dialog
# that appears, closes VBA and Office message boxes so the call can return, and ends
# Access if a call hangs.
#
# Used by AccessReviewMcp.ps1 (the MCP server), and usable directly:
#
#   . "$PSScriptRoot\AccessReview.ps1"
#   Open-ReviewApp -Path 'C:\Review\MyApp.accdb'
#   Get-ReviewErrorLog -Mark
#   Open-ReviewForm frmMain -WidthPx 1256 -HeightPx 620
#   Save-ReviewCapture frmMain -Label before
#   Get-ReviewErrorLog
#   Save-ReviewSnapshot -Label pr
#   Close-ReviewApp
#   Compare-ReviewSnapshot -Before main -After pr
#
# Rules it keeps: it only opens an existing file (a new file path brings back the
# Enable Content prompt), it never sends input to the desktop, and it leaves the
# palette on the default (first configured) palette when it closes.
#
# Settings come from environment variables, all optional (see README.md):
#   ACCESS_REVIEW_ALLOWED_FOLDERS    folders open_app may open front ends from, separated by ';'
#   ACCESS_REVIEW_LOCKS              lock files shared by reviewers, separated by ';'
#   ACCESS_REVIEW_OUTPUT             folder for captures (default %LOCALAPPDATA%\AccessReview)
#   ACCESS_REVIEW_PALETTES           name=value pairs for set_palette, separated by ';'
#   ACCESS_REVIEW_PALETTE_PROCEDURE  public VBA procedure that switches palettes

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'FormSnapshot.ps1')

if (-not ('KitReviewWin' -as [type])) {
    Add-Type @'
using System; using System.Text; using System.Threading; using System.Diagnostics;
using System.Collections.Generic; using System.Runtime.InteropServices;
using System.Drawing; using System.Drawing.Drawing2D; using System.Drawing.Imaging;

public static class KitReviewWin {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int i);
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);

    const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10;
    const uint PW_RENDERFULLCONTENT = 0x2;

    // Physical pixels everywhere, so window rectangles and PrintWindow bitmaps agree on a
    // scaled display. Must run before the first window call.
    public static void DpiAware() {
        try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return; } catch { }
        try { SetProcessDPIAware(); } catch { }
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);

    public static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
    static string Text(IntPtr h) { var s = new StringBuilder(4096); GetWindowText(h, s, 4096); return s.ToString(); }
    public static uint PidOf(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }

    // "[title] text | text" for a dialog: child Static/Edit text for VBA and Win32 boxes,
    // UI Automation text for Office-style alerts (NUIDialog), which have no child windows.
    public static string Describe(IntPtr h) {
        var parts = new List<string>();
        EnumChildWindows(h, (ch, l) => {
            var c = Cls(ch); var t = Text(ch).Trim();
            if (t.Length > 0 && (c == "Static" || c == "Edit" || c == "RichEdit20W")) parts.Add(t);
            return true;
        }, IntPtr.Zero);
        if (parts.Count == 0) {
            try {
                var el = System.Windows.Automation.AutomationElement.FromHandle(h);
                var cond = new System.Windows.Automation.PropertyCondition(
                    System.Windows.Automation.AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text);
                foreach (System.Windows.Automation.AutomationElement t in el.FindAll(System.Windows.Automation.TreeScope.Descendants, cond)) {
                    var n = t.Current.Name.Trim();
                    if (n.Length > 0 && n != Text(h)) parts.Add(n);
                }
            } catch { }
        }
        var body = string.Join(" | ", parts).Replace("\r\n", " ").Replace("\n", " ");
        while (body.Contains("  ")) body = body.Replace("  ", " ");
        return "[" + Text(h) + "] " + (body.Length > 0 ? body : "(no text)");
    }

    // Left, top, width, height in physical pixels.
    public static int[] Rect(IntPtr h) {
        RECT r; GetWindowRect(h, out r);
        return new int[] { r.L, r.T, r.R - r.L, r.B - r.T };
    }

    public static IntPtr Root(IntPtr h) { return GetAncestor(h, 2); }   // GA_ROOT

    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);

    // Physical pixels per 96 dpi pixel for a window (1.5 at 150% scaling).
    public static double Scale(IntPtr h) {
        try { uint d = GetDpiForWindow(h); if (d > 0) return d / 96.0; } catch { }
        return 1.0;
    }

    [DllImport("user32.dll")] static extern bool SystemParametersInfo(uint action, uint param, ref RECT r, uint ini);

    // Width and height of the main display less the taskbar: the size a maximized window gets.
    public static int[] WorkArea() {
        var r = new RECT();
        if (SystemParametersInfo(0x30, 0, ref r, 0)) return new int[] { r.R - r.L, r.B - r.T };   // SPI_GETWORKAREA
        return new int[] { GetSystemMetrics(0), GetSystemMetrics(1) };
    }

    public static int[] VirtualScreen() {
        return new int[] { GetSystemMetrics(76), GetSystemMetrics(77), GetSystemMetrics(78), GetSystemMetrics(79) };
    }

    // Moves a top-level window just past the right edge of the whole desktop, without
    // activating it. It stays restored, so it keeps drawing and PrintWindow still works.
    public static bool Park(IntPtr h) {
        var v = VirtualScreen();
        return SetWindowPos(h, IntPtr.Zero, v[0] + v[2] + 100, v[1], 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    public static bool Show(IntPtr h, int x, int y) {
        return SetWindowPos(h, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    public static bool Size(IntPtr h, int w, int ht) {
        return SetWindowPos(h, IntPtr.Zero, 0, 0, w, ht, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    public static bool Close(IntPtr h) { return PostMessage(h, 0x10, IntPtr.Zero, IntPtr.Zero); }   // WM_CLOSE

    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr h, int id);

    static bool Gone(IntPtr h, int ms) {
        for (int i = 0; i < ms / 100; i++) { Thread.Sleep(100); if (!IsWindow(h) || !IsWindowVisible(h)) return true; }
        return false;
    }

    // Closes a message box. WM_CLOSE acts as Cancel; a Yes/No box has no Cancel and
    // ignores it, so then press No (the safer answer in a review), else the default button.
    public static string Dismiss(IntPtr h) {
        PostMessage(h, 0x10, IntPtr.Zero, IntPtr.Zero);
        if (Gone(h, 1000)) return "closed";
        if (Cls(h) == "#32770") {
            int id = 0; string what = "";
            if (GetDlgItem(h, 7) != IntPtr.Zero) { id = 7; what = "pressed No"; }   // IDNO
            else {
                long r = SendMessage(h, 0x400, IntPtr.Zero, IntPtr.Zero).ToInt64();   // DM_GETDEFID
                if (((r >> 16) & 0xFFFF) == 0x534B) { id = (int)(r & 0xFFFF); what = "pressed the default button"; }
            }
            if (id != 0) {
                PostMessage(h, 0x111, new IntPtr(id), IntPtr.Zero);   // WM_COMMAND
                if (Gone(h, 1000)) return what;
            }
        }
        return "still open";
    }

    // Visible top-level windows of a process, skipping tooltips and empty windows.
    public static List<IntPtr> TopWindows(uint pid) {
        var r = new List<IntPtr>();
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p == pid && IsWindowVisible(h)) {
                var c = Cls(h); var rc = Rect(h);
                if (c != "tooltips_class32" && rc[2] > 0 && rc[3] > 0) r.Add(h);
            }
            return true;
        }, IntPtr.Zero);
        return r;
    }

    // Draws the window's top-level ancestor with PrintWindow, crops to the window itself,
    // then to an optional sub-rectangle (relative to the window, physical pixels; cw = 0
    // for none), enlarges by zoom, and saves a PNG. Returns width, height and whether
    // the image came back a single colour (a locked session or minimized window).
    public static int[] Save(IntPtr h, string path, int cx, int cy, int cw, int ch, int zoom) {
        IntPtr root = Root(h);
        if (root == IntPtr.Zero) root = h;
        if (IsIconic(root)) throw new InvalidOperationException("The window is minimized, so it can't be drawn. Restore it first.");
        var rr = Rect(root); var wr = Rect(h);
        if (rr[2] <= 0 || rr[3] <= 0) throw new InvalidOperationException("The window has no size.");
        using (var full = new Bitmap(rr[2], rr[3], PixelFormat.Format32bppArgb)) {
            using (var g = Graphics.FromImage(full)) {
                IntPtr dc = g.GetHdc();
                bool ok;
                try { ok = PrintWindow(root, dc, PW_RENDERFULLCONTENT); } finally { g.ReleaseHdc(dc); }
                if (!ok) throw new InvalidOperationException("PrintWindow failed.");
            }
            var crop = new Rectangle(wr[0] - rr[0], wr[1] - rr[1], wr[2], wr[3]);
            if (cw > 0 && ch > 0) crop = new Rectangle(crop.X + cx, crop.Y + cy, cw, ch);
            crop.Intersect(new Rectangle(0, 0, rr[2], rr[3]));
            if (crop.Width <= 0 || crop.Height <= 0) throw new InvalidOperationException("The crop lies outside the window.");
            if (zoom < 1) zoom = 1;
            using (var outBmp = new Bitmap(crop.Width * zoom, crop.Height * zoom, PixelFormat.Format32bppArgb)) {
                using (var g = Graphics.FromImage(outBmp)) {
                    g.InterpolationMode = InterpolationMode.NearestNeighbor;   // keep pixel edges sharp for spacing checks
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.DrawImage(full, new Rectangle(0, 0, outBmp.Width, outBmp.Height), crop, GraphicsUnit.Pixel);
                }
                outBmp.Save(path, ImageFormat.Png);
                return new int[] { outBmp.Width, outBmp.Height, Blank(full, crop) ? 1 : 0 };
            }
        }
    }

    static bool Blank(Bitmap b, Rectangle r) {
        int first = b.GetPixel(r.X, r.Y).ToArgb();
        for (int i = 0; i < 24; i++)
            for (int j = 0; j < 24; j++)
                if (b.GetPixel(r.X + r.Width * i / 24, r.Y + r.Height * j / 24).ToArgb() != first) return false;
        return true;
    }
}

// Keeps every top-level window of the Access process off the desktop for the whole
// session (off-screen mode), not just during a call: a startup popup such as a sign-in
// form, a dialog opened later, or one that re-centres itself on a timer. A WinEvent
// hook moves a window the moment it is shown or moved onto the desktop; a sweep every
// 50 ms catches anything the hook missed.
public class KitReviewKeeper {
    delegate void WinEventProc(IntPtr hook, uint ev, IntPtr h, int idObject, int idChild, uint thread, uint time);
    [StructLayout(LayoutKind.Sequential)] struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }
    [DllImport("user32.dll")] static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr mod, WinEventProc cb, uint pid, uint tid, uint flags);
    [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] static extern bool PeekMessage(out MSG m, IntPtr h, uint min, uint max, uint remove);
    [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
    [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG m);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint flags);

    readonly uint pid; volatile bool stop; Thread thread;
    WinEventProc proc;   // held so the delegate isn't collected while the hook is live
    public int Moved;

    public KitReviewKeeper(uint pid) { this.pid = pid; }
    public void Start() { thread = new Thread(Loop); thread.IsBackground = true; thread.Start(); }
    public void Stop() { stop = true; if (thread != null) thread.Join(2000); }

    static bool OnDesktop(IntPtr h) {
        var v = KitReviewWin.VirtualScreen(); var r = KitReviewWin.Rect(h);
        return r[0] < v[0] + v[2] && r[0] + r[2] > v[0] && r[1] < v[1] + v[3] && r[1] + r[3] > v[1];
    }

    void Check(IntPtr h) {
        if (!KitReviewWin.IsIconic(h) && OnDesktop(h)) { KitReviewWin.Park(h); Interlocked.Increment(ref Moved); }
    }

    void Loop() {
        proc = (hk, ev, h, idObject, idChild, t, time) => {
            if (idObject != 0 || idChild != 0 || h == IntPtr.Zero) return;   // the window itself, not a part of it
            if (GetAncestor(h, 2) != h) return;                               // top-level windows only
            try { Check(h); } catch { }
        };
        // EVENT_OBJECT_SHOW (0x8002) through EVENT_OBJECT_LOCATIONCHANGE (0x800B), for the
        // Access process only; out of context, so delivered to this thread's message loop.
        IntPtr hook = SetWinEventHook(0x8002, 0x800B, IntPtr.Zero, proc, pid, 0, 0x2);   // WINEVENT_SKIPOWNPROCESS
        var sw = Stopwatch.StartNew();
        MSG m;
        while (!stop) {
            while (PeekMessage(out m, IntPtr.Zero, 0, 0, 1)) { TranslateMessage(ref m); DispatchMessage(ref m); }   // PM_REMOVE
            if (sw.ElapsedMilliseconds >= 50) {
                sw.Reset(); sw.Start();
                try { foreach (var h in KitReviewWin.TopWindows(pid)) Check(h); } catch { }
            }
            Thread.Sleep(5);
        }
        if (hook != IntPtr.Zero) UnhookWinEvent(hook);
    }
}

// Watches one Access process while a COM call runs. Each new top-level window is parked
// off the desktop (in off-screen mode), captured and described. VBA and Office message
// boxes are closed at once so the blocked call returns; any other window (a kit dialog
// form opened modally) is closed if the call is still blocked a few seconds later
// (never, with modalSec 0, as while a front end starts up). If the call outlasts the
// timeout, Access is ended.
public class KitReviewGuard {
    readonly uint pid; readonly string dir; readonly bool park; readonly int timeoutMs, modalMs;
    readonly HashSet<IntPtr> known = new HashSet<IntPtr>();
    readonly Dictionary<IntPtr, long> pending = new Dictionary<IntPtr, long>();
    public readonly List<string> Events = new List<string>();
    public bool Killed;
    volatile bool stop; Thread thread; int n;

    public KitReviewGuard(uint pid, string dir, bool park, int timeoutSec, int modalSec) {
        this.pid = pid; this.dir = dir; this.park = park; timeoutMs = timeoutSec * 1000; modalMs = modalSec * 1000;
        foreach (var h in KitReviewWin.TopWindows(pid)) known.Add(h);
    }

    public void Start() { thread = new Thread(Loop); thread.IsBackground = true; thread.Start(); }

    void Note(string s) { lock (Events) Events.Add(s); }

    void Loop() {
        var sw = Stopwatch.StartNew();
        while (!stop) {
            Process proc;
            try { proc = Process.GetProcessById((int)pid); } catch { return; }
            if (proc.HasExited) { Note("[exited] Access closed during the call"); return; }
            foreach (var h in KitReviewWin.TopWindows(pid)) {
                if (!known.Add(h)) continue;
                if (park) KitReviewWin.Park(h);
                Thread.Sleep(600);   // let it finish drawing
                var cls = KitReviewWin.Cls(h);
                string shot = System.IO.Path.Combine(dir, "dialog-" + DateTime.Now.ToString("HHmmss") + "-" + (++n) + ".png");
                try { KitReviewWin.Save(h, shot, 0, 0, 0, 0, 1); } catch (Exception e) { shot = "(not captured: " + e.Message + ")"; }
                bool msgBox = cls == "#32770" || cls == "NUIDialog";
                string line = (msgBox ? "[message box] " : "[window " + cls + "] ") + KitReviewWin.Describe(h) + " | capture " + shot;
                if (msgBox) line += " | " + KitReviewWin.Dismiss(h);
                else if (modalMs > 0) pending[h] = sw.ElapsedMilliseconds;
                Note(line);
            }
            foreach (var kv in new List<KeyValuePair<IntPtr, long>>(pending)) {
                if (sw.ElapsedMilliseconds - kv.Value < modalMs) continue;
                pending.Remove(kv.Key);
                if (KitReviewWin.TopWindows(pid).Contains(kv.Key)) {
                    Note("[closed] window " + KitReviewWin.Cls(kv.Key) + " was still blocking the call after " + (modalMs / 1000) + "s");
                    KitReviewWin.Close(kv.Key);
                }
            }
            if (timeoutMs > 0 && sw.ElapsedMilliseconds > timeoutMs) {
                Note("[timeout] Access still busy after " + (timeoutMs / 1000) + "s; ended it");
                try { proc.Kill(); Killed = true; } catch { }
                return;
            }
            Thread.Sleep(200);
        }
    }

    public string[] Stop() {
        stop = true;
        if (thread != null) thread.Join(3000);
        lock (Events) return Events.ToArray();
    }
}
'@ -ReferencedAssemblies System.Drawing, UIAutomationClient, UIAutomationTypes
    [KitReviewWin]::DpiAware()
}

$script:Review = $null

# An environment setting, or $null when it is unset or empty. A host that fills
# settings from a template (a desktop extension's ${user_config.x}) can pass the
# placeholder through unexpanded when the user left it blank; that counts as unset.
function Get-ReviewSetting([string]$Name) {
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v) -or $v -match '^\$\{.*\}$') { return $null }
    $v.Trim()
}

# Palettes set_palette can switch to, as name = value passed to the palette procedure.
# The first is the default: the app opens on it and close_app goes back to it. The
# default list is access-modern-ui's (SetPalette 1, 2, 6).
function Get-ReviewPalettes {
    $spec = Get-ReviewSetting 'ACCESS_REVIEW_PALETTES'
    if (-not $spec) { $spec = 'Navy=1;Midnight=2;Teal Navy=6' }
    $list = [ordered]@{}
    foreach ($pair in $spec -split ';') {
        if (-not $pair.Trim()) { continue }
        $name, $value = $pair -split '=', 2
        $n = 0
        if (-not $name.Trim() -or -not [int]::TryParse(([string]$value).Trim(), [ref]$n)) {
            throw "ACCESS_REVIEW_PALETTES: '$pair' is not name=number."
        }
        $list[$name.Trim()] = $n
    }
    if (-not $list.Count) { throw 'ACCESS_REVIEW_PALETTES lists no palettes.' }
    $list
}

# A bad setting falls back to the default list rather than stopping the server, which
# a desktop app would only show as a server that failed to start. The problem is
# reported in the server's instructions and in every set_palette result.
$script:PaletteWarning = $null
try { $script:Palettes = Get-ReviewPalettes }
catch {
    $script:PaletteWarning = "$($_.Exception.Message) Using the default palettes instead."
    $env:ACCESS_REVIEW_PALETTES = $null
    $script:Palettes = Get-ReviewPalettes
}
$script:DefaultPalette = @($script:Palettes.Keys)[0]
$script:PaletteProcedure = Get-ReviewSetting 'ACCESS_REVIEW_PALETTE_PROCEDURE'
if (-not $script:PaletteProcedure) { $script:PaletteProcedure = 'SetPalette' }
$script:TwipsPerPx = 15   # kit layout values are pixels at 96 dpi

function Assert-ReviewOpen {
    if (-not $script:Review) { throw 'Access is not open. Call Open-ReviewApp first.' }
    if ($script:Review.Process.HasExited) { $script:Review = $null; throw 'Access has closed (it may have been ended after a hang). Open it again.' }
}

# Runs a script block against Access under a KitReviewGuard. Returns the block's result
# plus the dialogs seen, and throws with those dialogs listed if the call failed.
function Invoke-ReviewGuarded {
    param([Parameter(Mandatory)][scriptblock]$Body, [int]$TimeoutSec = 120, [int]$ModalSec = 4)
    Assert-ReviewOpen
    $r = $script:Review
    $guard = New-Object KitReviewGuard -ArgumentList ([uint32]$r.Pid), $r.Out, ($r.Mode -eq 'offscreen'), $TimeoutSec, $ModalSec
    $guard.Start()
    $value = $null; $failure = $null
    try { $value = & $Body } catch { $failure = $_.Exception.Message }
    $events = @($guard.Stop())
    $r.Dialogs += $events
    if ($guard.Killed) { $script:Review = $null }
    if ($failure) {
        $msg = "Access call failed: $failure"
        if ($events.Count) { $msg += "`nDialogs: " + ($events -join "`n") }
        throw $msg
    }
    [pscustomobject]@{ Value = $value; Dialogs = $events }
}

# PowerShell's COM adapter can't read Access Form, Control and Section objects (their
# type information fails, so $form.Hwnd comes back empty and InsideWidth "cannot be
# found"). Everything on a form is read with Application.Eval and set late-bound with
# InvokeMember.
function Get-FormValue([string]$Name, [string]$Expr) {
    $script:Review.Acc.Eval("Forms![$Name]$Expr")
}

function Find-FormValue([string]$Name, [string]$Expr) {
    try { Get-FormValue $Name $Expr } catch { $null }
}

function Assert-FormOpen([string]$Name) {
    Assert-ReviewOpen
    if ($Name -match '[\[\]!"]') { throw "Unsupported form name '$Name'." }
    try { $null = Get-FormValue $Name '.Name' } catch { throw "Form '$Name' is not open." }
}

function Get-FormHwnd([string]$Name) { [IntPtr][long](Get-FormValue $Name '.Hwnd') }

# Heights of the sections form view shows (header, detail, footer), in twips.
function Get-FormDesignHeight([string]$Name) {
    $h = 0
    foreach ($i in 1, 0, 2) {
        if (Find-FormValue $Name ".Section($i).Visible") { $h += [int](Get-FormValue $Name ".Section($i).Height") }
    }
    $h
}

function Release-Com($o) {
    if ($o -is [__ComObject]) { while ([Runtime.InteropServices.Marshal]::ReleaseComObject($o) -gt 0) {} }
}

# Sets InsideWidth/InsideHeight (twips) through InvokeMember; if that fails too, sizes
# the form window with DoCmd.MoveSize plus its frame (window size minus inside size).
function Set-FormInsideSize([string]$Name, [int]$Width, [int]$Height) {
    $f = $null
    try {
        $f = $script:Review.Acc.Forms.Item($Name)
        $set = [Reflection.BindingFlags]::SetProperty
        [void][__ComObject].InvokeMember('InsideWidth', $set, $null, $f, @($Width))
        [void][__ComObject].InvokeMember('InsideHeight', $set, $null, $f, @($Height))
    } catch {
        $fw = [int](Get-FormValue $Name '.WindowWidth') - [int](Get-FormValue $Name '.InsideWidth')
        $fh = [int](Get-FormValue $Name '.WindowHeight') - [int](Get-FormValue $Name '.InsideHeight')
        $script:Review.Acc.DoCmd.MoveSize([Type]::Missing, [Type]::Missing, $Width + $fw, $Height + $fh)
    } finally { Release-Com $f }
}

# Lock files shared by reviewers on one machine (ACCESS_REVIEW_LOCKS: full paths,
# separated by ';'). Takes all of them or none, then re-reads to catch a race.
function Enter-ReviewLock {
    param([Parameter(Mandatory)][string]$Owner, [string[]]$Paths = @(((Get-ReviewSetting 'ACCESS_REVIEW_LOCKS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }))
    if (-not $Paths.Count) { return [pscustomobject]@{ Held = $true; Note = 'No lock files configured (ACCESS_REVIEW_LOCKS); nothing to take.' } }
    $busy = @($Paths | Where-Object { Test-Path -LiteralPath $_ })
    if ($busy.Count) {
        return [pscustomobject]@{ Held = $false; Note = 'Busy: ' + (($busy | ForEach-Object { "$_ = " + ((Get-Content -LiteralPath $_ -Raw) -replace '\s+', ' ').Trim() }) -join '; ') }
    }
    $stamp = "$Owner | $(Get-Date -Format s) | access-review"
    foreach ($p in $Paths) { Set-Content -LiteralPath $p -Value $stamp -Encoding ASCII }
    Start-Sleep -Seconds 3
    $lost = @($Paths | Where-Object { -not (Test-Path -LiteralPath $_) -or ((Get-Content -LiteralPath $_ -Raw).Trim() -ne $stamp) })
    if ($lost.Count) {
        # Lost the race: take back the stamps that are still ours so no stale lock is left.
        foreach ($p in $Paths) { if ((Test-Path -LiteralPath $p) -and ((Get-Content -LiteralPath $p -Raw).Trim() -eq $stamp)) { Remove-Item -LiteralPath $p } }
        return [pscustomobject]@{ Held = $false; Note = 'Another reviewer took the lock at the same moment: ' + ($lost -join '; ') + '. Wait and try again.' }
    }
    [pscustomobject]@{ Held = $true; Note = "Holding: $($Paths -join '; ')" }
}

# Deletes only lock files this owner wrote.
function Exit-ReviewLock {
    param([Parameter(Mandatory)][string]$Owner, [string[]]$Paths = @(((Get-ReviewSetting 'ACCESS_REVIEW_LOCKS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }))
    $released = @()
    foreach ($p in $Paths) {
        if ((Test-Path -LiteralPath $p) -and ((Get-Content -LiteralPath $p -Raw) -like "$Owner |*")) { Remove-Item -LiteralPath $p; $released += $p }
    }
    [pscustomobject]@{ Released = $released }
}

# True for a fully qualified path: C:\x, a UNC path \\server\share\x, or /x off Windows.
# IsPathRooted alone also accepts C:x (relative to drive C's current folder) and \x
# (relative to the current drive), which still depend on where the process is.
function Test-FullPath([string]$Path) {
    $Path -match '^[A-Za-z]:[\\/]' -or $Path -match '^[\\/]{2}[^\\/]' -or ([IO.Path]::DirectorySeparatorChar -eq '/' -and $Path.StartsWith('/'))
}

# Folder for captures and snapshots: ACCESS_REVIEW_OUTPUT, or %LOCALAPPDATA%\AccessReview.
# A relative setting would resolve against whichever folder the server or Access happens
# to be in, so it is refused.
function Get-ReviewOutputBase {
    $base = Get-ReviewSetting 'ACCESS_REVIEW_OUTPUT'
    if (-not $base) { return (Join-Path $env:LOCALAPPDATA 'AccessReview') }
    if (-not (Test-FullPath $base)) { throw "ACCESS_REVIEW_OUTPUT: '$base' is not a full path." }
    [IO.Path]::GetFullPath($base)
}

# True when Path is FolderPath itself or inside it. Both are full paths; the comparison
# ignores case and a trailing separator, and a folder never matches a sibling that
# merely starts with the same name (C:\Review vs C:\Review2).
function Test-PathInFolder([string]$Path, [string]$FolderPath) {
    $f = [IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $Path.StartsWith($f, [StringComparison]::OrdinalIgnoreCase)
}

# Under automation Access runs a front end's startup code (AutoExec, startup form)
# without the Enable Content bar, and the run tool calls any public procedure, so
# open_app refuses a file that came from the internet (its Zone.Identifier stream
# says zone 3 or 4) and, when ACCESS_REVIEW_ALLOWED_FOLDERS is set, a file outside
# those folders.
function Assert-ReviewFileAllowed([string]$Path) {
    $allowed = @(((Get-ReviewSetting 'ACCESS_REVIEW_ALLOWED_FOLDERS') -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($a in $allowed) {
        if (-not (Test-FullPath $a)) { throw "ACCESS_REVIEW_ALLOWED_FOLDERS: '$a' is not a full path." }
    }
    $inside = @($allowed | Where-Object { Test-PathInFolder $Path $_ })
    if ($allowed.Count -and -not $inside.Count) {
        throw "'$Path' is not in a folder this tool may open ($($allowed -join '; '), from ACCESS_REVIEW_ALLOWED_FOLDERS)."
    }
    # The check above compares names, so a junction or symbolic link inside an allowed
    # folder could lead outside it. Refuse a path that goes through one below the folder.
    # A hard link to a file elsewhere is not detected (Windows can't tell it from the
    # original); making one needs write access to the allowed folder, which already lets
    # someone put any front end there.
    foreach ($a in $inside) {
        $root = [IO.Path]::GetFullPath($a).TrimEnd('\', '/')
        for ($p = $Path; $p -and $p.Length -gt $root.Length; $p = [IO.Path]::GetDirectoryName($p)) {
            $link = (Get-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue).LinkType
            if ($link -in 'Junction', 'SymbolicLink') { throw "'$Path' goes through the $link '$p', which may lead outside $a, so it is not opened. Use the file's real location." }
        }
    }
    $zone = $null
    try { $zone = Get-Content -LiteralPath $Path -Stream Zone.Identifier -ErrorAction Stop } catch {}
    if ($zone -and (($zone -join "`n") -match '(?m)^ZoneId=([34])\s*$')) {
        throw "'$Path' was downloaded from the internet (Zone.Identifier ZoneId=$($Matches[1])), so its startup code is not trusted. If you trust it, unblock it first (file Properties, Unblock) and open it again."
    }
}

# Starts a separate Access instance and opens an existing front end in it. Off-screen
# mode parks every Access window past the desktop edge; on-screen mode (for a dedicated
# review machine) puts the Access window at the top left of the main display.
function Open-ReviewApp {
    param([Parameter(Mandatory)][string]$Path, [ValidateSet('offscreen', 'onscreen')][string]$Mode = 'offscreen',
          [string]$OutputDirectory, [int]$AppWidth = 1920, [int]$AppHeight = 1200)
    if ($script:Review -and -not $script:Review.Process.HasExited) { throw "Access is already open on $($script:Review.Path). Close it first." }
    $Path = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue).ProviderPath
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Open only an existing, already trusted front end; this tool never creates one.' }
    Assert-ReviewFileAllowed $Path
    if ($OutputDirectory -and -not (Test-FullPath $OutputDirectory)) { throw "output_dir '$OutputDirectory' is not a full path." }
    if (-not $OutputDirectory) {
        $OutputDirectory = Join-Path (Get-ReviewOutputBase) (Get-Date -Format 'yyyyMMdd-HHmmss')
    }
    New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

    $acc = New-Object -ComObject Access.Application
    $hwnd = [IntPtr][long]$acc.hWndAccessApp()
    $proc = Get-Process -Id ([KitReviewWin]::PidOf($hwnd))
    $v = [KitReviewWin]::VirtualScreen()
    # A window can't be larger than the desktop, and a form bigger than the Access window is clipped.
    $w = [Math]::Min($AppWidth, $v[2]); $h = [Math]::Min($AppHeight, $v[3])
    $place = {
        [void][KitReviewWin]::Size($hwnd, $w, $h)
        if ($Mode -eq 'offscreen') { [void][KitReviewWin]::Park($hwnd) } else { [void][KitReviewWin]::Show($hwnd, 0, 0) }
    }
    $keeper = $null
    if ($Mode -eq 'offscreen') {
        $keeper = New-Object KitReviewKeeper -ArgumentList ([uint32]$proc.Id)
        $keeper.Start()
    }
    & $place                     # while still hidden, so it never appears on the desktop
    $acc.Visible = $true         # PrintWindow needs a shown (not minimized) window
    & $place
    try { $acc.UserControl = $false } catch {}   # so Access ends with this script

    $script:Review = [pscustomobject]@{
        Acc = $acc; Hwnd = $hwnd; Pid = $proc.Id; Process = $proc; Path = $Path; Mode = $Mode; Out = $OutputDirectory
        Palette = $script:DefaultPalette; Dialogs = @(); LogMark = $null; FileMark = $null; Sizes = @{}; Keeper = $keeper
    }
    try {
        # ModalSec 0: a splash or sign-in form shown at startup is left open.
        $open = Invoke-ReviewGuarded { $acc.OpenCurrentDatabase($Path, $false) } -TimeoutSec 180 -ModalSec 0
    } catch { Close-ReviewApp | Out-Null; throw }
    & $place
    [pscustomobject]@{
        Path = $Path; Pid = $proc.Id; Mode = $Mode; Output = $OutputDirectory; AccessWindow = "$($w)x$($h)"
        OpenForms = @(Get-ReviewOpenForms); Dialogs = $open.Dialogs
    }
}

function Get-ReviewOpenForms {
    $acc = $script:Review.Acc
    $n = [int]$acc.Eval('Forms.Count')
    for ($i = 0; $i -lt $n; $i++) { [string]$acc.Eval("Forms($i).Name") }
}

# Closes the database without saving design changes, quits, and makes sure the process
# is gone. Resets the palette to the default first if it was changed.
function Close-ReviewApp {
    if (-not $script:Review) { return [pscustomobject]@{ Closed = $true; Note = 'Access was not open.' } }
    $r = $script:Review
    $notes = @()
    if (-not $r.Process.HasExited) {
        $d = $script:DefaultPalette
        if ($r.Palette -ne $d) { try { $null = Set-ReviewPalette $d -NoReopen } catch { $notes += "Could not reset the palette to ${d}: $($_.Exception.Message)" } }
        try { $r.Acc.CloseCurrentDatabase() } catch {}
        try { $r.Acc.Quit(2) } catch {}   # acQuitSaveNone
    }
    # Access stays alive while this process holds any reference to it, including
    # temporaries PowerShell made along the way, so release them all before waiting.
    try { Release-Com $r.Acc } catch {}
    $r.Acc = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    if (-not $r.Process.WaitForExit(15000)) { $r.Process.Kill(); $notes += 'Access did not quit within 15s; ended it.' }
    if ($r.Keeper) { $r.Keeper.Stop(); $notes += "Windows moved off the desktop during the session: $($r.Keeper.Moved)." }
    $script:Review = $null
    [pscustomobject]@{ Closed = $true; Note = ($notes -join ' ') }
}

# Compiles and saves all modules. A failed compile under automation shows only a
# generic alert (captured by the guard); tests\native-smoke.ps1 names the line.
function Test-ReviewCompile {
    $r = Invoke-ReviewGuarded { $script:Review.Acc.DoCmd.RunCommand(126); $script:Review.Acc.IsCompiled }
    [pscustomobject]@{ Compiled = [bool]$r.Value; Dialogs = $r.Dialogs }
}

# Runs a public VBA procedure or function (Application.Run), e.g. an app's sign-in.
function Invoke-ReviewRun {
    param([Parameter(Mandatory)][string]$Name, [object[]]$Arguments = @(), [int]$TimeoutSec = 120)
    if ($Arguments.Count -gt 8) { throw 'At most 8 arguments.' }
    # Run passes its arguments ByRef, so each gets its own reference holder (a [ref] to
    # the loop variable would make every argument the same one) and output arguments
    # come back. JSON whole numbers arrive as Int64, which a VBA Long rejects (type
    # mismatch), so they become Int32, or Double when too large; decimals become Double.
    $refs = @(foreach ($a in $Arguments) {
        $v = $a
        if ($v -is [long]) { $v = if ($v -ge [int]::MinValue -and $v -le [int]::MaxValue) { [int]$v } else { [double]$v } }
        elseif ($v -is [decimal]) { $v = [double]$v }
        [Management.Automation.PSReference]::new($v)
    })
    $r = Invoke-ReviewGuarded -TimeoutSec $TimeoutSec {
        $acc = $script:Review.Acc
        switch ($refs.Count) {
            0 { $acc.Run($Name) }
            1 { $acc.Run($Name, $refs[0]) }
            2 { $acc.Run($Name, $refs[0], $refs[1]) }
            3 { $acc.Run($Name, $refs[0], $refs[1], $refs[2]) }
            4 { $acc.Run($Name, $refs[0], $refs[1], $refs[2], $refs[3]) }
            5 { $acc.Run($Name, $refs[0], $refs[1], $refs[2], $refs[3], $refs[4]) }
            6 { $acc.Run($Name, $refs[0], $refs[1], $refs[2], $refs[3], $refs[4], $refs[5]) }
            7 { $acc.Run($Name, $refs[0], $refs[1], $refs[2], $refs[3], $refs[4], $refs[5], $refs[6]) }
            8 { $acc.Run($Name, $refs[0], $refs[1], $refs[2], $refs[3], $refs[4], $refs[5], $refs[6], $refs[7]) }
        }
    }
    [pscustomobject]@{ Result = $r.Value; ArgumentsAfter = @($refs | ForEach-Object { $_.Value }); Dialogs = $r.Dialogs }
}

# Waits (up to 5 s) until a form's timer has stopped, so work the kit queues on a timer
# after opening it (pills, progress bars) has drawn before anyone captures it.
function Wait-ReviewFormIdle([string]$Name) {
    for ($k = 0; $k -lt 25; $k++) {
        $t = Find-FormValue $Name '.TimerInterval'
        if ($null -eq $t -or [int]$t -eq 0) { break }
        Start-Sleep -Milliseconds 200
    }
    Start-Sleep -Milliseconds 300
}

# True when a late-bound call failed because the object has no member of that name,
# rather than because the member itself failed: DISP_E_UNKNOWNNAME (which .NET reports
# as a MissingMemberException), or Access error 2465 (0x800A09A1), which an Access form
# raises for a name it doesn't have.
function Test-UnknownMember($err) {
    for ($e = $err; $e; $e = $e.InnerException) {
        if ($e -is [MissingMemberException] -or $e.HResult -eq -2147352570 -or $e.HResult -eq -2146825823) { return $true }
    }
    $false
}

# Calls a form's own public Reopen method (an app shell form may have one that reopens
# the form on a timer without running its sign-out on close), passing the current page when the
# form exposes PageKey. Returns 'reopened', or 'none' when the form has no Reopen method
# and is safe to close and reopen. Throws when the form has Reopen but the page can't be
# read or the call fails: closing such a form would sign the user out.
# Form-module members are read with InvokeMember on the form object; Eval can't reach
# a form module's Property Get.
function Invoke-FormReopen([string]$Name) {
    $f = $null
    $bf = [Reflection.BindingFlags]
    try {
        $f = $script:Review.Acc.Forms.Item($Name)
        $hasPage = $true; $page = $null
        for ($k = 0; $k -lt 10; $k++) {   # a navigation still queued on the form's timer leaves PageKey empty for a moment
            try { $page = [string][__ComObject].InvokeMember('PageKey', $bf::GetProperty, $null, $f, $null) }
            catch { if (Test-UnknownMember $_.Exception) { $hasPage = $false; break } else { throw "Form '$Name': could not read PageKey: $($_.Exception.Message)" } }
            if ($page) { break }
            Start-Sleep -Milliseconds 200
        }
        if ($hasPage -and -not $page) { throw "Form '$Name' has PageKey but it stayed empty, so it was not reopened (closing it would sign the user out)." }
        $callArgs = if ($hasPage) { @($page) } else { @() }
        try { [void][__ComObject].InvokeMember('Reopen', $bf::InvokeMethod, $null, $f, $callArgs) }
        catch {
            if (Test-UnknownMember $_.Exception) {
                if ($hasPage) { throw "Form '$Name' has PageKey but no Reopen method; left open rather than closed." }
                return 'none'
            }
            throw "Form '$Name': Reopen failed: $($_.Exception.Message)"
        }
        'reopened'
    } finally { Release-Com $f }
}

# Waits (up to 6 s) for a form that is closing and reopening on a timer to be open
# again with its timer stopped.
function Wait-ReviewFormReopened([string]$Name) {
    Start-Sleep -Milliseconds 300
    for ($k = 0; $k -lt 30; $k++) {
        $t = Find-FormValue $Name '.TimerInterval'
        if ($null -ne $t -and [int]$t -eq 0) { break }
        Start-Sleep -Milliseconds 200
    }
    Start-Sleep -Milliseconds 300
    if ($null -eq (Find-FormValue $Name '.Name')) { throw "Form '$Name' did not reopen after the palette switch." }
}

function Restore-ReviewFormSize([string]$Name) {
    $want = $script:Review.Sizes[$Name]
    $p = @{ Name = $Name }
    if ($want) { $p.WidthPx = $want.W; $p.HeightPx = $want.H; $p.Maximize = [bool]$want.Max }
    Set-ReviewFormSize @p
}

# Switches the palette, then reopens every open form, as the app's own Appearance switch
# does: a form left open keeps part of the old drawing (in Midnight, pills and progress
# bars went missing). A form with a public Reopen method (an app shell such as a
# frmMain that signs the user out on close) reopens itself, which keeps the user signed in and on the same page; closing
# it instead would run its sign-out. Other forms (dialogs) are closed first and reopened
# last with their OpenArgs, so they come back on top of the reopened shell. Each form
# returns to the size last asked for.
function Set-ReviewPalette {
    param([Parameter(Mandatory)][string]$Name, [switch]$NoReopen)
    Assert-ReviewOpen
    if (-not $script:Palettes.Contains($Name)) { throw "Unknown palette '$Name'. Configured: $(@($script:Palettes.Keys) -join ', ') (ACCESS_REVIEW_PALETTES).$(if ($script:PaletteWarning) { ' ' + $script:PaletteWarning })" }
    $i = [int]$script:Palettes[$Name]
    $proc = $script:PaletteProcedure
    $acc = $script:Review.Acc
    $open = @(foreach ($f in @(Get-ReviewOpenForms)) { [pscustomobject]@{ Name = $f; OpenArgs = [string](Find-FormValue $f '.OpenArgs') } })
    try { $r = Invoke-ReviewGuarded { $acc.Run($proc, [ref]$i) } }
    catch {
        # Access error 2517: Run can't find the procedure. Anything else is a real failure
        # inside it (or a timeout) and is reported as it is.
        if ($_.Exception.Message -match '2517|find the procedure') { throw "Could not run the palette procedure '$proc' (ACCESS_REVIEW_PALETTE_PROCEDURE); the front end needs a public Sub $proc(ByVal i As Long). $($_.Exception.Message)" }
        throw "The palette procedure '$proc' failed: $($_.Exception.Message)"
    }
    $script:Review.Palette = $Name
    $dialogs = @($r.Dialogs); $reopened = @()
    if (-not $NoReopen) {
        $selfReopen = @(); $plain = @()
        foreach ($f in $open) {
            $d = Invoke-ReviewGuarded { Invoke-FormReopen $f.Name }
            $dialogs += $d.Dialogs
            if ($d.Value -eq 'reopened') { $selfReopen += $f } else { $plain += $f }   # throws, never closes, if a Reopen form cannot reopen itself
        }
        foreach ($f in $plain) {
            $d = Invoke-ReviewGuarded { $acc.DoCmd.Close(2, $f.Name, 2) }   # acForm, acSaveNo
            $dialogs += $d.Dialogs
        }
        foreach ($f in $selfReopen) {
            Wait-ReviewFormReopened $f.Name
            $o = Restore-ReviewFormSize $f.Name
            Wait-ReviewFormIdle $f.Name
            $dialogs += $o.Dialogs
            $reopened += "$($f.Name) $($o.InsidePx) (own Reopen)"
        }
        foreach ($f in $plain) {
            $want = $script:Review.Sizes[$f.Name]
            $p = @{ Name = $f.Name; OpenArgs = $f.OpenArgs }
            if ($want) { $p.WidthPx = $want.W; $p.HeightPx = $want.H; $p.Maximize = [bool]$want.Max }
            $o = Open-ReviewForm @p
            $dialogs += $o.Dialogs
            $reopened += "$($f.Name) $($o.InsidePx)"
        }
    }
    [pscustomobject]@{ Palette = $Name; Index = $i; Reopened = $reopened; Dialogs = $dialogs; Warning = $script:PaletteWarning }
}

# Size and position facts for one open form, in pixels at 96 dpi (twips / 15) unless
# named Phys (physical screen pixels).
function Get-ReviewFormSize([string]$Name) {
    Assert-FormOpen $Name
    $px = { param($twips) [int]([int]$twips / $script:TwipsPerPx) }
    $hwnd = Get-FormHwnd $Name
    $root = [KitReviewWin]::Root($hwnd)
    $fr = [KitReviewWin]::Rect($hwnd); $ar = [KitReviewWin]::Rect($script:Review.Hwnd)
    $popup = $root -ne $script:Review.Hwnd
    $clipped = (-not $popup) -and ($fr[0] -lt $ar[0] -or $fr[1] -lt $ar[1] -or ($fr[0] + $fr[2]) -gt ($ar[0] + $ar[2]) -or ($fr[1] + $fr[3]) -gt ($ar[1] + $ar[3]))
    [pscustomobject]@{
        Form = $Name
        DesignPx = "$(& $px (Get-FormValue $Name '.Width'))x$(& $px (Get-FormDesignHeight $Name))"
        InsidePx = "$(& $px (Get-FormValue $Name '.InsideWidth'))x$(& $px (Get-FormValue $Name '.InsideHeight'))"
        WindowPhys = "$($fr[2])x$($fr[3])"
        Popup = $popup
        ClippedByAccessWindow = $clipped
    }
}

# Opens a form in form view and sizes it to its designed dimensions: the form's Width
# and the heights of its visible header, detail and footer. Pass WidthPx/HeightPx (96
# dpi pixels) for forms whose design size is set in code.
function Open-ReviewForm {
    param([Parameter(Mandatory)][string]$Name, [string]$OpenArgs, [int]$WidthPx, [int]$HeightPx, [switch]$Maximize)
    if ($Name -match '[\[\]!"]') { throw "Unsupported form name '$Name'." }
    $missing = [Type]::Missing
    $r = Invoke-ReviewGuarded {
        $acc = $script:Review.Acc
        $acc.DoCmd.OpenForm($Name, 0, $missing, $missing, $missing, 0, $(if ($OpenArgs) { $OpenArgs } else { $missing }))   # acNormal, acWindowNormal
    }
    $sized = Set-ReviewFormSize -Name $Name -WidthPx $WidthPx -HeightPx $HeightPx -Maximize:$Maximize
    Wait-ReviewFormIdle $Name
    $sized | Add-Member -NotePropertyName Dialogs -NotePropertyValue (@($r.Dialogs) + @($sized.Dialogs)) -Force
    $sized
}

# Resizes an open form: to its design size (no size given), to WidthPx x HeightPx, or
# maximized inside the Access window. Used for the small and maximized checks too.
function Set-ReviewFormSize {
    param([Parameter(Mandatory)][string]$Name, [int]$WidthPx, [int]$HeightPx, [switch]$Maximize)
    Assert-FormOpen $Name
    # Requested inside size in twips (design size unless given).
    $w = if ($WidthPx -gt 0) { $WidthPx * $script:TwipsPerPx } else { [int](Get-FormValue $Name '.Width') }
    $h = if ($HeightPx -gt 0) { $HeightPx * $script:TwipsPerPx } else { Get-FormDesignHeight $Name }
    $r = Invoke-ReviewGuarded {
        $acc = $script:Review.Acc
        $acc.DoCmd.SelectObject(2, $Name, $false)   # acForm
        if ($Maximize) { $acc.DoCmd.Maximize() } else {
            $acc.DoCmd.Restore()
            Set-FormInsideSize $Name $w $h
        }
    }
    $script:Review.Sizes[$Name] = @{ W = $WidthPx; H = $HeightPx; Max = [bool]$Maximize }
    $app = $script:Review.Hwnd
    $root = [KitReviewWin]::Root((Get-FormHwnd $Name))
    $note = $null
    if ($root -ne $app) {
        if ($script:Review.Mode -eq 'offscreen') { [void][KitReviewWin]::Park($root) }
    } elseif ($Maximize) {
        # The maximized check: the Access window at the size a maximized window gets on the
        # main display (its work area, without the taskbar).
        $ps = [KitReviewWin]::WorkArea()
        [void][KitReviewWin]::Size($app, $ps[0], $ps[1])
        Start-Sleep -Milliseconds 600
    } else {
        # A kit shell form (frmMain) fills the Access window and refits itself to it, so
        # setting its own size does nothing. Size the Access window instead until the
        # form's inside size matches, as reviewers do by hand.
        $scale = [KitReviewWin]::Scale($app)
        for ($k = 0; $k -lt 4; $k++) {
            $dx = $w - [int](Get-FormValue $Name '.InsideWidth'); $dy = $h - [int](Get-FormValue $Name '.InsideHeight')
            if ([Math]::Abs($dx) -le 15 -and [Math]::Abs($dy) -le 15) { break }   # within a pixel
            $ar = [KitReviewWin]::Rect($app)
            [void][KitReviewWin]::Size($app, $ar[2] + [int][Math]::Round($dx / 15 * $scale), $ar[3] + [int][Math]::Round($dy / 15 * $scale))
            Start-Sleep -Milliseconds 600   # let Access and the form's Resize handler catch up
        }
    }
    $size = Get-ReviewFormSize $Name
    if (-not $Maximize) {
        $wantPx = "$([int]($w / $script:TwipsPerPx))x$([int]($h / $script:TwipsPerPx))"
        $offW = [Math]::Abs($w - [int](Get-FormValue $Name '.InsideWidth')); $offH = [Math]::Abs($h - [int](Get-FormValue $Name '.InsideHeight'))
        if ($offW -gt 15 -or $offH -gt 15) {   # more than a pixel out
            $note = "Inside size is $($size.InsidePx), not the requested $wantPx. Captures are not at the designed size; the Access window may be limited by the desktop size."
        }
        $size | Add-Member -NotePropertyName SizeMatches -NotePropertyValue (-not $note)
    }
    if ($note) { $size | Add-Member -NotePropertyName Warning -NotePropertyValue $note }
    $size | Add-Member -NotePropertyName Requested -NotePropertyValue $(if ($Maximize) { 'maximized' } elseif ($WidthPx -or $HeightPx) { "${WidthPx}x${HeightPx} px" } else { 'design size' })
    $size | Add-Member -NotePropertyName Dialogs -NotePropertyValue $r.Dialogs
    $size
}

# Saves a PNG of an open form (or a crop of it, in physical pixels relative to the form
# window, enlarged by Zoom). The file name carries form, palette and label.
function Save-ReviewCapture {
    param([Parameter(Mandatory)][string]$Name, [string]$Label, [int[]]$Crop, [int]$Zoom = 1)
    Assert-FormOpen $Name
    $hwnd = Get-FormHwnd $Name
    if ($hwnd -eq [IntPtr]::Zero) { throw "Form '$Name' has no window handle." }
    if ($Label -and $Label -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') { throw "Capture label '$Label' must be letters, digits, '.', '_' or '-'." }
    if ($Name -match '[\\/:*?"<>|]') { throw "Form '$Name' has a character that can't be used in a file name." }
    $tag = ($script:Review.Palette -replace ' ', '').ToLower()
    $file = ($Name, $tag, $Label | Where-Object { $_ }) -join '-'
    if ($Crop) { $file += '-crop' }
    $path = Join-Path $script:Review.Out ($file + '.png')
    $c = if ($Crop) { if ($Crop.Count -ne 4) { throw 'Crop is x, y, width, height.' }; $Crop } else { @(0, 0, 0, 0) }
    # A form that was just opened or reopened may still be drawing (the kit queues its own
    # timer after SetPalette), so capture until two frames in a row are identical.
    $sha = [Security.Cryptography.SHA256]::Create()
    $prev = $null; $stable = $false
    for ($k = 0; $k -lt 10; $k++) {
        $res = [KitReviewWin]::Save($hwnd, $path, $c[0], $c[1], $c[2], $c[3], $Zoom)
        $hash = [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($path)))
        if ($hash -eq $prev) { $stable = $true; break }
        $prev = $hash
        Start-Sleep -Milliseconds 300
    }
    $size = Get-ReviewFormSize $Name
    [pscustomobject]@{
        Path = $path; ImagePx = "$($res[0])x$($res[1])"; Palette = $script:Review.Palette
        InsidePx = $size.InsidePx; ClippedByAccessWindow = $size.ClippedByAccessWindow
        Stable = $stable
        Warning = $(if ($res[2]) { 'The capture is one solid colour: the Windows session may be locked or disconnected, or the form has not drawn yet.' }
                    elseif (-not $stable) { 'The form was still changing after 3 s of captures (an animation or a timer); this frame may be mid-draw.' } else { $null })
    }
}

# Every control on an open form with its position in 96 dpi pixels (twips / 15) and the
# colours that matter for the kit, so gaps and the 16 px card inset can be computed.
function Get-ReviewControls {
    param([Parameter(Mandatory)][string]$Name)
    Assert-FormOpen $Name
    $sectionNames = @{ 0 = 'Detail'; 1 = 'Header'; 2 = 'Footer' }
    $n = [int](Get-FormValue $Name '.Controls.Count')
    $list = for ($i = 0; $i -lt $n; $i++) {
        $at = ".Controls($i)"
        $type = [int](Get-FormValue $Name "$at.ControlType")
        $o = [ordered]@{
            Name = [string](Get-FormValue $Name "$at.Name"); Type = $type
            Section = $sectionNames[[int](Find-FormValue $Name "$at.Section")]
            Left = [int]([int](Get-FormValue $Name "$at.Left") / 15); Top = [int]([int](Get-FormValue $Name "$at.Top") / 15)
            Width = [int]([int](Get-FormValue $Name "$at.Width") / 15); Height = [int]([int](Get-FormValue $Name "$at.Height") / 15)
            Visible = [bool](Get-FormValue $Name "$at.Visible")
        }
        foreach ($p in 'Caption', 'BackColor', 'ForeColor', 'BorderColor', 'BorderStyle', 'BackStyle', 'FontName', 'FontSize', 'TextAlign') {
            $v = Find-FormValue $Name "$at.$p"   # not every control type has every property
            if ($null -ne $v) { $o[$p] = $v }
        }
        if ($type -in 109, 110, 111) { $o['Value'] = [string](Find-FormValue $Name "$at.Value") }   # text box, list, combo
        [pscustomobject]$o
    }
    $sections = foreach ($i in 1, 0, 2) {
        $vis = Find-FormValue $Name ".Section($i).Visible"
        if ($null -ne $vis) {
            [pscustomobject]@{ Section = $sectionNames[$i]; Height = [int]([int](Get-FormValue $Name ".Section($i).Height") / 15)
                               Visible = [bool]$vis; BackColor = (Find-FormValue $Name ".Section($i).BackColor") }
        }
    }
    [pscustomobject]@{ Form = $Name; WidthPx = [int]([int](Get-FormValue $Name '.Width') / 15); Sections = @($sections); Controls = @($list) }
}

# Snapshots live outside the per-session capture folders, because the two sides of a
# comparison come from two sessions: the base build and the PR build, usually the same
# front-end file rebuilt in between.
function Get-ReviewSnapshotRoot { Join-Path (Get-ReviewOutputBase) 'snapshots' }

function Get-ReviewSnapshotDir([string]$Label) {
    if ($Label -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') { throw "Snapshot label '$Label' must be letters, digits, '.', '_' or '-'." }
    Join-Path (Get-ReviewSnapshotRoot) $Label
}

# Names of every form in the open front end.
function Get-ReviewFormNames {
    $acc = $script:Review.Acc
    try {
        $all = $acc.CurrentProject.AllForms
        $names = for ($i = 0; $i -lt $all.Count; $i++) { [string]$all.Item($i).Name }
    } catch {
        # Fallback when the COM adapter can't read AllForms: forms are type -32768.
        $rs = $acc.CurrentDb().OpenRecordset('SELECT Name FROM MSysObjects WHERE Type = -32768')
        $names = while (-not $rs.EOF) { [string]$rs.Fields.Item('Name').Value; $rs.MoveNext() }
        $rs.Close()
    }
    @($names | Sort-Object)
}

# Saves each form's definition (SaveAsText) under snapshots\<Label>\<form>.txt. Without
# -Forms, every form, and the label's folder is cleared first so a form the build no
# longer makes doesn't linger. Reads the saved design only: nothing in the file changes,
# and changes code makes while a form is open (colours set on load) are not included.
function Save-ReviewSnapshot {
    param([Parameter(Mandatory)][string]$Label, [string[]]$Forms)
    Assert-ReviewOpen
    $dir = Get-ReviewSnapshotDir $Label
    $names = if ($Forms) { $Forms } else { Get-ReviewFormNames }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    if (-not $Forms) { Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Extension -in '.txt', '.failed', '.list' } | Remove-Item }
    $dialogs = @(); $saved = @(); $failed = @()
    $badList = Join-Path $dir 'unsaved-names.list'
    foreach ($n in $names) {
        if ($n -match '[\\/:*?"<>|]') {
            # No file can carry this name, so it goes in a list compare_snapshots reads.
            $known = if (Test-Path -LiteralPath $badList) { @(Get-Content -LiteralPath $badList -Encoding UTF8) } else { @() }
            if ($known -notcontains $n) { Add-Content -LiteralPath $badList -Value $n -Encoding UTF8 }
            $failed += [pscustomobject]@{ Form = $n; Error = "The name has a character that can't be used in a file name." }; continue
        }
        $path = Join-Path $dir "$n.txt"; $mark = Join-Path $dir "$n.failed"
        # Clear this form's earlier result first, so a failed save never leaves an old
        # definition behind to be compared as if it were this build's.
        foreach ($f in $path, $mark) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f } }
        try {
            $r = Invoke-ReviewGuarded { $script:Review.Acc.SaveAsText(2, $n, $path) }   # 2 = acForm
            $dialogs += $r.Dialogs
            $saved += [pscustomobject]@{ Form = $n; Objects = (Read-FormSnapshot $path).Count }
        } catch {
            $msg = $_.Exception.Message
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
            # compare_snapshots reads this marker and reports the form as not compared
            # rather than removed.
            [IO.File]::WriteAllText($mark, $msg, (New-Object Text.UTF8Encoding($false)))
            $failed += [pscustomobject]@{ Form = $n; Error = $msg }
        }
    }
    [pscustomobject]@{ Label = $Label; Folder = $dir; Saved = $saved; Failed = $failed; Dialogs = $dialogs }
}

# Compares two snapshots by label (Access need not be open). Saves the full report as
# "snapshots\compare <before> to <after>.txt" (labels have no spaces, so the name is
# unambiguous) and returns it, with the differences per form.
function Compare-ReviewSnapshot {
    param([Parameter(Mandatory)][string]$Before, [Parameter(Mandatory)][string]$After, [string[]]$Forms, [int]$MaxLines = 400)
    $cmp = Compare-FormSnapshotFolder -BeforeDir (Get-ReviewSnapshotDir $Before) -AfterDir (Get-ReviewSnapshotDir $After) -Forms $Forms
    $text = Format-FormSnapshotReport $cmp "Form snapshots: $Before -> $After"
    $path = Join-Path (Get-ReviewSnapshotRoot) "compare $Before to $After.txt"
    [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($false)))
    $lines = $text -split "`r?`n"
    $shown = if ($lines.Count -gt $MaxLines) { (($lines[0..($MaxLines - 1)]) -join "`n") + "`n... $($lines.Count - $MaxLines) more lines in $path" } else { $text }
    [pscustomobject]@{
        Report = $path
        FormsAdded = $cmp.FormsAdded; FormsRemoved = $cmp.FormsRemoved; NotCompared = $cmp.NotCompared
        FormsChanged = @($cmp.Forms | Where-Object { $_.Differences -gt 0 } | ForEach-Object { "$($_.Form) ($($_.Differences))" })
        FormsUnchanged = @($cmp.Forms | Where-Object { $_.Differences -eq 0 }).Count
        Text = $shown
    }
}

# Open windows of the Access instance other than its main window: message boxes and
# popup forms, each described and captured. -Close closes them.
function Get-ReviewDialogs {
    param([switch]$Close)
    Assert-ReviewOpen
    $r = $script:Review
    $found = foreach ($h in [KitReviewWin]::TopWindows([uint32]$r.Pid)) {
        if ($h -eq $r.Hwnd) { continue }
        $shot = Join-Path $r.Out ('window-' + (Get-Date -Format 'HHmmss') + "-$([long]$h).png")
        try { [void][KitReviewWin]::Save($h, $shot, 0, 0, 0, 0, 1) } catch { $shot = "(not captured: $($_.Exception.Message))" }
        $item = [pscustomobject]@{ Class = [KitReviewWin]::Cls($h); Text = [KitReviewWin]::Describe($h); Capture = $shot; Closed = $null }
        if ($Close) { $item.Closed = [KitReviewWin]::Dismiss($h) }
        $item
    }
    [pscustomobject]@{ Open = @($found); SeenDuringCalls = @($r.Dialogs) }
}

# Error log check (modErrors): -Mark records where the log ends now; without -Mark it
# returns the entries added since the mark, from the error_log table and the
# errors.log file. Any new entry is a must-fix finding unless explained.
function Get-ReviewErrorLog {
    param([switch]$Mark)
    Assert-ReviewOpen
    $acc = $script:Review.Acc
    $table = $null; $maxId = 0; $rows = @()
    try {
        $db = $acc.CurrentDb()
        $rs = $db.OpenRecordset('SELECT Max(id) AS m FROM error_log')
        $m = $rs.Fields.Item('m').Value; $rs.Close()
        $maxId = if ($m -is [DBNull] -or $null -eq $m) { 0 } else { [int]$m }
        $table = 'error_log'
    } catch { $table = $null }
    $file = $null; $len = 0
    try { $file = [string]$acc.Run('LogFilePath') } catch {}
    if ($file -and (Test-Path -LiteralPath $file)) { $len = (Get-Item -LiteralPath $file).Length }

    if ($Mark -or -not $script:Review.LogMark) {
        $script:Review.LogMark = [pscustomobject]@{ Id = $maxId; File = $file; Length = $len }
        return [pscustomobject]@{ Marked = $true; Table = $table; LastId = $maxId; File = $file; FileLength = $len
                                  Note = $(if (-not $table -and -not $file) { 'This front end has no error_log table or LogFilePath; nothing to watch.' } else { $null }) }
    }
    $mk = $script:Review.LogMark
    if ($table -and $maxId -gt $mk.Id) {
        $rs = $db.OpenRecordset("SELECT id, logged_at, error_number, description, source, active_form, active_control FROM error_log WHERE id > $($mk.Id) ORDER BY id")
        while (-not $rs.EOF) {
            $rows += [pscustomobject]@{ Id = $rs.Fields.Item('id').Value; At = [string]$rs.Fields.Item('logged_at').Value
                                        Error = "$($rs.Fields.Item('error_number').Value): $($rs.Fields.Item('description').Value)"
                                        Procedure = [string]$rs.Fields.Item('source').Value
                                        Form = [string]$rs.Fields.Item('active_form').Value; Control = [string]$rs.Fields.Item('active_control').Value }
            $rs.MoveNext()
        }
        $rs.Close()
    }
    $fileText = $null
    if ($file -and $len -gt $mk.Length -and $file -eq $mk.File) {
        $fs = [IO.File]::Open($file, 'Open', 'Read', 'ReadWrite')
        try { [void]$fs.Seek($mk.Length, 'Begin'); $fileText = (New-Object IO.StreamReader($fs)).ReadToEnd() } finally { $fs.Close() }
    }
    [pscustomobject]@{ NewEntries = $rows.Count; Entries = @($rows); NewFileText = $fileText; SinceId = $mk.Id }
}
