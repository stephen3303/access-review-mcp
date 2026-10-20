# MCP server (stdio) exposing AccessReview.ps1 to a Claude Code session, so visual
# review drives Access through COM and PrintWindow instead of the desktop's mouse,
# keyboard and screen. One JSON-RPC message per line on stdin and stdout; nothing else
# may be written to stdout. See README.md for setup.

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$utf8 = New-Object Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8

. (Join-Path $PSScriptRoot 'AccessReview.ps1')

$ServerVersion = '0.5.0'

function Send-Message($obj) {
    [Console]::Out.WriteLine(($obj | ConvertTo-Json -Depth 20 -Compress))
    [Console]::Out.Flush()
}

function Log([string]$text) { [Console]::Error.WriteLine("access-review: $text") }

function Arg($a, [string]$name, $default = $null) {
    if ($a -and $a.PSObject.Properties[$name] -and $null -ne $a.$name) { $a.$name } else { $default }
}

function Schema([hashtable]$props, [string[]]$required = @()) {
    $s = @{ type = 'object'; properties = $props }
    if ($required.Count) { $s.required = $required }
    $s
}

$str = @{ type = 'string' }
$int = @{ type = 'integer' }
$formProp = @{ type = 'string'; description = 'Form name, e.g. frmMain.' }

# name, description, input schema, handler (receives the arguments object).
$Tools = @(
    @{ name = 'review_lock'; description = 'Take the shared reviewer lock files (ACCESS_REVIEW_LOCKS) before using Access on a machine other reviewers share. Fails if another reviewer holds them. Does nothing when no lock files are configured.'
       schema = (Schema @{ owner = @{ type = 'string'; description = 'Your thread name.' } } @('owner'))
       run = { param($a) Enter-ReviewLock -Owner (Arg $a 'owner') } }
    @{ name = 'review_unlock'; description = 'Release the lock files this owner took.'
       schema = (Schema @{ owner = $str } @('owner'))
       run = { param($a) Exit-ReviewLock -Owner (Arg $a 'owner') } }
    @{ name = 'open_app'; description = 'Start a separate Access instance and open an existing, already trusted front end (never creates a file). mode offscreen (default) parks every Access window past the desktop edge so it never covers anyone''s work; onscreen is for a dedicated review machine. Startup forms and AutoExec run as usual, so it refuses a file downloaded from the internet (Zone.Identifier) and, when ACCESS_REVIEW_ALLOWED_FOLDERS is set, a file outside those folders.'
       schema = (Schema @{ path = @{ type = 'string'; description = 'Full path of the .accdb/.accde.' }; mode = @{ type = 'string'; enum = @('offscreen', 'onscreen') }
                          output_dir = @{ type = 'string'; description = 'Where PNGs go. Default ACCESS_REVIEW_OUTPUT or %LOCALAPPDATA%\AccessReview, in a new <time> folder.' } } @('path'))
       run = { param($a) $p = @{ Path = (Arg $a 'path'); Mode = (Arg $a 'mode' 'offscreen') }; if (Arg $a 'output_dir') { $p.OutputDirectory = $a.output_dir }; Open-ReviewApp @p } }
    @{ name = 'close_app'; description = "Reset the palette to $DefaultPalette, close the database without saving design changes, quit Access and confirm the process is gone."
       schema = (Schema @{}); run = { param($a) Close-ReviewApp } }
    @{ name = 'compile'; description = 'Compile and save all modules; reports any VBA dialog text.'
       schema = (Schema @{}); run = { param($a) Test-ReviewCompile } }
    @{ name = 'error_log'; description = 'mark=true records where the app error log ends now (an error_log table, and the file a public LogFilePath function returns, as access-modern-ui apps have); call it before testing. Then call with mark=false to get entries added since. Treat any new entry as a finding unless explained. Says so when the app has neither.'
       schema = (Schema @{ mark = @{ type = 'boolean' } }); run = { param($a) Get-ReviewErrorLog -Mark:([bool](Arg $a 'mark' $false)) } }
    @{ name = 'run'; description = 'Run a public VBA procedure or function with Application.Run (e.g. a sign-in procedure with username, password and an output error argument). Returns the result and the arguments after the call.'
       schema = (Schema @{ procedure = $str; args = @{ type = 'array'; items = @{} } } @('procedure'))
       run = { param($a) Invoke-ReviewRun -Name (Arg $a 'procedure') -Arguments @(Arg $a 'args' @()) } }
    @{ name = 'set_palette'; description = "Switch the app's colour palette by running its palette procedure ($PaletteProcedure, as in access-modern-ui), then reopen every open form at its last requested size (a form left open keeps part of the old drawing). A form with a public Reopen method (an app shell that signs the user out on close) reopens itself on its current page (its PageKey) so the user stays signed in, and the call fails rather than closing it if that cannot be done; other forms are closed and reopened with their OpenArgs. close_app resets to $DefaultPalette. Fails with a clear message when the app has no such procedure."
       schema = (Schema @{ palette = @{ type = 'string'; enum = @($Palettes.Keys) } } @('palette'))
       run = { param($a) Set-ReviewPalette -Name (Arg $a 'palette') } }
    @{ name = 'open_form'; description = 'Open a form and size it to its designed dimensions (form width and visible section heights), or to width_px x height_px (96 dpi pixels) for forms sized in code. Reports design, inside and window sizes and whether the Access window clips it.'
       schema = (Schema @{ form = $formProp; open_args = $str; width_px = $int; height_px = $int } @('form'))
       run = { param($a) Open-ReviewForm -Name (Arg $a 'form') -OpenArgs (Arg $a 'open_args') -WidthPx (Arg $a 'width_px' 0) -HeightPx (Arg $a 'height_px' 0) } }
    @{ name = 'resize_form'; description = 'Resize an open form: to its design size (no size given), to width_px x height_px (the small check is about 1100x700), or maximized=true (the Access window at the main display''s work area, as a maximized window gets). A shell form that fills the Access window (kit frmMain) is sized by sizing the Access window; SizeMatches=false and a warning mean the size could not be reached.'
       schema = (Schema @{ form = $formProp; width_px = $int; height_px = $int; maximized = @{ type = 'boolean' } } @('form'))
       run = { param($a) Set-ReviewFormSize -Name (Arg $a 'form') -WidthPx (Arg $a 'width_px' 0) -HeightPx (Arg $a 'height_px' 0) -Maximize:([bool](Arg $a 'maximized' $false)) } }
    @{ name = 'capture'; description = 'PNG of an open form drawn with PrintWindow (no screen needed, works behind other windows). Optional crop [x, y, width, height] in physical pixels relative to the form window, enlarged by zoom. Returns the image and its saved path.'
       schema = (Schema @{ form = $formProp; label = @{ type = 'string'; pattern = '^[A-Za-z0-9][A-Za-z0-9_.-]*$'; description = 'Added to the file name: letters, digits, ''.'', ''_'' or ''-'', e.g. sidebar-collapsed.' }
                          crop = @{ type = 'array'; items = $int; minItems = 4; maxItems = 4 }; zoom = @{ type = 'integer'; minimum = 1; maximum = 8 } } @('form'))
       run = { param($a) $c = Arg $a 'crop'; Save-ReviewCapture -Name (Arg $a 'form') -Label (Arg $a 'label') -Crop $(if ($c) { [int[]]@($c) } else { $null }) -Zoom (Arg $a 'zoom' 1) }
       image = $true }
    @{ name = 'controls'; description = 'Every control on an open form with Left/Top/Width/Height in 96 dpi pixels, section, visibility, caption and colours, plus section heights, for measuring gaps, the 16 px card inset and overlaps.'
       schema = (Schema @{ form = $formProp } @('form')); run = { param($a) Get-ReviewControls -Name (Arg $a 'form') } }
    @{ name = 'snapshot'; description = 'Save the definition of forms in the open front end (Application.SaveAsText) as a snapshot under a label, e.g. main for a build of the base branch and pr for the PR build. Without forms, every form, replacing that label''s earlier snapshot. Snapshots outlive close_app, so build and open one side, snapshot, then rebuild and snapshot the other. Reads the saved design only; changes nothing.'
       schema = (Schema @{ label = @{ type = 'string'; description = 'Letters, digits, ., _ or -, e.g. main or pr-45.' }; forms = @{ type = 'array'; items = $str; description = 'Form names; default all.' } } @('label'))
       run = { param($a) $f = Arg $a 'forms'; Save-ReviewSnapshot -Label (Arg $a 'label') -Forms $(if ($f) { [string[]]@($f) } else { $null }) } }
    @{ name = 'compare_snapshots'; description = 'Compare two snapshots property by property (Access need not be open): forms and controls added or removed, and every changed property with before and after values, positions in 96 dpi px and colours as #RRGGBB. Properties Access rewrites on every save (checksums, GUIDs, printer data, cached layout) are ignored; pictures compare by hash; code behind the form is left to the PR diff. Saves the full report as a text file.'
       schema = (Schema @{ before = $str; after = $str; forms = @{ type = 'array'; items = $str; description = 'Only these forms; default all.' } } @('before', 'after'))
       run = { param($a) $f = Arg $a 'forms'; Compare-ReviewSnapshot -Before (Arg $a 'before') -After (Arg $a 'after') -Forms $(if ($f) { [string[]]@($f) } else { $null }) } }
    @{ name = 'export_design'; description = 'Write a snapshot''s forms as JSON designs (Access need not be open), one <form>.json per form plus index.json, under designs\<label> in the output folder, for rebuilding the forms in another app or as HTML pages. Each file has the form''s properties and its sections in order, each with its controls: position and size in 96 dpi px, colours as #RRGGBB, text unquoted, event handlers listed apart, controls on tab pages and in option groups nested under them. With theme_file (the app''s build palette, modTheme.bas, as a full path) each colour also names its kit token and tokens.json lists them. Holds the form as built: changes code makes when the form opens are not in it, and pictures are a hash.'
       schema = (Schema @{ label = $str; theme_file = @{ type = 'string'; description = 'Full path of the app''s modTheme.bas; optional.' }; forms = @{ type = 'array'; items = $str; description = 'Only these forms; default all.' } } @('label'))
       run = { param($a) $f = Arg $a 'forms'; Export-ReviewDesign -Label (Arg $a 'label') -ThemeFile (Arg $a 'theme_file') -Forms $(if ($f) { [string[]]@($f) } else { $null }) } }
    @{ name = 'dialogs'; description = 'List (and capture) open message boxes and popup windows of this Access instance, plus dialogs seen during earlier calls. close=true closes the open ones.'
       schema = (Schema @{ close = @{ type = 'boolean' } }); run = { param($a) Get-ReviewDialogs -Close:([bool](Arg $a 'close' $false)) } }
)

function Invoke-Tool($name, $arguments) {
    $tool = $Tools | Where-Object { $_.name -eq $name } | Select-Object -First 1
    if (-not $tool) { return @{ content = @(@{ type = 'text'; text = "Unknown tool: $name" }); isError = $true } }
    try {
        $result = & $tool.run $arguments
        $content = @(@{ type = 'text'; text = ($result | ConvertTo-Json -Depth 8) })
        if ($tool.image -and $result.Path -and (Test-Path -LiteralPath $result.Path)) {
            $content += @{ type = 'image'; mimeType = 'image/png'; data = [Convert]::ToBase64String([IO.File]::ReadAllBytes($result.Path)) }
        }
        @{ content = $content; isError = $false }
    } catch {
        @{ content = @(@{ type = 'text'; text = $_.Exception.Message }); isError = $true }
    }
}

function Get-ToolList {
    @(foreach ($t in $Tools) { @{ name = $t.name; description = $t.description; inputSchema = $t.schema } })
}

Log "started (version $ServerVersion)"
if ($PaletteWarning) { Log $PaletteWarning }
try {
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if (-not $line.Trim()) { continue }
        try { $msg = $line | ConvertFrom-Json } catch { Send-Message @{ jsonrpc = '2.0'; id = $null; error = @{ code = -32700; message = 'Parse error' } }; continue }
        $hasId = $null -ne $msg.PSObject.Properties['id']
        $reply = $null
        switch ($msg.method) {
            'initialize' {
                $pv = Arg $msg.params 'protocolVersion' '2025-06-18'
                $reply = @{ protocolVersion = $pv; capabilities = @{ tools = @{} }; serverInfo = @{ name = 'access-review'; version = $ServerVersion }
                            instructions = 'Drives Microsoft Access over COM for visual review without using the desktop. Typical run: review_lock, open_app, error_log mark=true, open_form, capture, controls, dialogs, error_log, close_app, review_unlock. To see exactly what a PR changed on its forms: snapshot label=main on a base build and label=pr on the PR build, then compare_snapshots. To hand form designs to another app or an HTML rebuild: snapshot, then export_design.' + $(if ($PaletteWarning) { " Settings problem: $PaletteWarning" }) }
            }
            'ping' { $reply = @{} }
            'tools/list' { $reply = @{ tools = (Get-ToolList) } }
            'tools/call' { $reply = Invoke-Tool $msg.params.name (Arg $msg.params 'arguments') }
            default { $reply = 'unknown' }
        }
        if (-not $hasId) { continue }   # a notification: no response
        if ($reply -eq 'unknown') {
            Send-Message @{ jsonrpc = '2.0'; id = $msg.id; error = @{ code = -32601; message = "Method not found: $($msg.method)" } }
        } else {
            Send-Message @{ jsonrpc = '2.0'; id = $msg.id; result = $reply }
        }
    }
} finally {
    # The client went away: never leave a hidden Access instance holding the file open.
    try { $null = Close-ReviewApp } catch {}
    Log 'stopped'
}
