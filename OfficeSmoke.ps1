# End-to-end check of the Office tools through the real MCP server, for a review
# machine with Office installed (never a PC someone is working on: it starts Excel,
# Word and, when installed, PowerPoint):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\review-mcp\OfficeSmoke.ps1
# It starts AccessReviewMcp.ps1 with a throwaway allowed folder and output folder under
# %TEMP%, talks JSON-RPC to it over stdio, prints each check and exits 1 if any fails.
# Apps that are not installed are reported as skipped.

$ErrorActionPreference = 'Stop'
$root = Join-Path $env:TEMP ('office-smoke-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$files = Join-Path $root 'files'; $out = Join-Path $root 'out'
New-Item -ItemType Directory -Force -Path $files, $out | Out-Null

$failures = 0; $skipped = @()
function Check([string]$what, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what $detail"; $script:failures++ }
}

$psi = New-Object Diagnostics.ProcessStartInfo
$psi.FileName = 'powershell.exe'
$psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + (Join-Path $PSScriptRoot 'AccessReviewMcp.ps1') + '"'
$psi.UseShellExecute = $false; $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
$psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
$psi.EnvironmentVariables['ACCESS_REVIEW_ALLOWED_FOLDERS'] = $files
$psi.EnvironmentVariables['ACCESS_REVIEW_OUTPUT'] = $out
$psi.EnvironmentVariables['ACCESS_REVIEW_OFFICE_APPS'] = 'excel;word;powerpoint;visio;access'
$server = [Diagnostics.Process]::Start($psi)
$stdin = New-Object IO.StreamWriter($server.StandardInput.BaseStream, (New-Object Text.UTF8Encoding($false)))
$stdin.AutoFlush = $true
$errTask = $server.StandardError.ReadToEndAsync()
$script:id = 0

function Rpc([string]$method, $params) {
    $script:id++
    $stdin.WriteLine((@{ jsonrpc = '2.0'; id = $script:id; method = $method; params = $params } | ConvertTo-Json -Depth 20 -Compress))
    $line = $server.StandardOutput.ReadLine()
    if ($null -eq $line) { throw "Server stopped. Its log: $($errTask.Result)" }
    $line | ConvertFrom-Json
}

# Calls a tool; returns @{ Ok; Text; Data (parsed JSON when the text is JSON); Image }.
function Tool([string]$name, [hashtable]$a = @{}) {
    $r = (Rpc 'tools/call' @{ name = $name; arguments = $a }).result
    $text = ($r.content | Where-Object { $_.type -eq 'text' } | Select-Object -First 1).text
    $data = $null; try { $data = $text | ConvertFrom-Json } catch {}
    [pscustomobject]@{ Ok = -not $r.isError; Text = $text; Data = $data; Image = [bool]($r.content | Where-Object { $_.type -eq 'image' }) }
}

try {
    $init = Rpc 'initialize' @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 'office-smoke'; version = '1' } }
    Check 'server starts (0.7.0)' ($init.result.serverInfo.version -eq '0.7.0')
    $names = @((Rpc 'tools/list' @{}).result.tools | ForEach-Object name)
    foreach ($t in 'office_apps', 'office_start', 'office_invoke', 'office_members', 'excel_range', 'word_text', 'powerpoint_slides', 'outlook_items', 'office_run_macro') { Check "tool $t listed" ($names -contains $t) }

    $apps = Tool office_apps
    $has = @{}; foreach ($x in $apps.Data) { $has[$x.App] = [bool]$x.Installed }
    Write-Host ("     installed: " + (@($apps.Data | Where-Object Installed | ForEach-Object App) -join ', '))
    Check 'Outlook off by default' (-not (Tool office_start @{ app = 'outlook' }).Ok)
    Check 'OneNote sent to onenote tools' ((Tool office_start @{ app = 'onenote' }).Text -like '*onenote_*')

    # --- Excel ---
    if ($has.excel) {
        $s = Tool office_start @{ app = 'excel' }
        Check 'Excel starts in its own instance' ($s.Ok -and $s.Data.StartedByTool) $s.Text
        $wb = (Tool office_new @{ app = 'excel' }).Data.Handle
        Check 'new workbook handle' ($wb -match '^h\d+$')
        $w = Tool excel_range @{ document = $wb; range = 'A1:B2'; values = @(@(1, 2), @('x', $true)) }
        Check 'excel_range writes and reads back' ($w.Ok -and $w.Data.Values.Rows[0][1] -eq 2 -and $w.Data.Values.Rows[1][0] -eq 'x') $w.Text
        $f = Tool office_invoke @{ target = $wb; path = "Worksheets(1).Range('C1').Formula"; value = '=SUM(A1:B1)' }
        Check 'office_invoke sets a formula' $f.Ok $f.Text
        $v = Tool office_invoke @{ target = $wb; path = "Worksheets(1).Range('C1').Value2" }
        Check 'formula result read' ($v.Ok -and $v.Data.Value -eq 3) $v.Text
        $blk = Tool office_invoke @{ target = $wb; path = "Worksheets(1).Range('A3:B3').Value2"; value = @(@(7, 8)) }
        Check 'block set through office_invoke' ($blk.Ok -and (Tool excel_range @{ document = $wb; range = 'A3' }).Data.Values.Rows[0][0] -eq 7) $blk.Text
        $ws = Tool office_invoke @{ target = $wb; path = 'Worksheets(1)'; fields = @('Name', 'UsedRange.Address(false, false)') }
        Check 'object comes back as a handle with fields' ($ws.Ok -and $ws.Data.Value.Handle -match '^h\d+$' -and $ws.Data.Value.Type -eq 'Worksheet' -and $ws.Data.Fields.'UsedRange.Address(false, false)' -eq 'A1:C3') $ws.Text
        $sheetH = $ws.Data.Value.Handle
        $cell = Tool office_invoke @{ target = $sheetH; path = 'Cells(1, 2).Value2' }
        Check 'handle used as a target' ($cell.Ok -and $cell.Data.Value -eq 2) $cell.Text
        $en = Tool office_invoke @{ target = $wb; path = 'Worksheets'; enumerate = $true }
        Check 'collection enumerated' ($en.Ok -and $en.Data.Count -ge 1 -and $en.Data.Items[0].Fields.Name) $en.Text
        $m = Tool office_members @{ target = $wb; filter = 'Save*' }
        Check 'office_members reads the type library' ($m.Ok -and $m.Data.Type -eq 'Workbook' -and @($m.Data.Members | ForEach-Object Name) -contains 'SaveAs') $m.Text
        $mx = Tool office_members @{ target = 'excel'; filter = 'ExecuteExcel4Macro' }
        Check 'office_members marks refused members' ($mx.Ok -and $mx.Data.Members[0].Refused) $mx.Text
        $x4 = Tool office_invoke @{ target = 'excel'; path = "ExecuteExcel4Macro('EXEC(""calc"")')" }
        Check 'ExecuteExcel4Macro refused' (-not $x4.Ok -and $x4.Text -like '*not allowed*') $x4.Text
        $dde = Tool excel_range @{ document = $wb; range = 'D1'; values = "=cmd|' /c calc'!A0" }
        Check 'DDE formula refused' (-not $dde.Ok -and $dde.Text -like '*DDE*') $dde.Text
        $outside = Tool office_invoke @{ target = $wb; path = "SaveCopyAs('C:\Windows\Temp\smoke.xlsx')" }
        Check 'save outside the allowed folders refused' (-not $outside.Ok -and $outside.Text -like '*outside*') $outside.Text
        $xlsx = Join-Path $files 'smoke.xlsx'
        $sv = Tool office_save_as @{ document = $wb; path = $xlsx; format = 51 }
        Check 'office_save_as writes xlsx' ($sv.Ok -and (Test-Path $xlsx)) $sv.Text
        Check 'no overwrite without asking' (-not (Tool office_save_as @{ document = $wb; path = $xlsx }).Ok)
        $pdf = Tool office_export_pdf @{ document = $wb }
        Check 'office_export_pdf writes a PDF' ($pdf.Ok -and (Test-Path $pdf.Data.Pdf)) $pdf.Text
        $cap = Tool office_capture @{ target = $wb; label = 'smoke' }
        Check 'office_capture returns an image' ($cap.Ok -and $cap.Image -and -not $cap.Data.Blank) $cap.Text
        Check 'office_close' (Tool office_close @{ document = $wb }).Ok
        $ro = Tool office_open @{ app = 'excel'; path = $xlsx }
        Check 'office_open read-only' ($ro.Ok -and $ro.Data.ReadOnly -and -not $ro.Data.MacrosEnabled) $ro.Text
        Check 'reopened file has the data' ((Tool excel_range @{ document = $ro.Data.Handle; range = 'C1' }).Data.Values.Rows[0][0] -eq 3)
        Check 'macro refused when macros are off' (-not (Tool office_run_macro @{ document = $ro.Data.Handle; macro = 'Foo' }).Ok)
        $q = Tool office_quit @{ app = 'excel' }
        Check 'office_quit' ($q.Ok) $q.Text
        Start-Sleep -Seconds 2
        Check 'Excel process gone' (-not (Get-Process -Id $s.Data.Pid -ErrorAction SilentlyContinue))
    } else { $skipped += 'Excel' }

    # --- Word ---
    if ($has.word) {
        $s = Tool office_start @{ app = 'word' }
        Check 'Word starts' ($s.Ok -and $s.Data.StartedByTool) $s.Text
        $d = (Tool office_new @{ app = 'word' }).Data.Handle
        $t = Tool office_invoke @{ target = $d; path = 'Content.Text'; value = "Review heading`rFirst paragraph." }
        Check 'Word text set' $t.Ok $t.Text
        $st = Tool office_invoke @{ target = $d; path = 'Paragraphs(1).Style'; value = 'Heading 1' }
        Check 'style set by name' $st.Ok $st.Text
        $wt = Tool word_text @{ document = $d }
        Check 'word_text reads paragraphs and styles' ($wt.Ok -and $wt.Data.Paragraphs -ge 2 -and $wt.Data.Shown[0].Text -eq 'Review heading' -and $wt.Data.Shown[0].OutlineLevel -eq 1) $wt.Text
        $docx = Join-Path $files 'smoke.docx'
        $na = Tool office_invoke @{ target = $d; path = "SaveAs2(FileName:='$docx', FileFormat:=16)" }
        Check 'named arguments (SaveAs2 FileName:=, FileFormat:=)' ($na.Ok -and (Test-Path $docx)) $na.Text
        $mix = Tool office_invoke @{ target = $d; path = "Range(0, 6).Text" }
        Check 'positional arguments' ($mix.Ok -and $mix.Data.Value -eq 'Review') $mix.Text
        Check 'Word PDF' (Tool office_export_pdf @{ document = $d }).Ok
        $cap = Tool office_capture @{ target = $d }
        Check 'Word capture' ($cap.Ok -and $cap.Image -and -not $cap.Data.Blank) $cap.Text
        Check 'WordBasic refused' (-not (Tool office_invoke @{ target = 'word'; path = 'WordBasic' }).Ok)
        $q = Tool office_quit @{ app = 'word' }
        Check 'Word quits' $q.Ok $q.Text
    } else { $skipped += 'Word' }

    # --- PowerPoint ---
    if ($has.powerpoint) {
        $s = Tool office_start @{ app = 'powerpoint' }
        Check 'PowerPoint starts' $s.Ok $s.Text
        if ($s.Ok) {
            $p = (Tool office_new @{ app = 'powerpoint' }).Data.Handle
            $sl = Tool office_invoke @{ target = $p; path = 'Slides.Add(1, 12)' }
            $tb = Tool office_invoke @{ target = $sl.Data.Value.Handle; path = 'Shapes.AddTextbox(1, 40, 40, 400, 60)' }
            $null = Tool office_invoke @{ target = $tb.Data.Value.Handle; path = 'TextFrame.TextRange.Text'; value = 'Overdue: 5' }
            $null = Tool office_invoke @{ target = $tb.Data.Value.Handle; path = "Tags.Add('spm.key', 'overdue')" }
            $ps = Tool powerpoint_slides @{ document = $p }
            Check 'powerpoint_slides reads shapes, text and tags' ($ps.Ok -and $ps.Data.Slides[0].Shapes[0].Text -eq 'Overdue: 5' -and $ps.Data.Slides[0].Shapes[0].Tags.'SPM.KEY' -eq 'overdue') $ps.Text
            Check 'PowerPoint PDF' (Tool office_export_pdf @{ document = $p }).Ok
            Check 'PowerPoint quits' (Tool office_quit @{ app = 'powerpoint' }).Ok
        }
    } else { $skipped += 'PowerPoint' }

    $audit = Join-Path $out 'office-audit.log'
    Check 'audit log written' ((Test-Path $audit) -and @(Get-Content $audit).Count -gt 10)
    Check 'refusals are in the audit log' ((Get-Content $audit -Raw) -match '"ok":false')
} catch {
    Write-Host "FAIL smoke run stopped: $($_.Exception.Message)"; $failures++
} finally {
    try { $stdin.Close() } catch {}
    if (-not $server.WaitForExit(20000)) { $server.Kill() }
}

Write-Host ''
if ($skipped.Count) { Write-Host "skipped (not installed): $($skipped -join ', ')" }
Write-Host "output: $root"
if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'all checks passed'
