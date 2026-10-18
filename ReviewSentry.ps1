# Optional reporting of review findings to Sentry, so new error-log entries and VBA
# message boxes from every review run collect in one place with a history across runs.
#
# Off unless ACCESS_REVIEW_SENTRY_DSN is set. Only these fields are sent: the front
# end's file name, the error number, the procedure (with its line), the form and
# control, the kind of finding, the Access version and the tool version. Never the
# error description, the call trail or context, the dialog's text, user names or the
# computer name: an Access message can quote field values, and those may be real data.
#
# Every send has a short timeout and never throws. After the first failure the rest of
# the session sends nothing, so an unreachable Sentry costs one timeout per session.
# Identical findings in one call go as one event with an occurrences count; distinct
# ones are all sent (the project's own rate limit caps a flood).
#
# Settings (environment variables, optional):
#   ACCESS_REVIEW_SENTRY_DSN          the project's DSN; unset means reporting is off
#   ACCESS_REVIEW_SENTRY_ENVIRONMENT  environment name on each event (default review)

$script:SentryVersion = '0.4.2'   # matches the server version
$script:SentryDown = $false

# The DSN split into what a send needs, or $null when reporting is off or the DSN is
# malformed. Takes Sentry's full form, https://KEY[:SECRET]@HOST[:PORT][/PATH]/PROJECT,
# so a self-hosted server on its own port or path works too.
function Get-ReviewSentryConfig {
    $dsn = [Environment]::GetEnvironmentVariable('ACCESS_REVIEW_SENTRY_DSN')
    if ([string]::IsNullOrWhiteSpace($dsn) -or $dsn -match '^\$\{.*\}$') { return $null }
    $dsn = $dsn.Trim()
    $uri = $null
    if (-not [Uri]::TryCreate($dsn, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { return $null }
    if ($uri.UserInfo -notmatch '^(?<key>[0-9a-fA-F]+)(:[0-9a-fA-F]+)?$') { return $null }
    $key = $Matches.key
    if ($uri.AbsolutePath -notmatch '^(?<prefix>(/[^/]+)*)/(?<project>\d+)/?$') { return $null }
    $url = "https://$($uri.Authority)$($Matches.prefix)/api/$($Matches.project)/envelope/"
    $envName = [Environment]::GetEnvironmentVariable('ACCESS_REVIEW_SENTRY_ENVIRONMENT')
    if ([string]::IsNullOrWhiteSpace($envName) -or $envName -match '^\$\{.*\}$') { $envName = 'review' }
    [pscustomobject]@{
        Dsn = $dsn; Key = $key; Url = $url
        Environment = $envName.Trim()
    }
}

# Keeps a value to the name at its start, so free text can't ride along in a field
# that should only hold one. A procedure keeps "module.procedure" and an optional
# " line N"; anything after (a context in brackets, a record) is dropped.
function Limit-ReviewSentryName([string]$Value, [switch]$Procedure) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $pattern = if ($Procedure) { '^[A-Za-z_]\w*(\.\w+)*( line \d+)?' } else { '^[\w .()\-]+' }
    if ($Value.Trim() -notmatch $pattern) { return $null }
    $v = $Matches[0].Trim()
    if ($v.Length -gt 100) { $v = $v.Substring(0, 100) }
    $v
}

# A finding as a Sentry event. Grouping is by app, kind, error number and procedure
# without its line, so the same error after an edit that moved it stays one issue.
#   Kind:      error_log, log_file, message_box, timeout or exited
#   App:       front end file name, e.g. KitDemo.accdb
function New-ReviewSentryEvent {
    param([Parameter(Mandatory)][string]$Kind, [string]$App, [long]$ErrorNumber = 0,
          [string]$Procedure, [string]$Form, [string]$Control, [string]$AccessVersion)
    $proc = Limit-ReviewSentryName $Procedure -Procedure
    $procNoLine = if ($proc) { $proc -replace ' line \d+$', '' } else { $null }
    $app = Limit-ReviewSentryName $App
    $title = switch ($Kind) {
        'message_box' { 'Message box during review' }
        'timeout'     { 'Access hung during review' }
        'exited'      { 'Access closed during review' }
        default       { 'Logged error' }
    }
    if ($ErrorNumber) { $title += " $ErrorNumber" }
    if ($proc) { $title += " in $proc" }
    $tags = [ordered]@{ kind = $Kind }
    if ($app) { $tags.app = $app }
    if ($ErrorNumber) { $tags.error_number = [string]$ErrorNumber }
    if ($procNoLine) { $tags.procedure = $procNoLine }
    $f = Limit-ReviewSentryName $Form; if ($f) { $tags.form = $f }
    $c = Limit-ReviewSentryName $Control; if ($c) { $tags.control = $c }
    $av = Limit-ReviewSentryName $AccessVersion; if ($av) { $tags.access_version = $av }
    $fp = @('access-review', $Kind, [string]$app, [string]$ErrorNumber, [string]$procNoLine)
    [ordered]@{
        event_id = [guid]::NewGuid().ToString('N')
        timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        platform = 'other'; level = 'error'; logger = 'access-review'
        release = "access-review@$script:SentryVersion"
        message = @{ formatted = $title }
        tags = $tags
        fingerprint = $fp
    }
}

# The error number in a guard note for a VBA or Office message box
# ("[message box] [Microsoft Visual Basic] Run-time error '2196': ..."), or 0.
# Only the number leaves this function; the note's text is never sent.
function Get-ReviewDialogErrorNumber([string]$Note) {
    if ($Note -match "error '(\d+)'") { return [long]$Matches[1] }
    if ($Note -match '\berror (\d{2,5})\b') { return [long]$Matches[1] }
    0
}

# Guard notes worth reporting, as events: message boxes, hangs and Access closing.
# Kit dialogs and popup forms ("[window ...]") are normal and are skipped.
function ConvertFrom-ReviewDialogNotes {
    param([string[]]$Notes, [string]$App, [string]$AccessVersion)
    foreach ($n in @($Notes)) {
        if (-not $n) { continue }
        $kind = $null
        if ($n.StartsWith('[message box]')) { $kind = 'message_box' }
        elseif ($n.StartsWith('[timeout]')) { $kind = 'timeout' }
        elseif ($n.StartsWith('[exited]')) { $kind = 'exited' }
        if (-not $kind) { continue }
        New-ReviewSentryEvent -Kind $kind -App $App -ErrorNumber (Get-ReviewDialogErrorNumber $n) -AccessVersion $AccessVersion
    }
}

# Entries in new errors.log text (modErrors.WriteLogFile), as events. Reads only the
# error number and the procedure name; the description, context, user and computer
# lines are skipped.
function ConvertFrom-ReviewLogText {
    param([string]$Text, [string]$App, [string]$AccessVersion)
    if (-not $Text) { return }
    $num = 0; $proc = $null; $open = $false
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d  error (-?\d+):') {
            if ($open) { New-ReviewSentryEvent -Kind log_file -App $App -ErrorNumber $num -Procedure $proc -AccessVersion $AccessVersion }
            $num = [long]$Matches[1]; $proc = $null; $open = $true
        } elseif ($open -and $line -match '^Procedure:\s+(\S+(?: line \d+)?)') {
            $proc = $Matches[1]
        }
    }
    if ($open) { New-ReviewSentryEvent -Kind log_file -App $App -ErrorNumber $num -Procedure $proc -AccessVersion $AccessVersion }
}

# The envelope body for one event.
function ConvertTo-ReviewSentryEnvelope($Config, $Item) {
    $head = [ordered]@{ event_id = $Item.event_id; dsn = $Config.Dsn; sent_at = $Item.timestamp } | ConvertTo-Json -Compress
    $body = $Item | ConvertTo-Json -Depth 5 -Compress
    $bytes = [Text.Encoding]::UTF8.GetByteCount($body)
    "$head`n" + (@{ type = 'event'; length = $bytes } | ConvertTo-Json -Compress) + "`n$body`n"
}

# One event per fingerprint, with an occurrences tag counting its repeats.
function Group-ReviewSentryEvents([object[]]$Items) {
    $groups = [ordered]@{}
    foreach ($i in @($Items | Where-Object { $_ })) {
        $k = $i.fingerprint -join '|'
        if ($groups.Contains($k)) { $groups[$k].tags.occurrences = [string]([int]$groups[$k].tags.occurrences + 1) }
        else { $i.tags.occurrences = '1'; $groups[$k] = $i }
    }
    @($groups.Values)
}

# Sends events; returns how many findings the accepted events covered. Never throws.
function Send-ReviewSentry {
    param([object[]]$Items)
    if ($script:SentryDown) { return 0 }
    $list = @(Group-ReviewSentryEvents $Items)
    if (-not $list.Count) { return 0 }
    $cfg = Get-ReviewSentryConfig
    if (-not $cfg) { return 0 }
    $sent = 0
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
    foreach ($e in $list) {
        $e.environment = $cfg.Environment
        try {
            $req = [Net.HttpWebRequest]::Create($cfg.Url)
            $req.Method = 'POST'; $req.Timeout = 5000; $req.ReadWriteTimeout = 5000
            $req.ContentType = 'application/x-sentry-envelope'
            $req.Headers.Add('X-Sentry-Auth', "Sentry sentry_version=7, sentry_key=$($cfg.Key), sentry_client=access-review/$script:SentryVersion")
            $body = [Text.Encoding]::UTF8.GetBytes((ConvertTo-ReviewSentryEnvelope $cfg $e))
            $s = $req.GetRequestStream(); try { $s.Write($body, 0, $body.Length) } finally { $s.Close() }
            $resp = $req.GetResponse(); $resp.Close()
            $sent += [int]$e.tags.occurrences
        } catch {
            $script:SentryDown = $true
            break
        }
    }
    $sent
}
