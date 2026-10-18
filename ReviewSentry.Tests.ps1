# Checks ReviewSentry.ps1 without Access or a network:
#   powershell -NoProfile -File ReviewSentry.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReviewSentry.ps1')

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}

$savedDsn = $env:ACCESS_REVIEW_SENTRY_DSN; $savedEnv = $env:ACCESS_REVIEW_SENTRY_ENVIRONMENT
try {
    # Settings
    $env:ACCESS_REVIEW_SENTRY_DSN = $null
    Check 'off when no DSN is set' ($null -eq (Get-ReviewSentryConfig))
    $env:ACCESS_REVIEW_SENTRY_DSN = '${user_config.sentry_dsn}'
    Check 'off when the extension leaves the placeholder' ($null -eq (Get-ReviewSentryConfig))
    $env:ACCESS_REVIEW_SENTRY_DSN = 'http://abc@example.com/1'
    Check 'off for a DSN that is not https' ($null -eq (Get-ReviewSentryConfig))
    $env:ACCESS_REVIEW_SENTRY_DSN = ' https://0123abcd@o1.ingest.us.sentry.io/42 '
    $cfg = Get-ReviewSentryConfig
    Check 'DSN key read' ($cfg.Key -eq '0123abcd')
    Check 'envelope URL built' ($cfg.Url -eq 'https://o1.ingest.us.sentry.io/api/42/envelope/')
    Check 'environment defaults to review' ($cfg.Environment -eq 'review')
    $env:ACCESS_REVIEW_SENTRY_DSN = 'https://abc123:def456@sentry.example.com:9443/sentry/7'
    $self = Get-ReviewSentryConfig
    Check 'self-hosted DSN with secret, port and path' ($self.Key -eq 'abc123' -and $self.Url -eq 'https://sentry.example.com:9443/sentry/api/7/envelope/')
    $env:ACCESS_REVIEW_SENTRY_DSN = 'https://not-hex@example.com/1'
    Check 'off for a DSN without a key' ($null -eq (Get-ReviewSentryConfig))
    $env:ACCESS_REVIEW_SENTRY_DSN = ' https://0123abcd@o1.ingest.us.sentry.io/42 '
    $env:ACCESS_REVIEW_SENTRY_ENVIRONMENT = 'REVIEW-02'
    Check 'environment from the setting' ((Get-ReviewSentryConfig).Environment -eq 'REVIEW-02')

    # Events carry names and numbers only
    $e = New-ReviewSentryEvent -Kind error_log -App 'KitDemo.accdb' -ErrorNumber 2196 -Procedure 'modSecret.ToggleSecret line 42' `
        -Form 'frmMain' -Control 'txtSecret' -AccessVersion '16.0'
    Check 'title names the error and procedure' ($e.message.formatted -eq 'Logged error 2196 in modSecret.ToggleSecret line 42')
    Check 'procedure tag drops the line' ($e.tags.procedure -eq 'modSecret.ToggleSecret')
    Check 'fingerprint ignores the line' (($e.fingerprint -join '|') -eq 'access-review|error_log|KitDemo.accdb|2196|modSecret.ToggleSecret')
    Check 'event id is 32 hex characters' ($e.event_id -match '^[0-9a-f]{32}$')
    $json = $e | ConvertTo-Json -Depth 5 -Compress
    Check 'no user or computer fields' ($json -notmatch 'user|computer|server_name|description')
    $odd = New-ReviewSentryEvent -Kind error_log -Procedure "pgX.Save (record 'SMITH, JOHN')"
    Check 'context after a procedure dropped' ($odd.tags.procedure -eq 'pgX.Save' -and ($odd | ConvertTo-Json -Depth 5) -notmatch 'SMITH')

    # Guard notes: message boxes, hangs and exits only, and never their text
    $notes = @(
        "[message box] [Microsoft Visual Basic] Run-time error '2196': Field SSN 123-45-6789 | capture C:\x.png | pressed OK",
        '[window OForm] [dlgMessage] Saved | capture C:\y.png',
        '[timeout] Access still busy after 120s; ended it',
        '[exited] Access closed during the call')
    $d = @(ConvertFrom-ReviewDialogNotes $notes -App 'KitDemo.accdb')
    Check 'kit dialogs skipped, three findings kept' ($d.Count -eq 3)
    Check 'message box error number read' ($d[0].tags.error_number -eq '2196')
    Check 'dialog text never sent' ((($d | ConvertTo-Json -Depth 5 -Compress)) -notmatch 'SSN|123-45|Field|capture')
    Check 'hang reported as timeout' ($d[1].tags.kind -eq 'timeout' -and $d[2].tags.kind -eq 'exited')

    # errors.log entries
    $log = @"
------------------------------------------------------------------------
2026-10-18 01:02:03  error 3022: The changes you requested were not successful because they would create duplicate values
Procedure:   pgPeople.btnSave_Click line 88  (person JONES)
Page:        people  form frmMain  control btnSave
User:        jdoe  Windows jdoe on PC123
Build:       1.2.3  Access 16.0.1  source DAO
------------------------------------------------------------------------
2026-10-18 01:05:00  error 91: Object variable not set
Procedure:   modKitUI.Paint
"@
    $l = @(ConvertFrom-ReviewLogText $log -App 'SPM.accdb')
    Check 'two log entries read' ($l.Count -eq 2)
    Check 'log entry number and procedure' ($l[0].tags.error_number -eq '3022' -and $l[0].message.formatted -eq 'Logged error 3022 in pgPeople.btnSave_Click line 88')
    Check 'log context, user and computer never sent' ((($l | ConvertTo-Json -Depth 5 -Compress)) -notmatch 'JONES|jdoe|PC123|duplicate')
    Check 'last entry without a line' ($l[1].tags.procedure -eq 'modKitUI.Paint')

    # Grouping: identical findings collapse with a count; distinct ones are all kept
    $many = @(1..25 | ForEach-Object { New-ReviewSentryEvent -Kind error_log -App 'A.accdb' -ErrorNumber $_ -Procedure 'm.p' })
    $many += New-ReviewSentryEvent -Kind error_log -App 'A.accdb' -ErrorNumber 1 -Procedure 'm.p line 9'
    $g = @(Group-ReviewSentryEvents $many)
    Check 'all 25 distinct findings kept' ($g.Count -eq 25)
    Check 'repeat counted on its group' ($g[0].tags.occurrences -eq '2' -and $g[1].tags.occurrences -eq '1')

    # Envelope
    $envelope = ConvertTo-ReviewSentryEnvelope $cfg $e
    $parts = $envelope -split "`n"
    Check 'envelope has header, item header and event' ($parts.Count -ge 3 -and ($parts[1] | ConvertFrom-Json).type -eq 'event')
    Check 'item length matches the event bytes' (($parts[1] | ConvertFrom-Json).length -eq [Text.Encoding]::UTF8.GetByteCount($parts[2]))

    # Sending: nothing without a DSN; an unreachable host stops the session after one try
    $env:ACCESS_REVIEW_SENTRY_DSN = $null
    Check 'nothing sent when off' ((Send-ReviewSentry @($e)) -eq 0)
    $env:ACCESS_REVIEW_SENTRY_DSN = 'https://0123abcd@sentry.invalid/42'
    $script:SentryDown = $false
    Check 'unreachable host sends nothing and does not throw' ((Send-ReviewSentry @($e, $e)) -eq 0)
    Check 'session stops trying after a failure' ($script:SentryDown)
} finally {
    $env:ACCESS_REVIEW_SENTRY_DSN = $savedDsn; $env:ACCESS_REVIEW_SENTRY_ENVIRONMENT = $savedEnv
}

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'all checks passed'
