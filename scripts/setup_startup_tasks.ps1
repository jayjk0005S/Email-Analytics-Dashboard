param(
    [switch]$SkipBrowser
)

$ErrorActionPreference = 'Stop'
$projectPath = Split-Path -Parent $PSScriptRoot
$pythonPath = Join-Path $projectPath '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Python environment not found: $pythonPath"
}

$startupArguments = "`"$pythonPath`" -m email_analytics.startup"
if ($SkipBrowser) {
    $startupArguments += ' --no-browser'
}

# Run only in the signed-in user's session so the browser tab can open normally.
schtasks.exe /Create /TN 'Email Analytics - Start Dashboard' /SC ONLOGON /TR $startupArguments /RL LIMITED /F | Out-Null

$mailSource = 'outlook_desktop'
$envFile = Join-Path $projectPath '.env'
if (Test-Path -LiteralPath $envFile) {
    $sourceLine = Select-String -LiteralPath $envFile -Pattern '^\s*MAIL_SOURCE\s*=\s*(.+?)\s*$' | Select-Object -First 1
    if ($sourceLine) {
        $mailSource = $sourceLine.Matches[0].Groups[1].Value.Trim().ToLowerInvariant()
    }
}

Write-Output 'Created: Email Analytics - Start Dashboard'
if ($mailSource -eq 'graph') {
    # Graph benefits from a periodic catch-up task when a webhook notification
    # is missed while this PC is asleep or offline.
    $syncArguments = "`"$pythonPath`" -m email_analytics.mail_sync"
    schtasks.exe /Create /TN 'Email Analytics - Sync Inbox' /SC MINUTE /MO 10 /TR $syncArguments /RL LIMITED /F | Out-Null
    Write-Output 'Created: Email Analytics - Sync Inbox'
} else {
    Write-Output 'Classic Outlook mode: immediate listener starts at sign-in; no separate sync task needed.'
}
