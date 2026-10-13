# Dependency-free tests for the pure functions and the demo report.
# Runs on Windows PowerShell 5.1 and PowerShell 7 (Windows / Linux / macOS):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\Run-Tests.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("oma_tests_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null

# Load the functions without running the menu
. (Join-Path $root 'OutlookMailAssistant.ps1') -Command demo -OutputFolder $tmp

$script:failed = 0
$script:passed = 0
function Assert {
    param([string]$Name, [bool]$Condition)
    if ($Condition) { $script:passed++; Write-Host "  [ok]   $Name" -ForegroundColor Green }
    else { $script:failed++; Write-Host "  [FAIL] $Name" -ForegroundColor Red }
}

Write-Host 'ConvertTo-Day'
$today = [datetime]'2026-10-13'
Assert 'empty = today'          ((ConvertTo-Day '' $today) -eq $today)
Assert 'y = yesterday'          ((ConvertTo-Day 'Y' $today) -eq $today.AddDays(-1))
Assert 'dd.MM.yyyy parsed'      ((ConvertTo-Day '05.10.2026' $today) -eq [datetime]'2026-10-05')
Assert 'invalid -> null'        ($null -eq (ConvertTo-Day '2026-10-05' $today))

Write-Host 'Test-IgnoredSubject'
Assert 'delivery receipt ignored'  (Test-IgnoredSubject 'Delivery Status Notification (Relay)')
Assert 'Delivered: ignored'        (Test-IgnoredSubject 'Delivered: your message')
Assert 'normal subject kept'       (-not (Test-IgnoredSubject 'Question about my file'))

Write-Host 'Helpers'
Assert 'unsafe file name chars'    ((Get-SafeFileName 'a:b/c*?.pdf') -eq 'a_b_c__.pdf')
Assert 'formula injection blocked' ((Protect-SpreadsheetValue '=SUM(A1)') -eq "'=SUM(A1)")
Assert 'plain text untouched'      ((Protect-SpreadsheetValue 'hello') -eq 'hello')
Assert 'text shortened'            ((Limit-Text 'abcdef' 3) -eq 'abc [...]')
Assert 'short text untouched'      ((Limit-Text 'abc' 3) -eq 'abc')
$reply = New-ReplyText 'Body text'
Assert 'reply has greeting, body, signature' ($reply -match '^Hello,' -and $reply -match 'Body text' -and $reply -match 'Your Name$')

Write-Host 'Config'
$cfgPath = Join-Path $tmp 'config.json'
'{"Signature":"Test Dept","Templates":{"9":"Custom text"},"_help":"x"}' | Set-Content -Path $cfgPath -Encoding UTF8
Import-AssistantConfig $cfgPath
Assert 'signature overridden'      ($Config.Signature -eq 'Test Dept')
Assert 'templates replaced'        ($Config.Templates.Contains('9') -and -not $Config.Templates.Contains('1'))
Assert '_help keys skipped'        (-not $Config.Contains('_help'))

Write-Host 'Demo report'
Initialize-Folders
$emails = New-DemoEmails $today
Assert 'four demo e-mails'         ($emails.Count -eq 4)
Assert 'numbered in time order'    (($emails | ForEach-Object Nr) -join ',' -eq '1,2,3,4')
$file = Save-Report $emails $today
$html = Get-Content $file -Raw -Encoding UTF8
Assert 'report file created'       (Test-Path $file)
Assert 'summary counts'            ($html -match '4 e-mails, 2 unread, 3 with attachments')
Assert 'HTML is escaped'           ((ConvertTo-ReportHtml @([pscustomobject]@{Nr=1;Mailbox='m';FromName='<b>x</b>';FromAddress='a';Subject='s';Received=$today;Unread=$false;Attachments=@();Body='b'}) $today) -match '&lt;b&gt;x&lt;/b&gt;')

Remove-Item $tmp -Recurse -Force
Write-Host ''
Write-Host "Passed: $script:passed  Failed: $script:failed"
if ($script:failed -gt 0) { exit 1 }
