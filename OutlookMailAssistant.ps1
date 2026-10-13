# =====================================================================
# Outlook Mail Assistant
# Daily e-mail report, guided attachment printing and template replies
# for every mailbox open in the local Outlook profile.
#
# Runs only when you start it. It never deletes or moves e-mails and
# never marks them as read.
#
#   .\OutlookMailAssistant.ps1 report      # HTML + Excel report for a day
#   .\OutlookMailAssistant.ps1 print       # print attachments, asks each time
#   .\OutlookMailAssistant.ps1 reply       # reply with a template or own text
#   .\OutlookMailAssistant.ps1 demo        # HTML report from fictitious e-mails
#                                          # (no Outlook needed)
#
# Requirements: Windows PowerShell 5.1 (built into Windows) and Outlook.
# Excel is optional (only for the .xlsx report). No modules, no admin
# rights, no internet access needed.
# =====================================================================
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('report', 'print', 'reply', 'demo')]
    [string]$Command = 'report',

    # Output folder (default: "Reports" next to the script)
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputFolder) { $OutputFolder = Join-Path $ScriptDir 'Reports' }
$DataFolder = Join-Path $OutputFolder 'app_data'   # CSV files used internally

# ---------------------------------------------------------------------
# Settings - defaults, overridden by config.json next to the script
# ---------------------------------------------------------------------
$Config = [ordered]@{
    Signature            = 'Your Name'
    Greeting             = 'Hello,'
    Closing              = 'Best regards,'
    Templates            = [ordered]@{
        '1' = 'Please send the missing document as soon as possible.'
        '2' = 'We have received your documents. Thank you.'
        '3' = 'The attached document is not readable. Please send it again.'
    }
    PrintableExtensions  = @('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.rtf', '.txt',
                             '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp')
    # Regular expressions matched against the subject; matching e-mails are
    # left out of the report (automatic delivery receipts and similar).
    IgnoredSubjects      = @(
        'Successful Mail Delivery Report',
        'Delivery Status Notification \((Relay|Delay|Success)\)',
        '^\s*(Delivered|Relayed)\s*:'
    )
    MaxBodyCharacters    = 5000
}

function Import-AssistantConfig {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    $json = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $json.PSObject.Properties) {
        if ($p.Name -like '_*') { continue }
        if ($p.Name -eq 'Templates') {
            $t = [ordered]@{}
            foreach ($tp in $p.Value.PSObject.Properties) { $t[$tp.Name] = "$($tp.Value)" }
            $Config.Templates = $t
        } else {
            $Config[$p.Name] = $p.Value
        }
    }
}

$PR_SMTP_SENDER = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$PR_ATTACHMENT_HIDDEN = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'

# ---------------------------------------------------------------------
# Pure helpers (no Outlook needed - covered by tests\Run-Tests.ps1)
# ---------------------------------------------------------------------
function ConvertTo-Day {
    # '' = today, 'y' = yesterday, or dd.MM.yyyy. Returns $null if invalid.
    param([string]$Text, [datetime]$Today = (Get-Date).Date)
    $s = "$Text".Trim().ToLowerInvariant()
    if ($s -eq '') { return $Today }
    if ($s -eq 'y') { return $Today.AddDays(-1) }
    try {
        return [datetime]::ParseExact($s, 'dd.MM.yyyy', [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $null
    }
}

function Test-IgnoredSubject {
    param([string]$Subject, [string[]]$Patterns = $Config.IgnoredSubjects)
    foreach ($m in $Patterns) { if ($Subject -match $m) { return $true } }
    return $false
}

function Get-SafeFileName {
    param([string]$Name)
    return ($Name -replace '[\\/:*?"<>|]', '_')
}

function Protect-SpreadsheetValue {
    # Keeps a value as plain text in Excel (blocks formula injection)
    param([string]$Value)
    if ($Value -match '^[=+\-@]') { return "'" + $Value }
    return $Value
}

function Limit-Text {
    param([string]$Text, [int]$Max, [string]$Suffix = ' [...]')
    if ($Text.Length -gt $Max) { return $Text.Substring(0, $Max) + $Suffix }
    return $Text
}

function New-ReplyText {
    param([string]$Body)
    return "$($Config.Greeting)`r`n`r`n$Body`r`n`r`n$($Config.Closing)`r`n$($Config.Signature)"
}

function Get-IndexPath  { param([datetime]$Day) Join-Path $DataFolder ("emails_{0}.csv" -f $Day.ToString('yyyy-MM-dd')) }
function Get-ExcelPath  { param([datetime]$Day) Join-Path $OutputFolder ("emails_{0}.xlsx" -f $Day.ToString('yyyy-MM-dd')) }
function Get-ReportPath { param([datetime]$Day) Join-Path $OutputFolder ("report_{0}.html" -f $Day.ToString('yyyy-MM-dd')) }

function Initialize-Folders {
    foreach ($f in @($OutputFolder, $DataFolder)) {
        if (-not (Test-Path $f)) { New-Item -ItemType Directory -Path $f | Out-Null }
    }
}

# ---------------------------------------------------------------------
# Outlook helpers
# ---------------------------------------------------------------------
function Get-Outlook {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        try {
            return New-Object -ComObject Outlook.Application
        } catch {
            throw ('Outlook could not be started. Open Outlook first and make sure a mailbox is ' +
                   'set up and fully loaded (no "Add Account" window), then try again. ' +
                   'If Outlook runs "as administrator", run this tool the same way. ' +
                   "Details: $($_.Exception.Message)")
        }
    }
}

function Read-Day {
    while ($true) {
        $d = ConvertTo-Day (Read-Host 'Day (Enter = today, y = yesterday, or dd.mm.yyyy)')
        if ($d) { return $d }
        Write-Host '  Invalid date. Example: 05.10.2026' -ForegroundColor Yellow
    }
}

function Get-SenderAddress {
    param($Item)
    $a = $null
    try { if ($Item.SenderEmailType -eq 'SMTP') { $a = $Item.SenderEmailAddress } } catch {}
    if (-not $a) { try { $a = $Item.PropertyAccessor.GetProperty($PR_SMTP_SENDER) } catch {} }
    if (-not $a) { try { $a = $Item.Sender.GetExchangeUser().PrimarySmtpAddress } catch {} }
    return "$a"
}

function Get-RealAttachments {
    # Only real attachments - skips inline images from signatures
    param($Item)
    $result = @()
    try { $n = $Item.Attachments.Count } catch { return $result }
    for ($k = 1; $k -le $n; $k++) {
        try {
            $a = $Item.Attachments.Item($k)
            if ($a.Type -ne 1 -and $a.Type -ne 5) { continue }
            $hidden = $false
            try { $hidden = [bool]$a.PropertyAccessor.GetProperty($PR_ATTACHMENT_HIDDEN) } catch {}
            if ($hidden) { continue }
            if ($a.FileName -match '^image\d+\.(png|jpe?g|gif)$') { continue }
            $result += [pscustomobject]@{ Name = $a.FileName; KB = [math]::Round($a.Size / 1KB) }
        } catch { continue }
    }
    return $result
}

function Set-ReplySender {
    # The reply leaves from the mailbox the e-mail arrived in
    param($Outlook, $Mail, [string]$Mailbox)
    try {
        $default = $Outlook.Session.Accounts.Item(1).SmtpAddress
        if ($default -and $default.ToLowerInvariant() -eq $Mailbox.ToLowerInvariant()) { return }
    } catch {}
    foreach ($acc in $Outlook.Session.Accounts) {
        if ($acc.SmtpAddress -and $acc.SmtpAddress.ToLowerInvariant() -eq $Mailbox.ToLowerInvariant()) {
            try { $Mail.SendUsingAccount = $acc; return } catch {}
        }
    }
    if ($Mailbox -like '*@*') { try { $Mail.SentOnBehalfOfName = $Mailbox } catch {} }
}

# ---------------------------------------------------------------------
# Excel output (through the installed Excel, optional)
# ---------------------------------------------------------------------
function Save-Excel {
    param([object[]]$Rows, [string[]]$Columns, [string]$Path, [hashtable]$Widths, [string]$WrapColumn = '')
    if (-not $Rows -or $Rows.Count -eq 0) { return $false }
    $xl = $null
    try {
        if (Test-Path $Path) { Remove-Item $Path -Force }
    } catch {
        Write-Host "  Cannot write $Path - is it open in Excel? Close it and try again." -ForegroundColor Yellow
        return $false
    }
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $nr = $Rows.Count + 1
        $nc = $Columns.Count
        $arr = New-Object 'object[,]' $nr, $nc
        for ($c = 0; $c -lt $nc; $c++) { $arr[0, $c] = $Columns[$c] }
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            for ($c = 0; $c -lt $nc; $c++) {
                $v = "$($Rows[$r].($Columns[$c]))"
                if ($v -match '^\d{1,9}$') { $arr[($r + 1), $c] = [int64]$v; continue }
                $arr[($r + 1), $c] = Protect-SpreadsheetValue $v
            }
        }
        $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nr, $nc))
        $range.Value2 = $arr
        $range.VerticalAlignment = -4160          # top
        $header = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $nc))
        $header.Font.Bold = $true
        $header.Interior.Color = 15652797          # light blue
        for ($c = 0; $c -lt $nc; $c++) {
            $col = $ws.Columns.Item($c + 1)
            if ($Widths -and $Widths.ContainsKey($Columns[$c])) { $col.ColumnWidth = $Widths[$Columns[$c]] }
            if ($Columns[$c] -eq $WrapColumn) { $col.WrapText = $true }
        }
        try { [void]$range.AutoFilter(1) } catch {}
        try {
            $ws.Activate()
            $xl.ActiveWindow.SplitRow = 1
            $xl.ActiveWindow.FreezePanes = $true
        } catch {}
        $wb.SaveAs($Path, 51)
        $wb.Close($false)
        return $true
    } catch {
        Write-Host "  Could not create the Excel file: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    } finally {
        if ($xl) {
            try { $xl.Quit() } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Update-ReplyLogExcel {
    param([string]$Log)
    if (-not (Test-Path $Log)) { return }
    $rows = @(Import-Csv -Path $Log -Delimiter ';' -Encoding UTF8)
    $path = Join-Path $OutputFolder 'sent_replies.xlsx'
    $ok = Save-Excel $rows @('Date', 'Mailbox', 'To', 'Subject', 'Text', 'Status') $path `
        @{ Date = 17; Mailbox = 28; To = 32; Subject = 45; Text = 70; Status = 14 } 'Text'
    if ($ok) { Write-Host "Reply log: $path" -ForegroundColor Green }
}

# ---------------------------------------------------------------------
# Reading the e-mails of one day
# ---------------------------------------------------------------------
function Get-DayEmails {
    param($Namespace, [datetime]$Day)
    $start = $Day.Date
    $end = $start.AddDays(1)
    $list = New-Object Collections.Generic.List[object]
    $skipped = 0

    foreach ($store in $Namespace.Stores) {
        $mailbox = $store.DisplayName
        try { $inbox = $store.GetDefaultFolder(6) } catch { continue }   # 6 = Inbox
        Write-Host "  Reading $mailbox / Inbox..." -ForegroundColor Cyan
        $items = $inbox.Items
        $items.Sort('[ReceivedTime]', $true)
        $it = $items.GetFirst()
        while ($null -ne $it) {
            try {
                $rt = $it.ReceivedTime
                if ($rt -lt $start) { break }
                if ($rt -lt $end -and $it.Class -eq 43) {             # 43 = MailItem
                    if (Test-IgnoredSubject "$($it.Subject)") {
                        $skipped++
                    } else {
                        $list.Add([pscustomobject]@{
                            Nr          = 0
                            Mailbox     = $mailbox
                            FromName    = "$($it.SenderName)"
                            FromAddress = Get-SenderAddress $it
                            Subject     = "$($it.Subject)"
                            Received    = $rt
                            Unread      = [bool]$it.UnRead
                            Attachments = @(Get-RealAttachments $it)
                            Body        = Limit-Text "$($it.Body)" ([int]$Config.MaxBodyCharacters) "`r`n[... shortened ...]"
                            EntryID     = $it.EntryID
                            StoreID     = $store.StoreID
                        })
                    }
                }
            } catch {}
            $it = $items.GetNext()
        }
    }
    if ($skipped -gt 0) { Write-Host "  Skipped $skipped automatic delivery receipts." -ForegroundColor DarkGray }
    return Set-EmailNumbers $list
}

function Set-EmailNumbers {
    param($Emails)
    $sorted = @($Emails | Sort-Object Received)
    for ($i = 0; $i -lt $sorted.Count; $i++) { $sorted[$i].Nr = $i + 1 }
    return , $sorted
}

function Save-Index {
    param($Emails, [datetime]$Day)
    $Emails | ForEach-Object {
        [pscustomobject]@{
            Nr          = $_.Nr
            Time        = $_.Received.ToString('HH:mm')
            Mailbox     = $_.Mailbox
            FromName    = $_.FromName
            FromAddress = $_.FromAddress
            Subject     = $_.Subject
            Attachments = (($_.Attachments | ForEach-Object { $_.Name }) -join ' | ')
            EntryID     = $_.EntryID
            StoreID     = $_.StoreID
        }
    } | Export-Csv -Path (Get-IndexPath $Day) -Delimiter ';' -NoTypeInformation -Encoding UTF8
}

function Get-Index {
    # Reads the day's list; creates it first if it does not exist
    param($Namespace, [datetime]$Day)
    $f = Get-IndexPath $Day
    if (-not (Test-Path $f)) {
        Write-Host 'No list for this day yet. Creating it now...' -ForegroundColor Cyan
        Save-Index (Get-DayEmails $Namespace $Day) $Day
    }
    return @(Import-Csv -Path $f -Delimiter ';' -Encoding UTF8)
}

# ---------------------------------------------------------------------
# HTML report (pure - also used by demo mode)
# ---------------------------------------------------------------------
function ConvertTo-ReportHtml {
    param($Emails, [datetime]$Day, [datetime]$Generated = (Get-Date))
    $enc = { param($t) [Net.WebUtility]::HtmlEncode("$t") }
    $dayText = $Day.ToString('dd.MM.yyyy')
    $withAtt = @($Emails | Where-Object { $_.Attachments.Count -gt 0 }).Count
    $unread = @($Emails | Where-Object { $_.Unread }).Count

    $sb = New-Object Text.StringBuilder
    [void]$sb.Append(@"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<title>E-mails $dayText</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f3f4f6;margin:0;padding:24px;color:#1f2937}
h1{margin:0 0 4px;font-size:22px} .summary{color:#4b5563;margin-bottom:20px}
.toc{background:#fff;border-radius:8px;padding:12px 16px;margin-bottom:20px}
.toc table{border-collapse:collapse;width:100%;font-size:14px}
.toc td{padding:4px 8px;border-bottom:1px solid #eee;vertical-align:top}
.toc a{color:#1d4ed8;text-decoration:none}
.email{background:#fff;border-radius:8px;padding:16px;margin-bottom:16px;border-left:5px solid #9ca3af}
.email.new{border-left-color:#2563eb}
.head{display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;font-size:13px;color:#6b7280}
.subject{font-size:17px;font-weight:600;margin:6px 0}
.from{font-size:14px;margin-bottom:8px}
.att{background:#fef3c7;border-radius:6px;padding:6px 10px;font-size:13px;margin:8px 0}
pre{white-space:pre-wrap;word-wrap:break-word;font-family:inherit;font-size:14px;background:#f9fafb;padding:10px;border-radius:6px;margin:0}
.tag{background:#2563eb;color:#fff;border-radius:4px;padding:1px 6px;font-size:12px}
</style></head><body>
<h1>E-mails received - $dayText</h1>
<div class="summary">$($Emails.Count) e-mails, $unread unread, $withAtt with attachments. Generated $($Generated.ToString('dd.MM.yyyy HH:mm')).</div>
"@)

    if ($Emails.Count -gt 0) {
        [void]$sb.Append('<div class="toc"><table>')
        foreach ($e in $Emails) {
            $clip = if ($e.Attachments.Count -gt 0) { ' &#128206;' } else { '' }
            [void]$sb.Append(("<tr><td>#{0}</td><td>{1}</td><td>{2}</td><td><a href='#e{0}'>{3}</a>{4}</td><td>{5}</td></tr>" -f `
                $e.Nr, $e.Received.ToString('HH:mm'), (& $enc $e.FromName), (& $enc $e.Subject), $clip, (& $enc $e.Mailbox)))
        }
        [void]$sb.Append('</table></div>')
    }

    foreach ($e in $Emails) {
        $class = if ($e.Unread) { 'email new' } else { 'email' }
        $new = if ($e.Unread) { ' <span class="tag">unread</span>' } else { '' }
        [void]$sb.Append("<div class='$class' id='e$($e.Nr)'>")
        [void]$sb.Append(("<div class='head'><span>#{0} &middot; {1}{2}</span><span>{3}</span></div>" -f `
            $e.Nr, $e.Received.ToString('dd.MM.yyyy HH:mm'), $new, (& $enc $e.Mailbox)))
        [void]$sb.Append("<div class='subject'>$(& $enc $e.Subject)</div>")
        [void]$sb.Append("<div class='from'>From: <b>$(& $enc $e.FromName)</b> &lt;$(& $enc $e.FromAddress)&gt;</div>")
        if ($e.Attachments.Count -gt 0) {
            $l = ($e.Attachments | ForEach-Object { "$(& $enc $_.Name) ($($_.KB) KB)" }) -join ' &middot; '
            [void]$sb.Append("<div class='att'>&#128206; Attachments: $l</div>")
        }
        [void]$sb.Append("<pre>$(& $enc "$($e.Body)".Trim())</pre></div>")
    }
    if ($Emails.Count -eq 0) { [void]$sb.Append('<p>No e-mails on this day.</p>') }
    [void]$sb.Append('</body></html>')
    return $sb.ToString()
}

function Save-Report {
    param($Emails, [datetime]$Day)
    $file = Get-ReportPath $Day
    [IO.File]::WriteAllText($file, (ConvertTo-ReportHtml $Emails $Day), (New-Object Text.UTF8Encoding($true)))
    return $file
}

function Open-File {
    param([string]$Path)
    if ($IsLinux -or $IsMacOS) { Write-Host "  (open manually: $Path)"; return }
    Start-Process $Path
}

# =====================================================================
# 1. REPORT
# =====================================================================
function Invoke-Report {
    $day = Read-Day
    $outlook = Get-Outlook
    $ns = $outlook.GetNamespace('MAPI')
    $emails = Get-DayEmails $ns $day
    Save-Index $emails $day

    $xlsx = Get-ExcelPath $day
    $okXlsx = $false
    if ($emails.Count -gt 0) {
        Write-Host '  Creating the Excel file...' -ForegroundColor Cyan
        $rows = $emails | ForEach-Object {
            [pscustomobject]@{
                'Nr'          = $_.Nr
                'Time'        = $_.Received.ToString('HH:mm')
                'Mailbox'     = $_.Mailbox
                'From'        = $_.FromName
                'Address'     = $_.FromAddress
                'Subject'     = $_.Subject
                'Attachments' = (($_.Attachments | ForEach-Object { $_.Name }) -join "`n")
                'Unread'      = $(if ($_.Unread) { 'YES' } else { '' })
                'Text'        = Limit-Text (($_.Body -replace "`r`n", "`n").Trim()) 1500
            }
        }
        $okXlsx = Save-Excel @($rows) @('Nr', 'Time', 'Mailbox', 'From', 'Address', 'Subject', 'Attachments', 'Unread', 'Text') `
            $xlsx @{ 'Nr' = 5; 'Time' = 7; 'Mailbox' = 26; 'From' = 24; 'Address' = 30; 'Subject' = 40; 'Attachments' = 30; 'Unread' = 8; 'Text' = 80 } 'Text'
    }

    $file = Save-Report $emails $day
    $withAtt = @($emails | Where-Object { $_.Attachments.Count -gt 0 }).Count
    Write-Host ''
    Write-Host "Done: $($emails.Count) e-mails ($withAtt with attachments)." -ForegroundColor Green
    Write-Host "Report (browser): $file"
    Open-File $file
    if ($okXlsx) {
        Write-Host "Excel: $xlsx" -ForegroundColor Green
        if ("$(Read-Host 'Open the Excel file too? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') { Open-File $xlsx }
    }
    if ($withAtt -gt 0) {
        if ("$(Read-Host 'Print attachments now? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') { Invoke-Print -Day $day -Namespace $ns }
    }
}

# =====================================================================
# 2. PRINT ATTACHMENTS (asks for each one)
# =====================================================================
function Invoke-Print {
    param([Nullable[datetime]]$Day = $null, $Namespace = $null)
    if (-not $Day) { $Day = Read-Day }
    if (-not $Namespace) { $Namespace = (Get-Outlook).GetNamespace('MAPI') }
    $index = Get-Index $Namespace $Day
    $withAtt = @($index | Where-Object { "$($_.Attachments)".Trim() -ne '' })
    if ($withAtt.Count -eq 0) {
        Write-Host 'No e-mails with attachments on the chosen day.' -ForegroundColor Yellow
        return
    }

    $attFolder = Join-Path $OutputFolder ("attachments_{0}" -f $Day.ToString('yyyy-MM-dd'))
    if (-not (Test-Path $attFolder)) { New-Item -ItemType Directory -Path $attFolder | Out-Null }

    Write-Host ''
    Write-Host 'For each attachment: P = print, N = no, V = view it, S = skip e-mail, Q = quit' -ForegroundColor Cyan
    $printed = 0

    foreach ($e in $withAtt) {
        Write-Host ''
        Write-Host ("#{0}  {1}  {2}" -f $e.Nr, $e.Time, $e.FromName) -ForegroundColor White
        Write-Host ("    {0}" -f $e.Subject) -ForegroundColor Gray
        try { $it = $Namespace.GetItemFromID($e.EntryID, $e.StoreID) } catch {
            Write-Host '  E-mail not found any more (moved or deleted?).' -ForegroundColor Yellow
            continue
        }
        $skipEmail = $false
        foreach ($attName in ("$($e.Attachments)" -split ' \| ')) {
            if ($skipEmail) { break }
            $att = $null
            for ($k = 1; $k -le $it.Attachments.Count; $k++) {
                if ($it.Attachments.Item($k).FileName -eq $attName) { $att = $it.Attachments.Item($k); break }
            }
            if (-not $att) { continue }

            $path = Join-Path $attFolder ("{0}_{1}" -f $e.Nr, (Get-SafeFileName $attName))
            $ext = [IO.Path]::GetExtension($attName).ToLowerInvariant()

            while ($true) {
                $r = "$(Read-Host ("    {0}  -> print? (P/N/V/S/Q)" -f $attName))".Trim().ToUpperInvariant()
                if ($r -eq 'Q') { Write-Host "Stopped. Printed: $printed." -ForegroundColor Cyan; return }
                if ($r -eq 'S') { $skipEmail = $true; break }
                if ($r -eq 'N' -or $r -eq '') { break }
                if (-not (Test-Path $path)) { $att.SaveAsFile($path) }
                if ($r -eq 'V') { Open-File $path; continue }
                if ($r -eq 'P') {
                    if ($ext -eq '.msg') {
                        # Attached e-mail (.msg): printed by Outlook in a separate job,
                        # so the script does not hang if Outlook stops responding
                        $job = Start-Job -ArgumentList $path -ScriptBlock {
                            param($f)
                            $ol = New-Object -ComObject Outlook.Application
                            $m = $ol.GetNamespace('MAPI').OpenSharedItem($f)
                            $m.PrintOut()
                            Start-Sleep -Seconds 2
                            try { $m.Close(1) } catch {}
                        }
                        $done = Wait-Job $job -Timeout 45
                        if ($done -and $job.State -eq 'Completed') {
                            $printed++
                            Write-Host '      sent to printer' -ForegroundColor Green
                        } else {
                            if (-not $done) { Stop-Job $job }
                            Write-Host '      Could not print (Outlook did not respond). Press V to open it and print manually.' -ForegroundColor Yellow
                        }
                        Remove-Job $job -Force
                    } elseif ($Config.PrintableExtensions -contains $ext) {
                        try {
                            Start-Process -FilePath $path -Verb Print
                            $printed++
                            Write-Host '      sent to printer' -ForegroundColor Green
                            Start-Sleep -Seconds 3
                        } catch {
                            Write-Host "      Could not print. The file is saved at: $path" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "      Cannot print $ext files automatically. Use V to open and print it. File: $path" -ForegroundColor Yellow
                    }
                    break
                }
            }
        }
    }
    Write-Host ''
    Write-Host "Done. Attachments printed: $printed. Copies are in: $attFolder" -ForegroundColor Green
}

# =====================================================================
# 3. REPLY
# =====================================================================
function Invoke-Reply {
    $day = Read-Day
    $outlook = Get-Outlook
    $ns = $outlook.GetNamespace('MAPI')
    $index = Get-Index $ns $day
    if ($index.Count -eq 0) { Write-Host 'No e-mails on the chosen day.' -ForegroundColor Yellow; return }
    $log = Join-Path $DataFolder 'sent_replies.csv'

    while ($true) {
        Write-Host ''
        Write-Host ("E-mails of {0}:" -f $day.ToString('dd.MM.yyyy')) -ForegroundColor Cyan
        foreach ($e in $index) {
            Write-Host ("  #{0,-3} {1}  {2,-28} {3}" -f $e.Nr, $e.Time, (Limit-Text $e.FromName 28 ''), (Limit-Text $e.Subject 55 ''))
        }
        Write-Host ''
        $nr = "$(Read-Host 'Number of the e-mail to reply to (Enter = exit)')".Trim()
        if ($nr -eq '') { Update-ReplyLogExcel $log; return }
        $e = $index | Where-Object { $_.Nr -eq $nr } | Select-Object -First 1
        if (-not $e) { Write-Host 'No such number.' -ForegroundColor Yellow; continue }

        try { $it = $ns.GetItemFromID($e.EntryID, $e.StoreID) } catch {
            Write-Host 'E-mail not found any more (moved or deleted?).' -ForegroundColor Yellow
            continue
        }

        Write-Host ''
        Write-Host "From:    $($e.FromName) <$($e.FromAddress)>" -ForegroundColor White
        Write-Host "Subject: $($e.Subject)" -ForegroundColor White
        Write-Host "Mailbox: $($e.Mailbox)" -ForegroundColor Gray
        Write-Host '----------------------------------------------------------'
        Write-Host (Limit-Text "$($it.Body)".Trim() 800)
        Write-Host '----------------------------------------------------------'

        Write-Host 'Which reply?' -ForegroundColor Cyan
        foreach ($k in $Config.Templates.Keys) { Write-Host ("  {0} = {1}" -f $k, $Config.Templates[$k]) }
        Write-Host '  W = write my own text'
        Write-Host '  C = cancel'
        $opt = "$(Read-Host 'Choose')".Trim().ToUpperInvariant()
        if ($opt -eq 'C' -or $opt -eq '') { continue }
        if ($opt -eq 'W') {
            Write-Host 'Type the text. Press Enter on an empty line to finish.' -ForegroundColor Gray
            $lines = @()
            while ($true) {
                $l = Read-Host
                if ($l -eq '') { break }
                $lines += $l
            }
            if ($lines.Count -eq 0) { continue }
            $text = $lines -join "`r`n"
        } elseif ($Config.Templates.Contains($opt)) {
            $text = $Config.Templates[$opt]
        } else {
            Write-Host 'Unknown option.' -ForegroundColor Yellow
            continue
        }

        $reply = $it.Reply()
        Set-ReplySender $outlook $reply $e.Mailbox
        $fullText = New-ReplyText $text
        try {
            if ($reply.BodyFormat -eq 2) {                 # HTML
                $html = [Net.WebUtility]::HtmlEncode($fullText) -replace "`r`n", '<br>'
                $h = $reply.HTMLBody
                $m = [regex]::Match($h, '<body[^>]*>', 'IgnoreCase')
                $block = "<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt'>$html</div><br>"
                if ($m.Success) { $reply.HTMLBody = $h.Insert($m.Index + $m.Length, $block) }
                else { $reply.HTMLBody = $block + $h }
            } else {
                $reply.Body = $fullText + "`r`n`r`n" + $reply.Body
            }
        } catch {
            $reply.Body = $fullText + "`r`n`r`n" + $reply.Body
        }

        Write-Host ''
        Write-Host "To:      $($reply.To)" -ForegroundColor White
        Write-Host "Subject: $($reply.Subject)" -ForegroundColor White
        Write-Host $fullText -ForegroundColor Gray
        $fin = "$(Read-Host 'D = save as draft (check it in Outlook), S = send now, C = cancel')".Trim().ToUpperInvariant()

        $status = $null
        if ($fin -eq 'S') {
            if ("$(Read-Host 'Type YES to send')".Trim().ToUpperInvariant() -eq 'YES') {
                try { $reply.Send(); $status = 'sent' } catch { $status = "error: $($_.Exception.Message)" }
            } else { $fin = 'D' }
        }
        if ($fin -eq 'D') {
            try { $reply.Save(); $status = 'draft' } catch { $status = "error: $($_.Exception.Message)" }
        }
        if (-not $status) {
            try { $reply.Close(1) } catch {}
            Write-Host 'Cancelled.' -ForegroundColor Gray
            continue
        }

        $color = if ($status -like 'error*') { 'Red' } else { 'Green' }
        Write-Host "  -> $status" -ForegroundColor $color
        [pscustomobject]@{
            Date    = Get-Date -Format 'dd.MM.yyyy HH:mm'
            Mailbox = $e.Mailbox
            To      = $e.FromAddress
            Subject = $e.Subject
            Text    = ($text -replace "[\r\n]+", ' ')
            Status  = $status
        } | Export-Csv -Path $log -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Append
    }
}

# =====================================================================
# DEMO - fictitious e-mails, no Outlook needed
# =====================================================================
function New-DemoEmails {
    param([datetime]$Day)
    $raw = @(
        @('Inbox - office@example.com', 'Maria Ionescu', 'maria.ionescu@example.org', 'Request for certificate', 9, 12, $true,
          @(@{ Name = 'request_signed.pdf'; KB = 184 }), "Hello,`r`n`r`nPlease find attached the signed request.`r`n`r`nThank you,`r`nMaria"),
        @('Inbox - office@example.com', 'John Smith', 'j.smith@example.net', 'Question about my file', 10, 47, $false,
          @(), "Good morning,`r`n`r`nCould you tell me what documents are still missing from my file?`r`n`r`nRegards,`r`nJohn"),
        @('Inbox - registry@example.com', 'Ana Popescu', 'ana.p@example.org', 'Scanned documents', 13, 5, $true,
          @(@{ Name = 'scan_page1.jpg'; KB = 912 }, @{ Name = 'scan_page2.jpg'; KB = 876 }), "Hello,`r`n`r`nAttached are the two pages you asked for.`r`n`r`nAna"),
        @('Inbox - registry@example.com', 'Peter Novak', 'p.novak@example.com', 'Forwarded: confirmation', 15, 31, $false,
          @(@{ Name = 'confirmation.msg'; KB = 64 }), "Hi,`r`n`r`nForwarding the confirmation e-mail below.`r`n`r`nPeter")
    )
    $list = foreach ($r in $raw) {
        [pscustomobject]@{
            Nr = 0; Mailbox = $r[0]; FromName = $r[1]; FromAddress = $r[2]; Subject = $r[3]
            Received = $Day.Date.AddHours($r[4]).AddMinutes($r[5]); Unread = $r[6]
            Attachments = @($r[7] | ForEach-Object { [pscustomobject]$_ }); Body = $r[8]
            EntryID = ''; StoreID = ''
        }
    }
    return Set-EmailNumbers $list
}

function Invoke-Demo {
    $day = (Get-Date).Date
    $emails = New-DemoEmails $day
    $file = Save-Report $emails $day
    Write-Host "Demo report with $($emails.Count) fictitious e-mails: $file" -ForegroundColor Green
    Open-File $file
}

# ---------------------------------------------------------------------
# Start (skipped when the script is dot-sourced by the tests)
# ---------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-AssistantConfig (Join-Path $ScriptDir 'config.json')
    try {
        Initialize-Folders
        switch ($Command) {
            'report' { Invoke-Report }
            'print'  { Invoke-Print }
            'reply'  { Invoke-Reply }
            'demo'   { Invoke-Demo }
        }
    } catch {
        Write-Host ''
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
