# =====================================================================
# Outlook Mail Assistant
# Daily e-mail report, guided attachment printing and template replies.
#
# Mail sources (any combination, set in config.json or with "setup"):
#   outlook  - every mailbox open in the local Outlook profile (default)
#   imap     - Gmail, Yahoo or any IMAP mailbox (app password)
#   folder   - a folder of saved .eml files
#   mbox     - Thunderbird mail folders (mbox files)
#
# Runs only when you start it. It never deletes or moves e-mails and
# never marks them as read.
#
#   .\OutlookMailAssistant.ps1 report      # HTML + Excel report for a day
#   .\OutlookMailAssistant.ps1 print       # print attachments, asks each time
#   .\OutlookMailAssistant.ps1 reply       # reply with a template or own text
#   .\OutlookMailAssistant.ps1 setup       # add / remove mailboxes
#   .\OutlookMailAssistant.ps1 demo        # HTML report from fictitious e-mails
#
# Requirements: Windows PowerShell 5.1 (built into Windows). Outlook only
# for the "outlook" source, Excel only for the .xlsx files. No modules,
# no admin rights.
# =====================================================================
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('report', 'print', 'reply', 'setup', 'demo')]
    [string]$Command = 'report',

    # Output folder (default: "Reports" next to the script)
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputFolder) { $OutputFolder = Join-Path $ScriptDir 'Reports' }
$DataFolder = Join-Path $OutputFolder 'app_data'   # CSV files and message copies used internally
$ConfigPath = Join-Path $ScriptDir 'config.json'
$CredentialFolder = Join-Path $ScriptDir 'credentials'

# ---------------------------------------------------------------------
# Settings - defaults, overridden by config.json next to the script
# ---------------------------------------------------------------------
$Config = [ordered]@{
    Sources              = @([pscustomobject]@{ Type = 'outlook' })
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
    # IMAP: larger messages are listed with their headers only
    MaxMessageMB         = 25
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
        } elseif ($p.Name -eq 'Sources') {
            $Config.Sources = @($p.Value | Where-Object { $null -ne $_ })
        } else {
            $Config[$p.Name] = $p.Value
        }
    }
}

$PR_SMTP_SENDER = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$PR_ATTACHMENT_HIDDEN = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'

# Text encodings: Latin-1 maps bytes 1:1 to characters, so raw e-mail
# data can be handled as text without losing anything.
$Latin1 = [Text.Encoding]::GetEncoding(28591)
$StrictUtf8 = New-Object Text.UTF8Encoding($false, $true)
$codePages = 'System.Text.CodePagesEncodingProvider' -as [type]
if ($codePages) { try { [Text.Encoding]::RegisterProvider($codePages::Instance) } catch { } }

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
function Get-CachePath  { param([datetime]$Day) Join-Path $DataFolder ("messages_{0}" -f $Day.ToString('yyyy-MM-dd')) }
function Get-ExcelPath  { param([datetime]$Day) Join-Path $OutputFolder ("emails_{0}.xlsx" -f $Day.ToString('yyyy-MM-dd')) }
function Get-ReportPath { param([datetime]$Day) Join-Path $OutputFolder ("report_{0}.html" -f $Day.ToString('yyyy-MM-dd')) }

function Initialize-Folders {
    foreach ($f in @($OutputFolder, $DataFolder)) {
        if (-not (Test-Path $f)) { New-Item -ItemType Directory -Path $f | Out-Null }
    }
}

# ---------------------------------------------------------------------
# E-mail format (MIME / .eml) - pure PowerShell, no libraries
# ---------------------------------------------------------------------
function Get-TextEncoding {
    param([string]$Charset)
    $cs = "$Charset".Trim().Trim('"').ToLowerInvariant()
    $alias = @{ 'utf8' = 'utf-8'; 'cp1250' = 'windows-1250'; 'cp1252' = 'windows-1252'
                'latin1' = 'iso-8859-1'; 'latin2' = 'iso-8859-2'; 'ansi_x3.4-1968' = 'us-ascii' }
    if ($alias.ContainsKey($cs)) { $cs = $alias[$cs] }
    if (-not $cs) { return $null }
    try { return [Text.Encoding]::GetEncoding($cs) } catch { return $null }
}

function ConvertFrom-Bytes {
    # Bytes -> text. Without a usable charset: UTF-8 if valid, else Windows-1252.
    param([byte[]]$Bytes, [string]$Charset = '')
    if ($null -eq $Bytes) { return '' }
    $cs = "$Charset".Trim().Trim('"').ToLowerInvariant()
    if ($cs -and @('us-ascii', 'ascii', 'unknown-8bit', 'x-unknown') -notcontains $cs) {
        $enc = Get-TextEncoding $cs
        if ($enc) { return $enc.GetString($Bytes) }
    }
    try { return $StrictUtf8.GetString($Bytes) } catch { }
    $w = Get-TextEncoding 'windows-1252'
    if ($w) { return $w.GetString($Bytes) }
    return $Latin1.GetString($Bytes)
}

function ConvertFrom-Base64Loose {
    param([string]$Text)
    $c = $Text -replace '[^A-Za-z0-9+/]', ''
    $rest = $c.Length % 4
    if ($rest -eq 1) { $c = $c.Substring(0, $c.Length - 1) }
    elseif ($rest -gt 0) { $c += '=' * (4 - $rest) }
    try { return , [Convert]::FromBase64String($c) } catch { return , [byte[]]@() }
}

function ConvertFrom-QuotedPrintable {
    # Returns bytes (the charset is applied by the caller)
    param([string]$Text)
    $t = $Text -replace '=\r?\n', ''
    $t = [regex]::Replace($t, '=([0-9A-Fa-f]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
    return , $Latin1.GetBytes($t)
}

function ConvertFrom-EncodedWord {
    # Decodes RFC 2047 words such as =?utf-8?B?...?= in headers
    param([string]$Text)
    if ($Text -notmatch '=\?') { return $Text }
    $s = [regex]::Replace($Text, '(\?=)\s+(?==\?)', '$1')
    $found = [regex]::Matches($s, '=\?([^?\s]+)\?([bBqQ])\?([^?\s]*)\?=')
    if ($found.Count -eq 0) { return $Text }
    $sb = New-Object Text.StringBuilder
    $pending = New-Object Collections.Generic.List[byte]
    $pendingCharset = ''
    $last = 0
    foreach ($m in $found) {
        $cs = ($m.Groups[1].Value -split '\*')[0].ToLowerInvariant()
        if ($m.Index -ne $last -or ($pending.Count -gt 0 -and $cs -ne $pendingCharset)) {
            if ($pending.Count -gt 0) { [void]$sb.Append((ConvertFrom-Bytes $pending.ToArray() $pendingCharset)); $pending.Clear() }
            [void]$sb.Append($s.Substring($last, $m.Index - $last))
        }
        $pendingCharset = $cs
        if ($m.Groups[2].Value -ieq 'B') {
            $bytes = ConvertFrom-Base64Loose $m.Groups[3].Value
        } else {
            $bytes = ConvertFrom-QuotedPrintable ($m.Groups[3].Value -replace '_', ' ')
        }
        $pending.AddRange([byte[]]$bytes)
        $last = $m.Index + $m.Length
    }
    if ($pending.Count -gt 0) { [void]$sb.Append((ConvertFrom-Bytes $pending.ToArray() $pendingCharset)) }
    [void]$sb.Append($s.Substring($last))
    return $sb.ToString()
}

function ConvertFrom-RawHeader {
    # Header value as read from the message (Latin-1 text) -> readable text
    param([string]$Text)
    if (-not $Text) { return '' }
    if ($Text -match '[\x80-\xff]') { $Text = ConvertFrom-Bytes $Latin1.GetBytes($Text) '' }
    return (ConvertFrom-EncodedWord $Text).Trim()
}

function Split-MimeHeaderBody {
    param([string]$Text)
    $m = [regex]::Match($Text, '\r?\n\r?\n')
    if (-not $m.Success) { return @($Text, '') }
    return @($Text.Substring(0, $m.Index), $Text.Substring($m.Index + $m.Length))
}

function Get-MimeHeaders {
    # First occurrence of every header, keys in lower case
    param([string]$HeaderText)
    $h = @{}
    $unfolded = $HeaderText -replace '\r?\n(?=[ \t])', ''
    foreach ($l in ($unfolded -split '\r?\n')) {
        $i = $l.IndexOf(':')
        if ($i -le 0) { continue }
        $k = $l.Substring(0, $i).Trim().ToLowerInvariant()
        if (-not $h.ContainsKey($k)) { $h[$k] = $l.Substring($i + 1).Trim() }
    }
    return $h
}

function Get-HeaderParams {
    # 'text/plain; charset="utf-8"' -> Value = text/plain, Params = @{charset = utf-8}
    # Handles quoted values, RFC 2231 (filename*=utf-8''...) and continuations.
    param([string]$Value)
    $res = [pscustomobject]@{ Value = ''; Params = @{} }
    if (-not $Value) { return $res }
    $parts = $Value -split ';', 2
    $res.Value = $parts[0].Trim().Trim('"').ToLowerInvariant()
    if ($parts.Count -lt 2) { return $res }
    $segments = @{}
    foreach ($m in [regex]::Matches($parts[1], '([^\s=;]+)\s*=\s*("(?:[^"\\]|\\.)*"|[^;]*)')) {
        $k = $m.Groups[1].Value.ToLowerInvariant()
        $v = $m.Groups[2].Value.Trim()
        if ($v.StartsWith('"') -and $v.EndsWith('"') -and $v.Length -ge 2) {
            $v = $v.Substring(1, $v.Length - 2) -replace '\\(.)', '$1'
        }
        if ($k -match '^(.+?)\*(\d+)(\*?)$') {
            $name = $Matches[1]; $idx = [int]$Matches[2]; $enc = ($Matches[3] -eq '*')
        } elseif ($k -match '^(.+?)\*$') {
            $name = $Matches[1]; $idx = 0; $enc = $true
        } else {
            $res.Params[$k] = ConvertFrom-RawHeader $v
            continue
        }
        if (-not $segments.ContainsKey($name)) { $segments[$name] = @() }
        $segments[$name] += [pscustomobject]@{ Idx = $idx; Enc = $enc; Val = $v }
    }
    foreach ($name in $segments.Keys) {
        $charset = 'utf-8'
        $bytes = New-Object Collections.Generic.List[byte]
        foreach ($seg in ($segments[$name] | Sort-Object Idx)) {
            $v = $seg.Val
            if ($seg.Enc) {
                if ($seg.Idx -eq 0 -and $v -match "^([^']*)'[^']*'(.*)$") {
                    if ($Matches[1]) { $charset = $Matches[1] }
                    $v = $Matches[2]
                }
                $v = [regex]::Replace($v, '%([0-9A-Fa-f]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
            }
            $bytes.AddRange([byte[]]$Latin1.GetBytes($v))
        }
        $res.Params[$name] = ConvertFrom-Bytes $bytes.ToArray() $charset
    }
    return $res
}

function Split-MailAddress {
    # '"Name" <a@b.c>', 'a@b.c (Name)' or 'a@b.c' -> Name, Address
    param([string]$Text)
    $t = "$Text".Trim()
    if ($t -match '^\s*"?([^"<]*?)"?\s*<([^>]+)>') {
        $name = $Matches[1].Trim(); $addr = $Matches[2].Trim()
    } elseif ($t -match '([^\s<>()",;]+@[^\s<>()",;]+)\s*(?:\(([^)]*)\))?') {
        $addr = $Matches[1]; $name = "$($Matches[2])".Trim()
    } else {
        $addr = ''; $name = $t
    }
    if (-not $name) { $name = $addr }
    return [pscustomobject]@{ Name = $name; Address = $addr }
}

function ConvertFrom-MailDate {
    # 'Tue, 13 Oct 2026 09:12:00 +0300 (EEST)' -> local [datetime], or $null
    param([string]$Text)
    if (-not $Text) { return $null }
    $s = ($Text -replace '\([^)]*\)', ' ').Trim()
    $s = $s -replace '^[A-Za-z]{3,9},?\s*', ''
    $s = ($s -replace '\s+', ' ').Trim()
    $zones = @{ 'UT' = '+0000'; 'UTC' = '+0000'; 'GMT' = '+0000'; 'Z' = '+0000'
                'EST' = '-0500'; 'EDT' = '-0400'; 'CST' = '-0600'; 'CDT' = '-0500'
                'MST' = '-0700'; 'MDT' = '-0600'; 'PST' = '-0800'; 'PDT' = '-0700' }
    if ($s -match ' ([A-Za-z]{1,3})$' -and $zones.ContainsKey($Matches[1].ToUpperInvariant())) {
        $s = $s.Substring(0, $s.Length - $Matches[1].Length) + $zones[$Matches[1].ToUpperInvariant()]
    }
    if ($s -notmatch '[+-]\d{4}$') { $s += ' +0000' }
    $s = $s -replace '([+-]\d{2})(\d{2})$', '$1:$2'
    $formats = [string[]]@('d MMM yyyy H:mm:ss zzz', 'd MMM yyyy H:mm zzz', 'd MMM yy H:mm:ss zzz', 'd MMM yy H:mm zzz')
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParseExact($s, $formats, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dto)) {
        return $dto.LocalDateTime
    }
    return $null
}

function ConvertFrom-Html {
    param([string]$Html)
    $t = $Html -replace '(?is)<(head|style|script)\b.*?</\1\s*>', ''
    $t = $t -replace '(?i)<br\s*/?>', "`n"
    $t = $t -replace '(?i)</(p|div|tr|li|h[1-6]|table|blockquote)\s*>', "`n"
    $t = $t -replace '(?i)<li\b[^>]*>', '- '
    $t = $t -replace '(?s)<[^>]+>', ''
    $t = [Net.WebUtility]::HtmlDecode($t) -replace [char]0xA0, ' '
    $t = (($t -split '\r?\n') | ForEach-Object { $_.Trim() }) -join "`r`n"
    return ($t -replace '(\r\n){3,}', "`r`n`r`n").Trim()
}

function Split-Multipart {
    param([string]$Body, [string]$Boundary)
    $result = @()
    $rx = '(?:^|\r?\n)--' + [regex]::Escape($Boundary) + '(--)?[ \t]*(?=\r?\n|$)'
    $found = [regex]::Matches($Body, $rx)
    for ($i = 0; $i -lt $found.Count; $i++) {
        if ($found[$i].Groups[1].Success) { break }                    # closing --boundary--
        $start = $found[$i].Index + $found[$i].Length
        $end = if ($i + 1 -lt $found.Count) { $found[$i + 1].Index } else { $Body.Length }
        $part = $Body.Substring($start, $end - $start) -replace '^\r?\n', ''
        $result += , $part
    }
    return , $result
}

function Read-MimeEntity {
    param([string]$Text, [int]$Depth = 0)
    $hb = Split-MimeHeaderBody $Text
    $h = Get-MimeHeaders $hb[0]
    $ct = Get-HeaderParams $h['content-type']
    $cd = Get-HeaderParams $h['content-disposition']
    $type = if ($ct.Value -match '/') { $ct.Value } else { 'text/plain' }
    $name = $cd.Params['filename']
    if (-not $name) { $name = $ct.Params['name'] }
    $e = [pscustomobject]@{
        Headers     = $h
        Type        = $type
        Params      = $ct.Params
        Disposition = $cd.Value
        Encoding    = "$($h['content-transfer-encoding'])".Trim().ToLowerInvariant()
        Name        = "$name".Trim()
        Body        = $hb[1]
        Parts       = @()
    }
    if ($type -like 'multipart/*' -and $ct.Params['boundary'] -and $Depth -lt 20) {
        $e.Parts = @(foreach ($p in (Split-Multipart $hb[1] $ct.Params['boundary'])) { Read-MimeEntity $p ($Depth + 1) })
    }
    return $e
}

function Get-PartBytes {
    param($Entity)
    switch ($Entity.Encoding) {
        'base64'           { return , [byte[]](ConvertFrom-Base64Loose $Entity.Body) }
        'quoted-printable' { return , [byte[]](ConvertFrom-QuotedPrintable $Entity.Body) }
        default            { return , $Latin1.GetBytes($Entity.Body) }
    }
}

function Get-PartText {
    param($Entity)
    $t = ConvertFrom-Bytes (Get-PartBytes $Entity) $Entity.Params['charset']
    return ($t -replace '\r?\n', "`r`n")
}

function Add-MimeLeaves {
    param($Entity, $List)
    if ($Entity.Type -like 'multipart/*' -and $Entity.Parts.Count -gt 0) {
        foreach ($p in $Entity.Parts) { Add-MimeLeaves $p $List }
    } else {
        $List.Add($Entity)
    }
}

function Test-MimeAttachment {
    param($Entity)
    $t = $Entity.Type
    if ($t -like 'multipart/*') { return $false }
    if ($t -eq 'message/rfc822') { return $true }
    $embedded = [bool]$Entity.Headers['content-id'] -and $Entity.Disposition -ne 'attachment'
    if ($Entity.Disposition -eq 'attachment') { return $true }
    if ($Entity.Name) {
        if ($embedded -and $t -like 'image/*') { return $false }       # picture inside the text / signature
        if ($Entity.Name -match '^smime\.p7s$') { return $false }
        return $true
    }
    if ($t -eq 'text/plain' -or $t -eq 'text/html') { return $false }
    if ($embedded) { return $false }
    if (@('application/pkcs7-signature', 'application/x-pkcs7-signature', 'application/pgp-signature',
          'message/delivery-status', 'text/rfc822-headers') -contains $t) { return $false }
    return $true
}

function ConvertFrom-Eml {
    # Raw e-mail (bytes of an .eml file / IMAP message) -> readable fields
    param([byte[]]$Bytes)
    $root = Read-MimeEntity $Latin1.GetString($Bytes)
    $h = $root.Headers
    $from = Split-MailAddress (ConvertFrom-RawHeader $h['from'])
    $replyTo = if ($h['reply-to']) { (Split-MailAddress (ConvertFrom-RawHeader $h['reply-to'])).Address } else { '' }

    $leaves = New-Object Collections.Generic.List[object]
    Add-MimeLeaves $root $leaves
    $plain = $null; $html = $null
    $atts = New-Object Collections.Generic.List[object]
    $usedNames = @{}
    $extensions = @{ 'application/pdf' = '.pdf'; 'image/jpeg' = '.jpg'; 'image/png' = '.png'; 'image/gif' = '.gif'
                     'text/calendar' = '.ics'; 'text/plain' = '.txt'; 'text/html' = '.html'; 'message/rfc822' = '.eml' }
    foreach ($leaf in $leaves) {
        if (Test-MimeAttachment $leaf) {
            $name = $leaf.Name
            if (-not $name -and $leaf.Type -eq 'message/rfc822') {
                $inner = Get-MimeHeaders (Split-MimeHeaderBody $leaf.Body)[0]
                $name = (Get-SafeFileName (Limit-Text (ConvertFrom-RawHeader $inner['subject']) 60 '')).Trim()
                if (-not $name) { $name = 'attached_message' }
                $name += '.eml'
            }
            if (-not $name) {
                $ext = if ($extensions.ContainsKey($leaf.Type)) { $extensions[$leaf.Type] } else { '.bin' }
                $name = 'attachment' + $ext
            }
            $key = $name.ToLowerInvariant()
            if ($usedNames.ContainsKey($key)) {
                $usedNames[$key]++
                $name = [IO.Path]::GetFileNameWithoutExtension($name) + " ($($usedNames[$key]))" + [IO.Path]::GetExtension($name)
            } else { $usedNames[$key] = 1 }
            $bytes = Get-PartBytes $leaf
            $atts.Add([pscustomobject]@{ Name = $name; KB = [math]::Round($bytes.Length / 1KB); Bytes = $bytes })
        } elseif ($leaf.Type -eq 'text/plain' -and $null -eq $plain) {
            $plain = Get-PartText $leaf
        } elseif ($leaf.Type -eq 'text/html' -and $null -eq $html) {
            $html = Get-PartText $leaf
        }
    }
    $body = if ($null -ne $plain -and "$plain".Trim()) { $plain } elseif ($null -ne $html) { ConvertFrom-Html $html } else { "$plain" }

    return [pscustomobject]@{
        Subject     = ConvertFrom-RawHeader $h['subject']
        FromName    = $from.Name
        FromAddress = $from.Address
        ReplyTo     = $replyTo
        To          = (Split-MailAddress (ConvertFrom-RawHeader $h['to'])).Address
        MessageId   = "$($h['message-id'])".Trim()
        References  = "$($h['references'])".Trim()
        Date        = ConvertFrom-MailDate $h['date']
        Body        = $body.Trim()
        Attachments = $atts.ToArray()
        Headers     = $h
    }
}

function Get-EmlHeaderDate {
    # Only the Date header - used to skip messages of other days quickly
    param([byte[]]$Bytes)
    $n = [Math]::Min($Bytes.Length, 65536)
    $head = (Split-MimeHeaderBody $Latin1.GetString($Bytes, 0, $n))[0]
    return ConvertFrom-MailDate (Get-MimeHeaders $head)['date']
}

# ---------------------------------------------------------------------
# Writing replies as e-mail files (used for IMAP drafts and .eml drafts)
# ---------------------------------------------------------------------
function ConvertTo-EncodedWord {
    param([string]$Text)
    if ($Text -notmatch '[^\x20-\x7e]') { return $Text }
    $words = @()
    $i = 0
    while ($i -lt $Text.Length) {
        $len = [Math]::Min(10, $Text.Length - $i)
        if ($i + $len -lt $Text.Length -and [char]::IsHighSurrogate($Text[$i + $len - 1])) { $len++ }
        $words += '=?utf-8?B?' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text.Substring($i, $len))) + '?='
        $i += $len
    }
    return ($words -join "`r`n ")
}

function Format-MailDate {
    param([datetime]$Date)
    $off = [TimeZoneInfo]::Local.GetUtcOffset($Date)
    $sign = if ($off -lt [TimeSpan]::Zero) { '-' } else { '+' }
    $zone = '{0}{1:00}{2:00}' -f $sign, [Math]::Abs($off.Hours), [Math]::Abs($off.Minutes)
    return $Date.ToString('ddd, dd MMM yyyy HH:mm:ss ', [Globalization.CultureInfo]::InvariantCulture) + $zone
}

function Get-ReplySubject {
    param([string]$Subject)
    if ($Subject -match '^\s*(re|aw|sv|antw)\s*:') { return $Subject.Trim() }
    return ('Re: ' + $Subject.Trim()).Trim()
}

function Get-QuotedText {
    param($Mail)
    $when = if ($Mail.Date) { $Mail.Date.ToString('dd.MM.yyyy HH:mm') } else { '' }
    $head = "On $when, $($Mail.FromName) <$($Mail.FromAddress)> wrote:"
    $body = Limit-Text "$($Mail.Body)" ([int]$Config.MaxBodyCharacters) "`r`n[...]"
    $quoted = (($body -split '\r?\n') | ForEach-Object { '> ' + $_ }) -join "`r`n"
    return "$head`r`n$quoted"
}

function New-ReplyMessage {
    # Complete reply e-mail (plain text, UTF-8) as bytes
    param([string]$From, [string]$To, [string]$Subject, [string]$Text,
          [string]$InReplyTo = '', [string]$References = '', [datetime]$Date = (Get-Date), [switch]$Unsent)
    $domain = if ($From -match '@([^>\s]+)') { $Matches[1] } else { 'localhost' }
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add('Date: ' + (Format-MailDate $Date))
    $lines.Add("From: $From")
    $lines.Add("To: $To")
    $lines.Add('Subject: ' + (ConvertTo-EncodedWord (Get-ReplySubject $Subject)))
    $lines.Add("Message-ID: <$([guid]::NewGuid().ToString('N'))@$domain>")
    if ($InReplyTo) {
        $refs = @("$References" -split '\s+' | Where-Object { $_ }) + $InReplyTo
        $refs = @($refs | Select-Object -Last 10)
        $lines.Add("In-Reply-To: $InReplyTo")
        $lines.Add('References: ' + ($refs -join "`r`n "))
    }
    $lines.Add('MIME-Version: 1.0')
    $lines.Add('Content-Type: text/plain; charset=utf-8')
    $lines.Add('Content-Transfer-Encoding: base64')
    if ($Unsent) { $lines.Add('X-Unsent: 1') }          # opens as a new message in Outlook / Windows Mail
    $lines.Add('')
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($Text -replace '\r?\n', "`r`n")))
    for ($i = 0; $i -lt $b64.Length; $i += 76) { $lines.Add($b64.Substring($i, [Math]::Min(76, $b64.Length - $i))) }
    return , [Text.Encoding]::ASCII.GetBytes(($lines -join "`r`n") + "`r`n")
}

# ---------------------------------------------------------------------
# Mail sources
# ---------------------------------------------------------------------
function Get-SourceType { param($Source) return "$($Source.Type)".Trim().ToLowerInvariant() }

function Get-SourceLabel {
    param($Source)
    if ($Source.Name) { return "$($Source.Name)" }
    switch (Get-SourceType $Source) {
        'imap'    { return "$($Source.User)" }
        'outlook' { return 'Outlook' }
        default   { if ($Source.Path) { return (Split-Path -Leaf "$($Source.Path)") } else { return 'Mail' } }
    }
}

function New-MailRecord {
    param($Mail, [string]$Mailbox, [datetime]$Received, [bool]$Unread, [string]$Source, [byte[]]$Raw)
    return [pscustomobject]@{
        Nr          = 0
        Mailbox     = $Mailbox
        FromName    = $Mail.FromName
        FromAddress = $Mail.FromAddress
        Subject     = $Mail.Subject
        Received    = $Received
        Unread      = $Unread
        Attachments = @($Mail.Attachments | ForEach-Object { [pscustomobject]@{ Name = $_.Name; KB = $_.KB } })
        Body        = Limit-Text "$($Mail.Body)" ([int]$Config.MaxBodyCharacters) "`r`n[... shortened ...]"
        EntryID     = ''
        StoreID     = ''
        Source      = $Source
        SourceIndex = 0
        Raw         = $Raw
    }
}

# --- folder of .eml files --------------------------------------------
function Get-FolderDayEmails {
    param($Source, [datetime]$Day)
    $path = "$($Source.Path)"
    if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Container)) { throw "Folder not found: $path" }
    $label = Get-SourceLabel $Source
    Write-Host "  Reading $label (folder)..." -ForegroundColor Cyan
    $files = @(Get-ChildItem -LiteralPath $path -Filter '*.eml' -File -Recurse:([bool]$Source.Recurse) | Where-Object { $_.Extension -eq '.eml' })
    $list = New-Object Collections.Generic.List[object]
    foreach ($f in $files) {
        try {
            $bytes = [IO.File]::ReadAllBytes($f.FullName)
            $d = Get-EmlHeaderDate $bytes
            if (-not $d) { $d = $f.LastWriteTime }
            if ($d.Date -ne $Day.Date) { continue }
            $mail = ConvertFrom-Eml $bytes
            if (Test-IgnoredSubject $mail.Subject) { continue }
            $list.Add((New-MailRecord $mail $label $d $false 'folder' $bytes))
        } catch {
            Write-Host "    Skipped $($f.Name): $($_.Exception.Message)" -ForegroundColor DarkGray
        }
    }
    return $list.ToArray()
}

# --- Thunderbird (mbox) ------------------------------------------------
function Read-MboxDay {
    # Streams an mbox file and returns the raw messages dated on $Day.
    # Only the messages of that day are kept in memory.
    param([string]$Path, [datetime]$Day)
    $results = New-Object Collections.Generic.List[object]
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $rd = New-Object IO.StreamReader($fs, $Latin1)
    $lines = $null; $inHead = $false; $keep = $false; $status = 0; $date = $null; $prevBlank = $true
    try {
        while ($true) {
            $line = $rd.ReadLine()
            # a new message starts with a "From " line (after a blank line, or in the usual
            # "From - Tue Oct 13 09:12:00 2026" form)
            $isStart = ($null -ne $line) -and $line.StartsWith('From ') -and
                ($prevBlank -or $line -match '^From \S+ +[A-Z][a-z]{2} [A-Z][a-z]{2} +\d{1,2} \d{1,2}:\d{2}(:\d{2})? \d{4}')
            if ($null -eq $line -or $isStart) {
                if ($null -ne $lines -and $keep) {
                    $results.Add([pscustomobject]@{
                        Raw = $Latin1.GetBytes(($lines -join "`r`n").TrimEnd() + "`r`n")
                        Unread = -not ($status -band 1)
                        Date = $date
                    })
                }
                if ($null -eq $line) { break }
                $lines = New-Object Collections.Generic.List[string]
                $inHead = $true; $keep = $false; $status = 0; $date = $null; $prevBlank = $false
                continue
            }
            $prevBlank = ($line.Length -eq 0)
            if ($null -eq $lines) { continue }
            if ($inHead) {
                $lines.Add($line)
                if ($line.Length -eq 0) {
                    $inHead = $false
                    $h = Get-MimeHeaders ($lines -join "`r`n")
                    if ($h['x-mozilla-status']) { try { $status = [Convert]::ToInt32($h['x-mozilla-status'].Trim(), 16) } catch { } }
                    $date = ConvertFrom-MailDate $h['date']
                    $keep = ($null -ne $date) -and ($date.Date -eq $Day.Date) -and -not ($status -band 8)   # 8 = deleted
                }
                continue
            }
            if ($keep) {
                if ($line -match '^>+From ') { $line = $line.Substring(1) }
                $lines.Add($line)
            }
        }
    } finally {
        $rd.Dispose()
    }
    return $results.ToArray()
}

function Get-MboxDayEmails {
    param($Source, [datetime]$Day)
    $path = "$($Source.Path)"
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { throw "Mail file or folder not found: $path" }
    $label = Get-SourceLabel $Source
    Write-Host "  Reading $label (Thunderbird)..." -ForegroundColor Cyan
    if (Test-Path -LiteralPath $path -PathType Container) {
        # Thunderbird folders are files without an extension (Inbox, Sent...)
        $files = @(Get-ChildItem -LiteralPath $path -File | Where-Object { $_.Extension -eq '' })
    } else {
        $files = @(Get-Item -LiteralPath $path)
    }
    $list = New-Object Collections.Generic.List[object]
    foreach ($f in $files) {
        foreach ($m in (Read-MboxDay $f.FullName $Day)) {
            try {
                $mail = ConvertFrom-Eml $m.Raw
                if (Test-IgnoredSubject $mail.Subject) { continue }
                $box = if ($files.Count -gt 1) { "$label / $($f.Name)" } else { $label }
                $list.Add((New-MailRecord $mail $box $m.Date ([bool]$m.Unread) 'mbox' $m.Raw))
            } catch {
                Write-Host "    Skipped one message: $($_.Exception.Message)" -ForegroundColor DarkGray
            }
        }
    }
    return $list.ToArray()
}

# --- IMAP (Gmail, Yahoo, company servers) ------------------------------
# A small IMAP client over TLS. The mailbox is opened read-only (EXAMINE)
# and messages are fetched with BODY.PEEK, so nothing is marked as read.
function New-ImapSession {
    param([IO.Stream]$InStream, [IO.Stream]$OutStream)
    return @{ In = $InStream; Out = $OutStream; Buf = (New-Object byte[] 65536); Pos = 0; Len = 0; Tag = 0; Tcp = $null }
}

function Read-ImapLine {
    param($Session)
    $ms = New-Object IO.MemoryStream
    while ($true) {
        if ($Session.Pos -ge $Session.Len) {
            $Session.Len = $Session.In.Read($Session.Buf, 0, $Session.Buf.Length)
            $Session.Pos = 0
            if ($Session.Len -le 0) { throw 'The mail server closed the connection.' }
        }
        $i = [Array]::IndexOf($Session.Buf, [byte]10, $Session.Pos, $Session.Len - $Session.Pos)
        if ($i -lt 0) {
            $ms.Write($Session.Buf, $Session.Pos, $Session.Len - $Session.Pos)
            $Session.Pos = $Session.Len
            continue
        }
        $ms.Write($Session.Buf, $Session.Pos, $i - $Session.Pos)
        $Session.Pos = $i + 1
        break
    }
    return ($Latin1.GetString($ms.ToArray()) -replace '\r$', '')
}

function Read-ImapBytes {
    param($Session, [int]$Count)
    $out = New-Object byte[] $Count
    $got = 0
    while ($got -lt $Count) {
        if ($Session.Pos -lt $Session.Len) {
            $k = [Math]::Min($Count - $got, $Session.Len - $Session.Pos)
            [Array]::Copy($Session.Buf, $Session.Pos, $out, $got, $k)
            $Session.Pos += $k
        } else {
            $k = $Session.In.Read($out, $got, $Count - $got)
            if ($k -le 0) { throw 'The mail server closed the connection.' }
        }
        $got += $k
    }
    return , $out
}

function Read-ImapResponse {
    # One response; literals {n} are read as bytes and marked {LIT0}, {LIT1}...
    param($Session)
    $text = New-Object Text.StringBuilder
    $lits = New-Object Collections.Generic.List[object]
    while ($true) {
        $line = Read-ImapLine $Session
        $m = [regex]::Match($line, '\{(\d+)\+?\}$')
        if ($m.Success) {
            [void]$text.Append($line.Substring(0, $m.Index)).Append("{LIT$($lits.Count)}")
            $lits.Add((Read-ImapBytes $Session ([int]$m.Groups[1].Value)))
            continue
        }
        [void]$text.Append($line)
        break
    }
    return [pscustomobject]@{ Text = $text.ToString(); Literals = $lits.ToArray() }
}

function Send-ImapData {
    param($Session, [byte[]]$Bytes)
    $Session.Out.Write($Bytes, 0, $Bytes.Length)
    $Session.Out.Flush()
}

function Invoke-ImapCommand {
    param($Session, [string]$Command, [byte[]]$Literal = $null)
    $Session.Tag++
    $tag = 'A{0:D4}' -f $Session.Tag
    if ($null -ne $Literal) {
        Send-ImapData $Session $Latin1.GetBytes("$tag $Command {$($Literal.Length)}`r`n")
        $r = Read-ImapResponse $Session
        if (-not $r.Text.StartsWith('+')) {
            return [pscustomobject]@{ Status = 'NO'; StatusText = $r.Text; Responses = @() }
        }
        Send-ImapData $Session $Literal
        Send-ImapData $Session $Latin1.GetBytes("`r`n")
    } else {
        Send-ImapData $Session $Latin1.GetBytes("$tag $Command`r`n")
    }
    $untagged = New-Object Collections.Generic.List[object]
    while ($true) {
        $r = Read-ImapResponse $Session
        if ($r.Text.StartsWith("$tag ")) {
            $st = $r.Text.Substring($tag.Length + 1)
            return [pscustomobject]@{ Status = ($st -split ' ', 2)[0].ToUpperInvariant(); StatusText = $st; Responses = $untagged.ToArray() }
        }
        $untagged.Add($r)
    }
}

function Format-ImapString {
    param([string]$Text)
    return '"' + ($Text -replace '(["\\])', '\$1') + '"'
}

function ConvertTo-ImapUtf7 {
    # Folder names with non-English letters use "modified UTF-7" (RFC 3501)
    param([string]$Name)
    $sb = New-Object Text.StringBuilder
    $i = 0
    while ($i -lt $Name.Length) {
        $c = $Name[$i]
        if ([int]$c -ge 0x20 -and [int]$c -le 0x7e) {
            if ($c -eq '&') { [void]$sb.Append('&-') } else { [void]$sb.Append($c) }
            $i++
            continue
        }
        $j = $i
        while ($j -lt $Name.Length -and ([int]$Name[$j] -lt 0x20 -or [int]$Name[$j] -gt 0x7e)) { $j++ }
        $b64 = [Convert]::ToBase64String([Text.Encoding]::BigEndianUnicode.GetBytes($Name.Substring($i, $j - $i)))
        [void]$sb.Append('&').Append(($b64.TrimEnd('=') -replace '/', ',')).Append('-')
        $i = $j
    }
    return $sb.ToString()
}

function ConvertFrom-ImapDate {
    # INTERNALDATE '13-Oct-2026 09:12:00 +0300' -> local [datetime]
    param([string]$Text)
    $s = ($Text.Trim() -replace '([+-]\d{2})(\d{2})$', '$1:$2')
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParseExact($s, 'd-MMM-yyyy HH:mm:ss zzz', [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dto)) { return $dto.LocalDateTime }
    return $null
}

function Get-ImapFetchInfo {
    # Parses '* 5 FETCH (UID 12 FLAGS (\Seen) INTERNALDATE "..." RFC822.SIZE 345)'
    param([string]$Text)
    if ($Text -notmatch '^\* \d+ FETCH \(') { return $null }
    $uid = if ($Text -match '\bUID (\d+)') { [int64]$Matches[1] } else { $null }
    $flags = if ($Text -match '\bFLAGS \(([^)]*)\)') { $Matches[1] } else { '' }
    $date = if ($Text -match '\bINTERNALDATE "([^"]+)"') { ConvertFrom-ImapDate $Matches[1] } else { $null }
    $size = if ($Text -match '\bRFC822\.SIZE (\d+)') { [int64]$Matches[1] } else { 0 }
    return [pscustomobject]@{ Uid = $uid; Seen = ($flags -match '\\Seen'); Date = $date; Size = $size }
}

function Connect-Imap {
    param($Source)
    $server = "$($Source.Host)"
    $port = if ($Source.Port) { [int]$Source.Port } else { 993 }
    if (-not $server -or -not $Source.User) { throw 'IMAP source needs "Host" and "User" in config.json (or run setup).' }
    $secret = Get-SourcePassword $Source
    $tcp = New-Object Net.Sockets.TcpClient
    $tcp.ReceiveTimeout = 60000
    $tcp.SendTimeout = 60000
    try { $tcp.Connect($server, $port) } catch { throw "Cannot reach ${server}:$port - check the server name and the internet connection." }
    $ssl = New-Object Net.Security.SslStream($tcp.GetStream(), $false)
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        $ssl.AuthenticateAsClient($server, $null, [Security.Authentication.SslProtocols]::Tls12, $true)
    } else {
        $ssl.AuthenticateAsClient($server)
    }
    $s = New-ImapSession $ssl $ssl
    $s.Tcp = $tcp
    $greeting = Read-ImapLine $s
    if ($greeting -notmatch '^\* (OK|PREAUTH)') { Disconnect-Imap $s; throw "Unexpected answer from ${server}: $greeting" }
    $r = Invoke-ImapCommand $s ('LOGIN ' + (Format-ImapString "$($Source.User)") + ' ' + (Format-ImapString $secret))
    if ($r.Status -ne 'OK') {
        Disconnect-Imap $s
        Clear-SavedPassword $Source
        throw "Login refused by $server ($($r.StatusText)). Check the address and the app password (it was not kept)."
    }
    return $s
}

function Disconnect-Imap {
    param($Session)
    try { [void](Invoke-ImapCommand $Session 'LOGOUT') } catch { }
    try { $Session.In.Dispose() } catch { }
    if ($Session.Tcp) { try { $Session.Tcp.Close() } catch { } }
}

function Open-ImapFolder {
    param($Session, [string]$Folder)
    $r = Invoke-ImapCommand $Session ('EXAMINE ' + (Format-ImapString (ConvertTo-ImapUtf7 $Folder)))
    if ($r.Status -ne 'OK') { throw "Folder '$Folder' cannot be opened: $($r.StatusText)" }
    $exists = 0
    foreach ($x in $r.Responses) { if ($x.Text -match '^\* (\d+) EXISTS') { $exists = [int]$Matches[1] } }
    return $exists
}

function Get-ImapDayEmails {
    param($Source, [datetime]$Day)
    $label = Get-SourceLabel $Source
    $folder = if ($Source.Folder) { "$($Source.Folder)" } else { 'INBOX' }
    Write-Host "  Reading $label / $folder (IMAP)..." -ForegroundColor Cyan
    $list = New-Object Collections.Generic.List[object]
    $s = Connect-Imap $Source
    try {
        [void](Open-ImapFolder $s $folder)
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $since = $Day.AddDays(-1).ToString('d-MMM-yyyy', $inv)
        $before = $Day.AddDays(2).ToString('d-MMM-yyyy', $inv)
        $r = Invoke-ImapCommand $s "UID SEARCH SINCE $since BEFORE $before"
        if ($r.Status -ne 'OK') { throw "Search failed: $($r.StatusText)" }
        $uids = @()
        foreach ($x in $r.Responses) {
            if ($x.Text -match '^\* SEARCH(.*)$') { $uids += @($Matches[1].Trim() -split '\s+' | Where-Object { $_ }) }
        }
        if ($uids.Count -eq 0) { return $list.ToArray() }

        $r = Invoke-ImapCommand $s ('UID FETCH ' + ($uids -join ',') + ' (UID FLAGS INTERNALDATE RFC822.SIZE)')
        $wanted = @($r.Responses | ForEach-Object { Get-ImapFetchInfo $_.Text } |
                    Where-Object { $_ -and $_.Uid -and $_.Date -and $_.Date.Date -eq $Day.Date })
        $maxBytes = [int64]([double]$Config.MaxMessageMB * 1MB)
        foreach ($w in $wanted) {
            $part = if ($w.Size -gt $maxBytes) { 'BODY.PEEK[HEADER]' } else { 'BODY.PEEK[]' }
            $r = Invoke-ImapCommand $s "UID FETCH $($w.Uid) ($part)"
            $resp = $r.Responses | Where-Object { $_.Literals.Count -gt 0 } | Select-Object -First 1
            if (-not $resp) { continue }
            $raw = [byte[]]$resp.Literals[0]
            try {
                $mail = ConvertFrom-Eml $raw
            } catch {
                Write-Host "    Skipped one message: $($_.Exception.Message)" -ForegroundColor DarkGray
                continue
            }
            if (Test-IgnoredSubject $mail.Subject) { continue }
            if ($part -like '*HEADER*') { $mail.Body = "[Message too large ($([math]::Round($w.Size / 1MB)) MB) - open it in the webmail.]" }
            $list.Add((New-MailRecord $mail $label $w.Date (-not $w.Seen) 'imap' $raw))
        }
    } finally {
        Disconnect-Imap $s
    }
    return $list.ToArray()
}

function Find-ImapDraftsFolder {
    param($Session, $Source)
    if ($Source.DraftsFolder) { return (ConvertTo-ImapUtf7 "$($Source.DraftsFolder)") }
    $r = Invoke-ImapCommand $Session 'LIST "" "*"'
    $names = @()
    foreach ($x in $r.Responses) {
        if ($x.Text -notmatch '^\* LIST \(([^)]*)\) (?:NIL|"(?:[^"\\]|\\.)*") (.+)$') { continue }
        $flags = $Matches[1]
        $n = $Matches[2].Trim()
        if ($n -match '^\{LIT(\d+)\}$') { $n = $Latin1.GetString($x.Literals[[int]$Matches[1]]) }
        elseif ($n.StartsWith('"')) { $n = $n.Substring(1, $n.Length - 2) -replace '\\(.)', '$1' }
        if ($flags -match '\\Drafts') { return $n }
        $names += $n
    }
    foreach ($n in $names) { if ($n -match '^(INBOX[./])?Drafts?$') { return $n } }
    return 'Drafts'
}

function Save-ImapDraft {
    # Puts the reply into the Drafts folder of the mailbox (APPEND)
    param($Source, [byte[]]$Message)
    $s = Connect-Imap $Source
    try {
        $folder = Find-ImapDraftsFolder $s $Source
        $r = Invoke-ImapCommand $s ('APPEND ' + (Format-ImapString $folder) + ' (\Draft \Seen)') $Message
        if ($r.Status -ne 'OK') { throw "The server did not accept the draft: $($r.StatusText)" }
        return $folder
    } finally {
        Disconnect-Imap $s
    }
}

# --- passwords (Windows DPAPI: readable only by this Windows user on this PC)
$script:PasswordCache = @{}

function Get-CredentialPath {
    param($Source)
    $user = if ($Source.User) { $Source.User } else { $Source.SmtpUser }
    $server = if ($Source.Host) { $Source.Host } else { $Source.SmtpHost }
    return (Join-Path $CredentialFolder ((Get-SafeFileName "$($user)_$server") + '.cred'))
}

function Clear-SavedPassword {
    param($Source)
    $f = Get-CredentialPath $Source
    $script:PasswordCache.Remove($f)
    if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
}

function Save-SourcePassword {
    param($Source, [Security.SecureString]$Secret)
    if (-not (Test-Path $CredentialFolder)) { New-Item -ItemType Directory -Path $CredentialFolder | Out-Null }
    ConvertFrom-SecureString $Secret | Set-Content -LiteralPath (Get-CredentialPath $Source) -Encoding ASCII
}

function Get-SourcePassword {
    param($Source)
    $fromEnv = [Environment]::GetEnvironmentVariable('MAILASSISTANT_PASSWORD')
    if ($fromEnv) { return $fromEnv }
    $f = Get-CredentialPath $Source
    if ($script:PasswordCache.ContainsKey($f)) { return $script:PasswordCache[$f] }
    if (Test-Path -LiteralPath $f) {
        try {
            $sec = (Get-Content -LiteralPath $f -Raw).Trim() | ConvertTo-SecureString
            $plain = (New-Object Net.NetworkCredential('', $sec)).Password
            $script:PasswordCache[$f] = $plain
            return $plain
        } catch {
            Write-Host '  The saved password cannot be read for this Windows user. Please enter it again.' -ForegroundColor Yellow
        }
    }
    $sec = Read-Host -AsSecureString "  App password for $(Get-SourceLabel $Source)"
    $plain = (New-Object Net.NetworkCredential('', $sec)).Password
    if (-not $plain) { throw 'No password entered.' }
    if ("$(Read-Host '  Save it (encrypted, only for your Windows user)? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') {
        Save-SourcePassword $Source $sec
    }
    $script:PasswordCache[$f] = $plain
    return $plain
}

# --- sending through SMTP ----------------------------------------------
function Send-SmtpReply {
    param($Source, [string]$From, [string]$To, [string]$Subject, [string]$Body, [string]$InReplyTo = '', [string]$References = '')
    $port = if ($Source.SmtpPort) { [int]$Source.SmtpPort } else { 587 }
    if ($port -eq 465) { throw 'SMTP port 465 is not supported here. Use port 587 (Gmail and Yahoo support it).' }
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    $user = if ($Source.SmtpUser) { "$($Source.SmtpUser)" } else { "$($Source.User)" }
    $secret = Get-SourcePassword $Source
    $msg = New-Object Net.Mail.MailMessage
    $smtp = New-Object Net.Mail.SmtpClient("$($Source.SmtpHost)", $port)
    try {
        $msg.From = New-Object Net.Mail.MailAddress($From)
        $msg.To.Add($To)
        $msg.Subject = Get-ReplySubject $Subject
        $msg.Body = $Body
        $msg.SubjectEncoding = [Text.Encoding]::UTF8
        $msg.BodyEncoding = [Text.Encoding]::UTF8
        if ($InReplyTo) {
            $msg.Headers.Add('In-Reply-To', $InReplyTo)
            $msg.Headers.Add('References', ("$References $InReplyTo").Trim())
        }
        $smtp.EnableSsl = $true
        $smtp.Timeout = 60000
        $smtp.Credentials = New-Object Net.NetworkCredential($user, $secret)
        $smtp.Send($msg)
    } finally {
        $msg.Dispose()
        $smtp.Dispose()
    }
}

# ---------------------------------------------------------------------
# Outlook helpers
# ---------------------------------------------------------------------
$script:OutlookApp = $null
$script:OutlookNs = $null

function Test-OutlookProfile {
    # True when Outlook has at least one mail profile. Starting Outlook without
    # one opens the "New Profile" window, which the tool must never cause.
    foreach ($key in @('HKCU:\Software\Microsoft\Office\*\Outlook\Profiles\*',
                       'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\*')) {
        try { if (@(Get-Item -Path $key -ErrorAction SilentlyContinue).Count -gt 0) { return $true } } catch { }
    }
    return $false
}

function Get-Outlook {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        if (-not (Test-OutlookProfile)) {
            throw ('Outlook has no e-mail account on this PC, so it was not started. ' +
                   'Use option 6 (setup): remove Outlook from the list and add Gmail / Yahoo / IMAP or a mail folder.')
        }
        try {
            return New-Object -ComObject Outlook.Application
        } catch {
            throw ('Outlook could not be started. Open Outlook first and make sure a mailbox is ' +
                   'set up and fully loaded (no "Add Account" window), then try again. ' +
                   'If Outlook runs "as administrator", run this tool the same way. ' +
                   'Without Outlook, add a Gmail / Yahoo / IMAP mailbox or a mail folder with option 6 (setup). ' +
                   "Details: $($_.Exception.Message)")
        }
    }
}

function Get-OutlookNamespace {
    if (-not $script:OutlookNs) {
        $script:OutlookApp = Get-Outlook
        $script:OutlookNs = $script:OutlookApp.GetNamespace('MAPI')
    }
    return $script:OutlookNs
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
# Reading the e-mails of one day (all sources)
# ---------------------------------------------------------------------
function Get-OutlookDayEmails {
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
                            Source      = 'outlook'
                            SourceIndex = 0
                            Raw         = $null
                        })
                    }
                }
            } catch {}
            $it = $items.GetNext()
        }
    }
    if ($skipped -gt 0) { Write-Host "  Skipped $skipped automatic delivery receipts." -ForegroundColor DarkGray }
    return $list.ToArray()
}

function Get-DayEmails {
    param([datetime]$Day)
    $sources = @($Config.Sources)
    if ($sources.Count -eq 0) { throw 'No mailbox is set up. Run option 6 (setup) first.' }
    $list = New-Object Collections.Generic.List[object]
    for ($i = 0; $i -lt $sources.Count; $i++) {
        $src = $sources[$i]
        $found = $null
        try {
            switch (Get-SourceType $src) {
                'outlook' { $found = Get-OutlookDayEmails (Get-OutlookNamespace) $Day }
                'imap'    { $found = Get-ImapDayEmails $src $Day }
                'folder'  { $found = Get-FolderDayEmails $src $Day }
                'mbox'    { $found = Get-MboxDayEmails $src $Day }
                default   { throw "Unknown source type '$($src.Type)' (use outlook, imap, folder or mbox)." }
            }
            foreach ($x in @($found | Where-Object { $null -ne $_ })) { $x.SourceIndex = $i; $list.Add($x) }
        } catch {
            if ($sources.Count -eq 1) { throw }
            Write-Host "  Could not read $(Get-SourceLabel $src): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
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
    $cache = Get-CachePath $Day
    if (Test-Path $cache) { Get-ChildItem -LiteralPath $cache -Filter '*.eml' | Remove-Item -Force }
    $Emails | ForEach-Object {
        $file = ''
        if ($null -ne $_.Raw) {
            if (-not (Test-Path $cache)) { New-Item -ItemType Directory -Path $cache | Out-Null }
            $file = Join-Path $cache ("{0}.eml" -f $_.Nr)
            [IO.File]::WriteAllBytes($file, [byte[]]$_.Raw)
        }
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
            Source      = $_.Source
            SourceIndex = $_.SourceIndex
            Cache       = $file
        }
    } | Export-Csv -Path (Get-IndexPath $Day) -Delimiter ';' -NoTypeInformation -Encoding UTF8
}

function Get-Index {
    # Reads the day's list; creates it first if it does not exist
    param([datetime]$Day)
    $f = Get-IndexPath $Day
    if (-not (Test-Path $f)) {
        Write-Host 'No list for this day yet. Creating it now...' -ForegroundColor Cyan
        Save-Index (Get-DayEmails $Day) $Day
    }
    return @(Import-Csv -Path $f -Delimiter ';' -Encoding UTF8)
}

function Get-RowSource {
    param($Row)
    $s = "$($Row.Source)".Trim().ToLowerInvariant()
    if ($s) { return $s } else { return 'outlook' }      # lists made by version 1
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
    $emails = Get-DayEmails $day
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
        if ("$(Read-Host 'Print attachments now? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') { Invoke-Print -Day $day }
    }
}

# =====================================================================
# 2. PRINT ATTACHMENTS (asks for each one)
# =====================================================================
function Get-EmailAttachments {
    # Attachments of one listed e-mail, ready to be saved. $null = e-mail not found.
    param($Row)
    $list = New-Object Collections.Generic.List[object]
    if ((Get-RowSource $Row) -eq 'outlook') {
        try { $it = (Get-OutlookNamespace).GetItemFromID($Row.EntryID, $Row.StoreID) } catch { return $null }
        for ($k = 1; $k -le $it.Attachments.Count; $k++) {
            $a = $it.Attachments.Item($k)
            $list.Add([pscustomobject]@{ Name = $a.FileName; Com = $a; Bytes = $null })
        }
    } else {
        if (-not $Row.Cache -or -not (Test-Path -LiteralPath $Row.Cache)) { return $null }
        $mail = ConvertFrom-Eml ([IO.File]::ReadAllBytes($Row.Cache))
        foreach ($a in $mail.Attachments) { $list.Add([pscustomobject]@{ Name = $a.Name; Com = $null; Bytes = $a.Bytes }) }
    }
    return , $list.ToArray()
}

function Save-Attachment {
    param($Attachment, [string]$Path)
    if ($null -ne $Attachment.Com) { $Attachment.Com.SaveAsFile($Path) }
    else { [IO.File]::WriteAllBytes($Path, [byte[]]$Attachment.Bytes) }
}

function Invoke-Print {
    param([Nullable[datetime]]$Day = $null)
    if (-not $Day) { $Day = Read-Day }
    $index = Get-Index $Day
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
        $all = Get-EmailAttachments $e
        if ($null -eq $all) {
            Write-Host '  E-mail not found any more (moved or deleted?).' -ForegroundColor Yellow
            continue
        }
        $skipEmail = $false
        foreach ($attName in ("$($e.Attachments)" -split ' \| ')) {
            if ($skipEmail) { break }
            $att = $all | Where-Object { $_.Name -eq $attName } | Select-Object -First 1
            if (-not $att) { continue }

            $path = Join-Path $attFolder ("{0}_{1}" -f $e.Nr, (Get-SafeFileName $attName))
            $ext = [IO.Path]::GetExtension($attName).ToLowerInvariant()

            while ($true) {
                $r = "$(Read-Host ("    {0}  -> print? (P/N/V/S/Q)" -f $attName))".Trim().ToUpperInvariant()
                if ($r -eq 'Q') { Write-Host "Stopped. Printed: $printed." -ForegroundColor Cyan; return }
                if ($r -eq 'S') { $skipEmail = $true; break }
                if ($r -eq 'N' -or $r -eq '') { break }
                if (-not (Test-Path $path)) { Save-Attachment $att $path }
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
function Read-ReplyText {
    # Template number or own text. Returns $null when cancelled.
    Write-Host 'Which reply?' -ForegroundColor Cyan
    foreach ($k in $Config.Templates.Keys) { Write-Host ("  {0} = {1}" -f $k, $Config.Templates[$k]) }
    Write-Host '  W = write my own text'
    Write-Host '  C = cancel'
    $opt = "$(Read-Host 'Choose')".Trim().ToUpperInvariant()
    if ($opt -eq 'C' -or $opt -eq '') { return $null }
    if ($opt -eq 'W') {
        Write-Host 'Type the text. Press Enter on an empty line to finish.' -ForegroundColor Gray
        $lines = @()
        while ($true) {
            $l = Read-Host
            if ($l -eq '') { break }
            $lines += $l
        }
        if ($lines.Count -eq 0) { return $null }
        return ($lines -join "`r`n")
    }
    if ($Config.Templates.Contains($opt)) { return $Config.Templates[$opt] }
    Write-Host 'Unknown option.' -ForegroundColor Yellow
    return $null
}

function Write-ReplyLog {
    param([string]$Log, [string]$Mailbox, [string]$To, [string]$Subject, [string]$Text, [string]$Status)
    $color = if ($Status -like 'error*') { 'Red' } else { 'Green' }
    Write-Host "  -> $Status" -ForegroundColor $color
    [pscustomobject]@{
        Date    = Get-Date -Format 'dd.MM.yyyy HH:mm'
        Mailbox = $Mailbox
        To      = $To
        Subject = $Subject
        Text    = ($Text -replace "[\r\n]+", ' ')
        Status  = $Status
    } | Export-Csv -Path $Log -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Append
}

function Invoke-OutlookReply {
    param($Row, [string]$Log)
    try { $it = (Get-OutlookNamespace).GetItemFromID($Row.EntryID, $Row.StoreID) } catch {
        Write-Host 'E-mail not found any more (moved or deleted?).' -ForegroundColor Yellow
        return
    }
    Write-Host ''
    Write-Host "From:    $($Row.FromName) <$($Row.FromAddress)>" -ForegroundColor White
    Write-Host "Subject: $($Row.Subject)" -ForegroundColor White
    Write-Host "Mailbox: $($Row.Mailbox)" -ForegroundColor Gray
    Write-Host '----------------------------------------------------------'
    Write-Host (Limit-Text "$($it.Body)".Trim() 800)
    Write-Host '----------------------------------------------------------'

    $text = Read-ReplyText
    if ($null -eq $text) { return }

    $reply = $it.Reply()
    Set-ReplySender $script:OutlookApp $reply $Row.Mailbox
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
        return
    }
    Write-ReplyLog $Log $Row.Mailbox $Row.FromAddress $Row.Subject $text $status
}

function Invoke-MimeReply {
    # Replies for IMAP / .eml folder / Thunderbird sources
    param($Row, [string]$Log)
    $sources = @($Config.Sources)
    $idx = [int]"0$($Row.SourceIndex)"
    $src = if ($idx -lt $sources.Count) { $sources[$idx] } else { $null }
    if (-not $Row.Cache -or -not (Test-Path -LiteralPath $Row.Cache)) {
        Write-Host 'The copy of this e-mail is missing. Run the report for this day again.' -ForegroundColor Yellow
        return
    }
    $mail = ConvertFrom-Eml ([IO.File]::ReadAllBytes($Row.Cache))
    Write-Host ''
    Write-Host "From:    $($mail.FromName) <$($mail.FromAddress)>" -ForegroundColor White
    Write-Host "Subject: $($mail.Subject)" -ForegroundColor White
    Write-Host "Mailbox: $($Row.Mailbox)" -ForegroundColor Gray
    Write-Host '----------------------------------------------------------'
    Write-Host (Limit-Text "$($mail.Body)".Trim() 800)
    Write-Host '----------------------------------------------------------'

    $text = Read-ReplyText
    if ($null -eq $text) { return }

    $from = "$($src.Address)"
    if (-not $from -and "$($src.User)" -like '*@*') { $from = "$($src.User)" }
    if ($from -notlike '*@*') {
        # suggest the address the e-mail was sent to; Enter accepts it
        $from = Read-Value 'Your e-mail address (the reply is sent from it)' $mail.To
        if ($from -notlike '*@*') { Write-Host 'No valid address - cancelled.' -ForegroundColor Yellow; return }
        if ($src -and -not $src.Address) { $src | Add-Member -NotePropertyName Address -NotePropertyValue $from -Force }
    }
    $to = if ($mail.ReplyTo) { $mail.ReplyTo } else { $mail.FromAddress }
    if (-not $to) { Write-Host 'The e-mail has no sender address - cannot reply.' -ForegroundColor Yellow; return }

    $fullText = New-ReplyText $text
    $body = $fullText + "`r`n`r`n" + (Get-QuotedText $mail)
    $type = Get-SourceType $src
    $canSend = [bool]$src.SmtpHost

    Write-Host ''
    Write-Host "To:      $to" -ForegroundColor White
    Write-Host "Subject: $(Get-ReplySubject $mail.Subject)" -ForegroundColor White
    Write-Host $fullText -ForegroundColor Gray
    $where = if ($type -eq 'imap') { 'in the Drafts folder of the mailbox' } else { 'as an .eml file you open in your mail program' }
    $prompt = "D = save as draft ($where)" + $(if ($canSend) { ', S = send now' } else { '' }) + ', C = cancel'
    $fin = "$(Read-Host $prompt)".Trim().ToUpperInvariant()

    $status = $null
    try {
        if ($fin -eq 'S' -and $canSend) {
            if ("$(Read-Host 'Type YES to send')".Trim().ToUpperInvariant() -eq 'YES') {
                Send-SmtpReply $src $from $to $mail.Subject $body $mail.MessageId $mail.References
                $status = 'sent'
            } else { $fin = 'D' }
        }
        if ($fin -eq 'D') {
            if ($type -eq 'imap') {
                $eml = New-ReplyMessage -From $from -To $to -Subject $mail.Subject -Text $body -InReplyTo $mail.MessageId -References $mail.References
                $folder = Save-ImapDraft $src $eml
                $status = 'draft'
                Write-Host "  Saved in '$folder'. Open the mailbox (webmail or app) to check and send it." -ForegroundColor Gray
            } else {
                $eml = New-ReplyMessage -From $from -To $to -Subject $mail.Subject -Text $body -InReplyTo $mail.MessageId -References $mail.References -Unsent
                $dir = Join-Path $OutputFolder 'replies'
                if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
                $file = Join-Path $dir ("reply_{0}_{1}.eml" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $Row.Nr)
                [IO.File]::WriteAllBytes($file, [byte[]]$eml)
                $status = 'draft'
                Write-Host "  Saved: $file" -ForegroundColor Gray
                if ("$(Read-Host 'Open it in your mail program now? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') { Open-File $file }
            }
        }
    } catch {
        $status = "error: $($_.Exception.Message)"
    }
    if (-not $status) { Write-Host 'Cancelled.' -ForegroundColor Gray; return }
    Write-ReplyLog $Log $Row.Mailbox $to $mail.Subject $text $status
}

function Invoke-Reply {
    $day = Read-Day
    $index = Get-Index $day
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

        if ((Get-RowSource $e) -eq 'outlook') { Invoke-OutlookReply $e $log }
        else { Invoke-MimeReply $e $log }
    }
}

# =====================================================================
# 6. SETUP - add / remove mailboxes (writes config.json)
# =====================================================================
function Find-ThunderbirdFolders {
    $found = @()
    if (-not $env:APPDATA) { return $found }
    $profiles = Join-Path $env:APPDATA 'Thunderbird\Profiles'
    if (-not (Test-Path $profiles)) { return $found }
    foreach ($p in Get-ChildItem -LiteralPath $profiles -Directory) {
        foreach ($kind in @('Mail', 'ImapMail')) {
            $root = Join-Path $p.FullName $kind
            if (-not (Test-Path $root)) { continue }
            foreach ($acc in Get-ChildItem -LiteralPath $root -Directory) {
                $inbox = Get-ChildItem -LiteralPath $acc.FullName -File | Where-Object { $_.Name -ieq 'Inbox' } | Select-Object -First 1
                if ($inbox) { $found += $inbox.FullName }
            }
        }
    }
    return $found
}

function Read-Value {
    param([string]$Prompt, [string]$Default = '')
    $p = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
    $v = "$(Read-Host $p)".Trim()
    if (-not $v) { return $Default }
    return $v
}

function Test-ImapSource {
    param($Source)
    $s = Connect-Imap $Source
    try {
        $folder = if ($Source.Folder) { "$($Source.Folder)" } else { 'INBOX' }
        return (Open-ImapFolder $s $folder)
    } finally {
        Disconnect-Imap $s
    }
}

function Invoke-Setup {
    $cfg = if (Test-Path $ConfigPath) { Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json } else { [pscustomobject]@{} }
    $sources = New-Object Collections.Generic.List[object]
    if ($cfg.PSObject.Properties['Sources']) { foreach ($x in @($cfg.Sources)) { if ($x) { $sources.Add($x) } } }
    else { $sources.Add([pscustomobject]@{ Type = 'outlook' }) }
    $implicitOutlook = -not $cfg.PSObject.Properties['Sources']

    while ($true) {
        Write-Host ''
        Write-Host 'Mailboxes read by the assistant:' -ForegroundColor Cyan
        if ($sources.Count -eq 0) { Write-Host '  (none)' }
        for ($i = 0; $i -lt $sources.Count; $i++) {
            $x = $sources[$i]
            $detail = switch (Get-SourceType $x) {
                'imap'    { "IMAP $($x.Host)" }
                'outlook' { 'all mailboxes in the Outlook profile' }
                default   { "$($x.Path)" }
            }
            Write-Host ("  {0}. {1} - {2}" -f ($i + 1), (Get-SourceLabel $x), $detail)
        }
        Write-Host ''
        Write-Host 'Add:  1 = Outlook (desktop)   2 = Gmail   3 = Yahoo Mail   4 = other IMAP server'
        Write-Host '      5 = folder with .eml files   6 = Thunderbird'
        Write-Host 'R = remove a mailbox   P = change a saved password   Q = save and exit   X = exit without saving'
        $opt = "$(Read-Host 'Choose')".Trim().ToUpperInvariant()

        if ($opt -eq 'X') { return }
        if ($opt -eq 'Q' -or $opt -eq '') { break }
        if ($opt -eq 'R') {
            $n = [int]"0$(Read-Host 'Number to remove')"
            if ($n -ge 1 -and $n -le $sources.Count) { $sources.RemoveAt($n - 1) }
            continue
        }
        if ($opt -eq 'P') {
            $n = [int]"0$(Read-Host 'Number of the mailbox')"
            if ($n -ge 1 -and $n -le $sources.Count -and ($sources[$n - 1].Host -or $sources[$n - 1].SmtpHost)) {
                Clear-SavedPassword $sources[$n - 1]
                Save-SourcePassword $sources[$n - 1] (Read-Host -AsSecureString 'New app password')
                Write-Host '  Saved.' -ForegroundColor Green
            } else { Write-Host '  That mailbox has no password.' -ForegroundColor Yellow }
            continue
        }

        $new = $null
        switch ($opt) {
            '1' {
                if (@($sources | Where-Object { (Get-SourceType $_) -eq 'outlook' }).Count -gt 0) { Write-Host '  Outlook is already in the list.' -ForegroundColor Yellow; break }
                $new = [pscustomobject]@{ Type = 'outlook' }
            }
            { $_ -in '2', '3' } {
                $addr = Read-Value 'E-mail address'
                if ($addr -notlike '*@*') { Write-Host '  Not an e-mail address.' -ForegroundColor Yellow; break }
                if ($opt -eq '2') {
                    Write-Host '  Gmail needs an APP PASSWORD (Google account > Security > 2-Step Verification > App passwords).' -ForegroundColor Gray
                    $new = [pscustomobject]@{ Type = 'imap'; Name = "Gmail - $addr"; Host = 'imap.gmail.com'; Port = 993; User = $addr
                                              Address = $addr; SmtpHost = 'smtp.gmail.com'; SmtpPort = 587 }
                } else {
                    Write-Host '  Yahoo needs an APP PASSWORD (Account info > Account security > Generate app password).' -ForegroundColor Gray
                    $new = [pscustomobject]@{ Type = 'imap'; Name = "Yahoo - $addr"; Host = 'imap.mail.yahoo.com'; Port = 993; User = $addr
                                              Address = $addr; SmtpHost = 'smtp.mail.yahoo.com'; SmtpPort = 587 }
                }
            }
            '4' {
                $server = Read-Value 'IMAP server (e.g. mail.company.com)'
                if (-not $server) { break }
                $user = Read-Value 'User name (usually the e-mail address)'
                $addr = Read-Value 'E-mail address' $(if ($user -like '*@*') { $user } else { '' })
                $new = [pscustomobject]@{ Type = 'imap'; Name = (Read-Value 'Name shown in the report' $addr); Host = $server
                                          Port = [int](Read-Value 'IMAP port (SSL)' '993'); User = $user; Address = $addr }
                $smtpHost = Read-Value 'SMTP server for sending (Enter = no sending, drafts only)'
                if ($smtpHost) {
                    $new | Add-Member SmtpHost $smtpHost
                    $new | Add-Member SmtpPort ([int](Read-Value 'SMTP port (STARTTLS)' '587'))
                }
            }
            '5' {
                $path = Read-Value 'Folder with the .eml files'
                if (-not (Test-Path -LiteralPath $path -PathType Container)) { Write-Host '  Folder not found.' -ForegroundColor Yellow; break }
                $new = [pscustomobject]@{ Type = 'folder'; Name = (Read-Value 'Name shown in the report' (Split-Path -Leaf $path)); Path = $path
                                          Recurse = ("$(Read-Host 'Include subfolders? (Y/N)')".Trim().ToUpperInvariant() -eq 'Y') }
            }
            '6' {
                $found = @(Find-ThunderbirdFolders)
                for ($i = 0; $i -lt $found.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i + 1), $found[$i]) }
                $pick = Read-Value 'Number from the list, or the full path of a Thunderbird mail file'
                $path = if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $found.Count) { $found[[int]$pick - 1] } else { $pick }
                if (-not $path -or -not (Test-Path -LiteralPath $path)) { Write-Host '  Not found.' -ForegroundColor Yellow; break }
                $acc = Split-Path -Leaf (Split-Path -Parent $path)
                $new = [pscustomobject]@{ Type = 'mbox'; Name = (Read-Value 'Name shown in the report' "Thunderbird - $acc"); Path = $path }
            }
            default { Write-Host '  Unknown option.' -ForegroundColor Yellow }
        }
        if (-not $new) { continue }

        if ((Get-SourceType $new) -eq 'imap') {
            Clear-SavedPassword $new
            $sec = Read-Host -AsSecureString '  App password'
            Save-SourcePassword $new $sec
            try {
                $n = Test-ImapSource $new
                Write-Host "  Connected. $n e-mails in the inbox." -ForegroundColor Green
            } catch {
                Write-Host "  Connection test failed: $($_.Exception.Message)" -ForegroundColor Yellow
                if ("$(Read-Host '  Keep this mailbox anyway? (Y/N)')".Trim().ToUpperInvariant() -ne 'Y') { Clear-SavedPassword $new; continue }
                Save-SourcePassword $new $sec
            }
        }
        if ($implicitOutlook -and (Get-SourceType $new) -ne 'outlook' -and $sources.Count -eq 1 -and (Get-SourceType $sources[0]) -eq 'outlook') {
            $sources.RemoveAt(0)
            Write-Host '  Outlook was taken off the list (add it back with 1 if this PC uses Outlook).' -ForegroundColor Gray
        }
        $implicitOutlook = $false
        $sources.Add($new)
    }

    $cfg | Add-Member -NotePropertyName Sources -NotePropertyValue @($sources.ToArray()) -Force
    [IO.File]::WriteAllText($ConfigPath, ($cfg | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
    $Config.Sources = @($sources.ToArray())
    Write-Host "Saved in $ConfigPath" -ForegroundColor Green
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
        @('Gmail - registry@example.com', 'Ana Popescu', 'ana.p@example.org', 'Scanned documents', 13, 5, $true,
          @(@{ Name = 'scan_page1.jpg'; KB = 912 }, @{ Name = 'scan_page2.jpg'; KB = 876 }), "Hello,`r`n`r`nAttached are the two pages you asked for.`r`n`r`nAna"),
        @('Gmail - registry@example.com', 'Peter Novak', 'p.novak@example.com', 'Forwarded: confirmation', 15, 31, $false,
          @(@{ Name = 'confirmation.eml'; KB = 64 }), "Hi,`r`n`r`nForwarding the confirmation e-mail below.`r`n`r`nPeter")
    )
    $list = foreach ($r in $raw) {
        [pscustomobject]@{
            Nr = 0; Mailbox = $r[0]; FromName = $r[1]; FromAddress = $r[2]; Subject = $r[3]
            Received = $Day.Date.AddHours($r[4]).AddMinutes($r[5]); Unread = $r[6]
            Attachments = @($r[7] | ForEach-Object { [pscustomobject]$_ }); Body = $r[8]
            EntryID = ''; StoreID = ''; Source = 'demo'; SourceIndex = 0; Raw = $null
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
    try {
        Import-AssistantConfig $ConfigPath
        Initialize-Folders
        switch ($Command) {
            'report' { Invoke-Report }
            'print'  { Invoke-Print }
            'reply'  { Invoke-Reply }
            'setup'  { Invoke-Setup }
            'demo'   { Invoke-Demo }
        }
    } catch {
        Write-Host ''
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
