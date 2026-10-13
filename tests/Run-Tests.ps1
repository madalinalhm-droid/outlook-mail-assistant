# Dependency-free tests for the pure functions, the mail-file readers and the demo report.
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
function Get-Utf8B64 { param([string]$Text) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)) }
function ConvertTo-Bytes { param([string]$Text) , [Text.Encoding]::GetEncoding(28591).GetBytes(($Text -replace '\r?\n', "`r`n")) }

# Non-ASCII test strings built from code points (keeps this file pure ASCII)
$a_ = [string][char]0x103    # a with breve
$s_ = [string][char]0x219    # s with comma
$t_ = [string][char]0x21B    # t with comma
$roName = "Mar${a_}rescu ${s_}tefan"
$roSubject = "Cerere adeverin${t_}${a_}"

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
'{"Signature":"Test Dept","Templates":{"9":"Custom text"},"_help":"x","Sources":[{"Type":"folder","Path":"C:\\Mail"},{"Type":"imap","Host":"imap.example.com"}]}' |
    Set-Content -Path $cfgPath -Encoding UTF8
Import-AssistantConfig $cfgPath
Assert 'signature overridden'      ($Config.Signature -eq 'Test Dept')
Assert 'templates replaced'        ($Config.Templates.Contains('9') -and -not $Config.Templates.Contains('1'))
Assert '_help keys skipped'        (-not $Config.Contains('_help'))
Assert 'sources loaded'            (@($Config.Sources).Count -eq 2 -and (Get-SourceType $Config.Sources[1]) -eq 'imap')
$Config.Signature = 'Your Name'

Write-Host 'Header decoding'
Assert 'encoded word (B)'          ((ConvertFrom-EncodedWord ("=?utf-8?B?" + (Get-Utf8B64 $roName) + "?=")) -eq $roName)
Assert 'encoded word (Q) + joined' ((ConvertFrom-EncodedWord '=?iso-8859-1?Q?Caf=E9_au?= =?iso-8859-1?Q?_lait?=') -eq ("Caf" + [char]0xE9 + " au lait"))
Assert 'plain text untouched'      ((ConvertFrom-EncodedWord 'Hello world') -eq 'Hello world')
$raw8 = [Text.Encoding]::GetEncoding(28591).GetString([Text.Encoding]::UTF8.GetBytes($roSubject))
Assert 'raw UTF-8 header'          ((ConvertFrom-RawHeader $raw8) -eq $roSubject)
$p = Get-HeaderParams "attachment; filename*=utf-8''cerere%C8%99.pdf"
Assert 'RFC 2231 file name'        ($p.Value -eq 'attachment' -and $p.Params['filename'] -eq "cerere$s_.pdf")
$p = Get-HeaderParams 'text/plain; charset="windows-1250"; format=flowed'
Assert 'quoted parameter'          ($p.Value -eq 'text/plain' -and $p.Params['charset'] -eq 'windows-1250' -and $p.Params['format'] -eq 'flowed')
$p = Get-HeaderParams 'application/pdf; name*0="very_long_"; name*1="name.pdf"'
Assert 'parameter continuation'    ($p.Params['name'] -eq 'very_long_name.pdf')
$x = Split-MailAddress '"Smith, John" <john@example.com>'
Assert 'address with quoted name'  ($x.Name -eq 'Smith, John' -and $x.Address -eq 'john@example.com')
$x = Split-MailAddress 'ana@example.org (Ana P)'
Assert 'address with comment name' ($x.Name -eq 'Ana P' -and $x.Address -eq 'ana@example.org')
$x = Split-MailAddress 'bob@example.net'
Assert 'bare address'              ($x.Name -eq 'bob@example.net' -and $x.Address -eq 'bob@example.net')

Write-Host 'Dates'
$exp = ([DateTimeOffset]::new(2026, 10, 13, 9, 12, 0, [TimeSpan]::FromHours(3))).LocalDateTime
Assert 'RFC 5322 date'             ((ConvertFrom-MailDate 'Tue, 13 Oct 2026 09:12:00 +0300 (EEST)') -eq $exp)
$exp = ([DateTimeOffset]::new(2026, 10, 13, 9, 12, 0, [TimeSpan]::Zero)).LocalDateTime
Assert 'GMT, no day name'          ((ConvertFrom-MailDate '13 Oct 2026 09:12:00 GMT') -eq $exp)
Assert 'no seconds'                ((ConvertFrom-MailDate 'Tue, 13 Oct 2026 09:12 +0000') -eq $exp)
Assert 'bad date -> null'          ($null -eq (ConvertFrom-MailDate 'yesterday'))
Assert 'IMAP internal date'        ((ConvertFrom-ImapDate '13-Oct-2026 09:12:00 +0000') -eq $exp)
Assert 'IMAP date, short day'      ($null -ne (ConvertFrom-ImapDate ' 3-Oct-2026 09:12:00 +0000'))

Write-Host 'E-mail files (MIME)'
$pdf = [byte[]](37, 80, 68, 70, 45, 49, 46, 52, 10, 1, 2, 3, 250, 251, 252)
$body = "Bun${a_} ziua,`r`nv${a_} rog s${a_} g${a_}si${t_}i cererea.`r`nMaria"
$qp = ([Text.Encoding]::UTF8.GetBytes($body) | ForEach-Object {
    if ($_ -eq 13 -or $_ -eq 10 -or ($_ -ge 32 -and $_ -lt 127 -and $_ -ne 61)) { [string][char]$_ } else { '={0:X2}' -f $_ } }) -join ''
$eml1 = @"
Return-Path: <maria@example.org>
From: =?utf-8?B?$(Get-Utf8B64 $roName)?= <maria@example.org>
To: office@example.com
Subject: =?utf-8?Q?Cerere_adeverin=C8=9B?=
 =?utf-8?Q?=C4=83?=
Date: Tue, 13 Oct 2026 12:00:00 +0000
Message-ID: <abc123@example.org>
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="==outer=="

This is a multi-part message in MIME format.
--==outer==
Content-Type: multipart/related; boundary="inner"

--inner
Content-Type: multipart/alternative; boundary=alt

--alt
Content-Type: text/plain; charset=utf-8
Content-Transfer-Encoding: quoted-printable

$qp
--alt
Content-Type: text/html; charset=utf-8

<p>HTML version</p>
--alt--
--inner
Content-Type: image/png; name="image001.png"
Content-Disposition: inline; filename="image001.png"
Content-ID: <image001@x>
Content-Transfer-Encoding: base64

iVBORw0KGgo=
--inner--
--==outer==
Content-Type: application/pdf; name="cerere.pdf"
Content-Disposition: attachment; filename*=utf-8''cerere%C8%99.pdf
Content-Transfer-Encoding: base64

$([Convert]::ToBase64String($pdf))
--==outer==
Content-Type: application/pdf
Content-Disposition: attachment; filename="scan.pdf"
Content-Transfer-Encoding: base64

JVBERg==
--==outer==
Content-Type: application/pdf
Content-Disposition: attachment; filename="scan.pdf"
Content-Transfer-Encoding: base64

JVBERg==
--==outer==
Content-Type: message/rfc822

From: someone@example.net
Subject: Original confirmation

Inner text
--==outer==--
"@
$m = ConvertFrom-Eml (ConvertTo-Bytes $eml1)
Assert 'sender name decoded'       ($m.FromName -eq $roName -and $m.FromAddress -eq 'maria@example.org')
Assert 'folded subject decoded'    ($m.Subject -eq $roSubject)
Assert 'plain text preferred, QP'  ($m.Body -eq $body)
Assert 'message id'                ($m.MessageId -eq '<abc123@example.org>')
$names = @($m.Attachments | ForEach-Object { $_.Name })
Assert 'signature image skipped'   ($names -notcontains 'image001.png')
Assert 'attachment names'          (($names -join '|') -eq "cerere$s_.pdf|scan.pdf|scan (2).pdf|Original confirmation.eml")
Assert 'attachment bytes exact'    ([Convert]::ToBase64String($m.Attachments[0].Bytes) -eq [Convert]::ToBase64String($pdf))
Assert 'attached e-mail kept'      ([Text.Encoding]::ASCII.GetString($m.Attachments[3].Bytes) -match 'Inner text')

$w1250 = Get-TextEncoding 'windows-1250'
$tc = [string][char]0x163   # t with cedilla (the Windows-1250 letter)
$htmlBytes = $w1250.GetBytes("<html><head><style>p{color:red}</style></head><body><p>Bun${a_} ziua,</p><p>Mul${tc}umesc&nbsp;&amp; o zi bun${a_}!<br>Ion</p></body></html>")
$eml2 = "From: Ion <ion@example.ro>`r`nSubject: HTML only`r`nDate: 13 Oct 2026 12:00:00 +0000`r`nContent-Type: text/html; charset=windows-1250`r`nContent-Transfer-Encoding: 8bit`r`n`r`n"
$m2 = ConvertFrom-Eml ([byte[]]((ConvertTo-Bytes $eml2) + $htmlBytes))
Assert 'HTML-only -> text'         ($m2.Body -eq "Bun${a_} ziua,`r`nMul${tc}umesc & o zi bun${a_}!`r`nIon")
Assert 'Windows-1250 decoded'      ($m2.Body -match "Mul${tc}umesc")

Write-Host 'Reply e-mail'
$r = New-ReplyMessage -From 'office@example.com' -To 'maria@example.org' -Subject $roSubject -Text "Thank you.`nLine 2" `
        -InReplyTo '<abc123@example.org>' -References '<first@x>' -Unsent
$rm = ConvertFrom-Eml $r
Assert 'reply subject Re:'         ($rm.Subject -eq "Re: $roSubject")
Assert 'reply body round trip'     ($rm.Body -eq "Thank you.`r`nLine 2")
Assert 'threading headers'         ($rm.Headers['in-reply-to'] -eq '<abc123@example.org>' -and $rm.Headers['references'] -match '<first@x>\s+<abc123@example.org>')
Assert 'draft flag for .eml'       ($rm.Headers['x-unsent'] -eq '1')
Assert 'reply is pure ASCII'       (-not ($r | Where-Object { $_ -gt 127 }))
Assert 'Re: not doubled'           ((Get-ReplySubject 'RE: Hello') -eq 'RE: Hello')
$q = Get-QuotedText $m
Assert 'original quoted'           ($q -match '^On 13\.10\.2026' -and $q -match "(?m)^> Bun$a_ ziua,")

Write-Host 'Folder of .eml files'
$Config.IgnoredSubjects = @('^\s*(Delivered|Relayed)\s*:')
$mailDir = Join-Path $tmp 'mail'
New-Item -ItemType Directory -Path $mailDir | Out-Null
[IO.File]::WriteAllBytes((Join-Path $mailDir 'a.eml'), (ConvertTo-Bytes $eml1))
[IO.File]::WriteAllBytes((Join-Path $mailDir 'b.eml'), (ConvertTo-Bytes ($eml2 + 'x').Replace('13 Oct', '12 Oct')))
[IO.File]::WriteAllBytes((Join-Path $mailDir 'c.eml'), (ConvertTo-Bytes "From: a@b.c`r`nSubject: Delivered: test`r`nDate: 13 Oct 2026 12:30:00 +0000`r`n`r`nx"))
$folderSrc = [pscustomobject]@{ Type = 'folder'; Name = 'Saved mail'; Path = $mailDir }
$f = @(Get-FolderDayEmails $folderSrc $today)
Assert 'only the chosen day, receipts skipped' ($f.Count -eq 1 -and $f[0].Subject -eq $roSubject)
Assert 'record fields'             ($f[0].Mailbox -eq 'Saved mail' -and $f[0].Attachments.Count -eq 4 -and $f[0].Source -eq 'folder')

Write-Host 'Thunderbird mbox'
$mbox = @"
From - Mon Oct 12 12:00:00 2026
X-Mozilla-Status: 0001
From: old@example.com
Subject: Yesterday
Date: Mon, 12 Oct 2026 12:00:00 +0000

old
From - Tue Oct 13 12:00:00 2026
X-Mozilla-Status: 0000
From: Ana <ana@example.org>
Subject: Today unread
Date: Tue, 13 Oct 2026 12:05:00 +0000

first line
>From the archive
end

From - Tue Oct 13 12:10:00 2026
X-Mozilla-Status: 0009
From: del@example.org
Subject: Deleted
Date: Tue, 13 Oct 2026 12:10:00 +0000

gone

From - Tue Oct 13 12:20:00 2026
X-Mozilla-Status: 0001
From: read@example.org
Subject: Today read
Date: Tue, 13 Oct 2026 12:20:00 +0000

read
"@
$mboxPath = Join-Path $tmp 'Inbox'
[IO.File]::WriteAllBytes($mboxPath, (ConvertTo-Bytes $mbox))
$mb = @(Get-MboxDayEmails ([pscustomobject]@{ Type = 'mbox'; Path = $mboxPath }) $today)
Assert 'mbox: day filter, deleted skipped' (($mb | ForEach-Object Subject) -join '|' -eq 'Today unread|Today read')
Assert 'mbox: read / unread flags'  ($mb[0].Unread -and -not $mb[1].Unread)
Assert 'mbox: >From unescaped'      ($mb[0].Body -eq "first line`r`nFrom the archive`r`nend")

Write-Host 'Index, cache and attachments'
Initialize-Folders
$list = Set-EmailNumbers (@($f) + @($mb))
Save-Index $list $today
$idx = @(Import-Csv -Path (Get-IndexPath $today) -Delimiter ';' -Encoding UTF8)
Assert 'index written'              ($idx.Count -eq 3 -and (Test-Path $idx[0].Cache))
$row = $idx | Where-Object { $_.Source -eq 'folder' }
$atts = Get-EmailAttachments $row
Assert 'attachments from the copy'  ($atts.Count -eq 4 -and [Convert]::ToBase64String($atts[0].Bytes) -eq [Convert]::ToBase64String($pdf))
Save-Attachment $atts[0] (Join-Path $tmp 'saved.pdf')
Assert 'attachment saved'           ((Get-Item (Join-Path $tmp 'saved.pdf')).Length -eq $pdf.Length)
Assert 'old lists = outlook'        ((Get-RowSource ([pscustomobject]@{ Source = '' })) -eq 'outlook')

Write-Host 'IMAP protocol'
$msg = ConvertTo-Bytes "Subject: Hi`r`n`r`nBody"
$serverText = "* 1 FETCH (UID 7 FLAGS (\Seen) BODY[] {$($msg.Length)}`r`n"
$server = [byte[]]([Text.Encoding]::ASCII.GetBytes($serverText) + $msg + [Text.Encoding]::ASCII.GetBytes(")`r`nA0001 OK FETCH completed`r`n"))
$out = New-Object IO.MemoryStream
$s = New-ImapSession (New-Object IO.MemoryStream(, $server)) $out
$res = Invoke-ImapCommand $s 'UID FETCH 7 (BODY.PEEK[])'
Assert 'command sent with tag'      ([Text.Encoding]::ASCII.GetString($out.ToArray()) -eq "A0001 UID FETCH 7 (BODY.PEEK[])`r`n")
Assert 'tagged status'              ($res.Status -eq 'OK')
Assert 'literal read exactly'       ($res.Responses.Count -eq 1 -and [Convert]::ToBase64String($res.Responses[0].Literals[0]) -eq [Convert]::ToBase64String($msg))
Assert 'literal marked in text'     ($res.Responses[0].Text -eq '* 1 FETCH (UID 7 FLAGS (\Seen) BODY[] {LIT0})')
$info = Get-ImapFetchInfo '* 3 FETCH (UID 45 FLAGS (\Answered) INTERNALDATE "13-Oct-2026 09:12:00 +0000" RFC822.SIZE 2048)'
Assert 'fetch info parsed'          ($info.Uid -eq 45 -and -not $info.Seen -and $info.Size -eq 2048 -and $info.Date -eq $exp)
Assert 'folder name UTF-7'          ((ConvertTo-ImapUtf7 ("Entw" + [char]0xFC + "rfe")) -eq 'Entw&APw-rfe' -and (ConvertTo-ImapUtf7 'A&B') -eq 'A&-B')
Assert 'quoted IMAP string'         ((Format-ImapString 'p"a\ss') -eq '"p\"a\\ss"')

Write-Host 'Demo report'
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
