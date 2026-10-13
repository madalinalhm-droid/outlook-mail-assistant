# Outlook Mail Assistant

A menu-driven **PowerShell** tool for desk work with a busy shared mailbox: it builds a
daily report of every e-mail received in all mailboxes of the local Outlook profile,
prints attachments one by one after asking, and sends replies from ready-made templates.

It was built for **locked-down office PCs**: it needs only what Windows and Office already
have – no installs, no admin rights, no Python, no internet API, no mail-server credentials.
It talks to the Outlook that is already open on the machine.

> Built from a real-world back-office workflow and published here in a generic form.
> The screenshot and demo use fictitious e-mails only.

![Demo report with fictitious e-mails](docs/demo_report.png)

## What it does

| Menu | What happens |
|---|---|
| **1. Report** | Reads the Inbox of every mailbox in the Outlook profile for a chosen day and creates an **HTML report** (table of contents, sender, subject, body, attachments, unread flag) plus an **Excel file** with filters. Delivery receipts are skipped. Nothing is changed in Outlook – e-mails are not even marked as read. |
| **2. Print attachments** | Walks through the day's attachments and asks for each one: **P**rint / **N**o / **V**iew / **S**kip e-mail / **Q**uit. Attached e-mails (`.msg`) are printed through Outlook in a background job with a timeout, so a hanging Outlook cannot freeze the tool. Signature images are ignored. |
| **3. Reply** | Lists the day's e-mails, shows the chosen one, and builds a reply from a numbered **template** or your own text, with greeting and signature. The reply is sent **from the mailbox the e-mail arrived in**, is saved as a **draft** by default, and is only sent after typing `YES`. Every reply is logged to CSV and Excel. |
| **5. Demo** | Generates the HTML report from fictitious e-mails – no Outlook needed. |

## Safety by design

- Read-only towards the mailbox: no deleting, moving or flagging.
- Nothing is sent without an explicit confirmation; drafts are the default.
- Values written to Excel are protected against formula injection.
- Personal settings (`config.json`) and all generated reports are excluded by `.gitignore`.

## Getting started

1. Download the repository (green **Code** button → **Download ZIP**) and unzip it.
2. Optional: copy `config.example.json` to `config.json` and set your signature and reply templates.
3. Double-click **`Start.bat`** and choose an option.

To try it without Outlook, choose **5 (Demo)**, or run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File OutlookMailAssistant.ps1 demo
```

Reports are written to the `Reports` folder next to the script.

### Requirements

- Windows with Windows PowerShell 5.1 (built in) and Microsoft Outlook (desktop).
- Microsoft Excel – optional, only for the `.xlsx` files.

## Configuration (`config.json`)

```json
{
  "Signature": "Your Name / Department",
  "Greeting": "Hello,",
  "Closing": "Best regards,",
  "Templates": {
    "1": "Please send the missing document as soon as possible.",
    "2": "We have received your documents. Thank you."
  },
  "IgnoredSubjects": ["Delivery Status Notification \\((Relay|Delay|Success)\\)"],
  "MaxBodyCharacters": 5000
}
```

## Tech

PowerShell 5.1/7 · Outlook & Excel COM automation (MAPI stores, `PropertyAccessor`,
`SendUsingAccount`) · background jobs with timeout · HTML/CSS report generation ·
CSV/Excel logging · dependency-free test script · GitHub Actions (tests + PSScriptAnalyzer).

## Tests

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Run-Tests.ps1
```

The tests cover date parsing, subject filtering, file-name sanitising, Excel injection
protection, configuration loading and the demo report. Outlook itself is not needed.

## License

MIT
