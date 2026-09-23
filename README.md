# Grant-MailboxDelegation

PowerShell script to grant **Full Access** and **Send As** permissions on a list of Exchange Online (Microsoft 365) mailboxes to a single delegate user.

## Features

- Grants `FullAccess` (read and manage) and `SendAs` on each mailbox in a list
- Idempotent: existing permissions are detected and skipped, so the script can be run several times safely
- Checks that the delegate and each mailbox exist before applying changes
- Supports `-WhatIf` to simulate without applying changes
- Optional control of Outlook AutoMapping
- Optional CSV report export
- Reuses an existing Exchange Online session if one is already open

## Requirements

- PowerShell 5.1 or PowerShell 7+
- [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) module v3 or later
- An account with the **Exchange Administrator** role (or equivalent)

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Input file

A plain text file with one mailbox UPN or email address per line. Empty lines, duplicates and lines starting with `#` are ignored.

```text
# Shared mailboxes
support@contoso.com
sales@contoso.com

# User mailboxes
jane.smith@contoso.com
```

## Usage

Basic usage:

```powershell
.\Grant-MailboxDelegation.ps1 -Delegate "john.doe@contoso.com" -MailboxListPath ".\mailboxes.txt"
```

Disable AutoMapping and export a report:

```powershell
.\Grant-MailboxDelegation.ps1 -Delegate "john.doe@contoso.com" -MailboxListPath "C:\temp\mailboxes.txt" -AutoMapping $false -ReportPath "C:\temp\report.csv"
```

Simulate without applying any change:

```powershell
.\Grant-MailboxDelegation.ps1 -Delegate "john.doe@contoso.com" -MailboxListPath ".\mailboxes.txt" -WhatIf
```

Built-in help:

```powershell
Get-Help .\Grant-MailboxDelegation.ps1 -Full
```

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Delegate` | Yes | UPN of the user who will **receive** the permissions |
| `-MailboxListPath` | Yes | Path to the text file listing the mailboxes to delegate |
| `-AutoMapping` | No | `$true` (default) auto-adds the mailboxes to the delegate's Outlook profile; `$false` disables it |
| `-ReportPath` | No | Path of a CSV file to export the results |
| `-WhatIf` | No | Shows what would be done without applying changes |

## Output

A summary table is displayed at the end of the run. Each mailbox gets one of these statuses for `FullAccess` and `SendAs`:

| Status | Meaning |
|---|---|
| `Granted` | Permission successfully added |
| `Already present` | Permission already existed, nothing changed |
| `Failed` | An error occurred (details in the `Error` column of the report) |
| `Skipped` | Mailbox not found |
| `WhatIf` | Simulation mode, nothing applied |

The CSV report uses `;` as delimiter so that it opens directly in Excel with European regional settings.

## Notes

- Permissions can take **up to 60 minutes** to become effective. Restarting Outlook may be required.
- `FullAccess` alone does **not** allow sending email; this is why the script also grants `SendAs`.
- If a user has both `SendAs` and `SendOnBehalf`, `SendAs` takes precedence.
- By default, emails sent as a shared mailbox are saved in the delegate's Sent Items. To keep a copy in the shared mailbox's Sent Items:

  ```powershell
  Set-Mailbox -Identity "support@contoso.com" -MessageCopyForSentAsEnabled $true
  ```

- With many mailboxes, consider `-AutoMapping $false` to avoid slowing down the delegate's Outlook.

## License

MIT (or replace with the license of your choice)
