```markdown
# Email Management Scripts

This directory contains PowerShell scripts designed to manage and automate email tasks in Microsoft 365.

## Forward Emails Script

The `ForwardEmails.ps1` script forwards emails from a specified Microsoft 365 mailbox to another mailbox based on a date range. It is designed for forwarding emails to external (third-party) mailboxes.

### Description

The script uses Exchange Online PowerShell to forward emails from a source mailbox to a target mailbox within a specified date range. It utilizes the Compliance Search feature for email discovery and forwarding.

### Parameters

- **SourceMailbox**: The email address of the source mailbox.
- **TargetMailbox**: The email address of the target mailbox where emails will be forwarded.
- **StartDate**: The start date for the email search range.
- **EndDate**: The end date for the email search range. Defaults to the current date if not specified.
- **LogFilePath**: The path where the log file will be created. Defaults to "EmailForwardLog_[timestamp].txt" in the current directory.
- **Verbose**: If specified, provides more detailed logging information.

### Usage

```powershell
.\ForwardEmails.ps1 -SourceMailbox "user@domain.com" -TargetMailbox "external@example.com" -StartDate "06/01/2024" -EndDate "06/30/2024" -Verbose
```

### Example

```powershell
.\ForwardEmails.ps1 -SourceMailbox "user@domain.com" -TargetMailbox "external@example.com" -StartDate "06/01/2024" -EndDate "06/30/2024" -Verbose
```

### Notes

This script requires the ExchangeOnlineManagement module and appropriate permissions to perform compliance searches.
```