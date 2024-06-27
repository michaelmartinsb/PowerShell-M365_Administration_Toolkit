```markdown
# Email Management Scripts

This directory contains PowerShell scripts designed to manage and automate email tasks in Microsoft 365.

## Forward Emails Script

The `ForwardEmails.ps1` script forwards emails from a specified Microsoft 365 mailbox to another mailbox based on a date range. It is designed for forwarding emails to external (third-party) mailboxes.

### Description

The script uses Exchange Online PowerShell to forward emails from a source mailbox to a target mailbox within a specified date range. It utilizes the `Search-Mailbox` cmdlet to ensure each email is forwarded individually. The script includes checks for necessary permissions and will attempt to use `Connect-IPPSSession` if required.

### Parameters

- **SourceMailbox**: The email address of the source mailbox.
- **TargetMailbox**: The email address of the target mailbox where emails will be forwarded.
- **StartDate**: The start date for the email search range.
- **EndDate**: The end date for the email search range. Defaults to the current date if not specified.
- **LogFilePath**: The path where the log file will be created. Defaults to "EmailForwardLog_[timestamp].txt" in the current directory.
- **TestMode**: If specified, the script will run in test mode without actually forwarding any emails.
- **Verbose**: If specified, provides more detailed logging information.

### Usage

.\ForwardEmails.ps1 -SourceMailbox "user@domain.com" -TargetMailbox "external@example.com" -StartDate "06/01/2024" -EndDate "06/30/2024" -TestMode -Verbose
```

### Example

```powershell
.\ForwardEmails.ps1 -SourceMailbox "user@domain.com" -TargetMailbox "external@example.com" -StartDate "06/01/2024" -EndDate "06/30/2024" -TestMode -Verbose
```

### Notes

This script requires the ExchangeOnlineManagement module and appropriate permissions to use the `Search-Mailbox` cmdlet. It may attempt to use `Connect-IPPSSession` if necessary permissions are not immediately available.

You may need to run `Connect-IPPSSession` before running this script to access compliance cmdlets. If the compliance cmdlets are not available, you will be prompted to run the following command:

```powershell
Connect-IPPSSession -UserPrincipalName your_admin@yourdomain.com
```

Ensure you have the necessary roles assigned to your account to use the `Search-Mailbox` cmdlet.
