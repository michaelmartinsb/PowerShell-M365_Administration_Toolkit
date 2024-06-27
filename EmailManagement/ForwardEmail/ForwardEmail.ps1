<#
.SYNOPSIS
    Forwards emails from a specified Microsoft 365 mailbox to another mailbox based on a date range.

.DESCRIPTION
    This script uses Exchange Online PowerShell to forward emails from a source mailbox to a target mailbox
    within a specified date range. It uses the Search-Mailbox cmdlet to ensure each email is forwarded individually.
    The script includes checks for necessary permissions and will attempt to use Connect-IPPSSession if required.

.PARAMETER SourceMailbox
    The email address of the source mailbox.

.PARAMETER TargetMailbox
    The email address of the target mailbox where emails will be forwarded.

.PARAMETER StartDate
    The start date for the email search range.

.PARAMETER EndDate
    The end date for the email search range. Defaults to the current date if not specified.

.PARAMETER LogFilePath
    The path where the log file will be created. Defaults to "EmailForwardLog_[timestamp].txt" in the current directory.

.PARAMETER TestMode
    If specified, the script will run in test mode without actually forwarding any emails.

.PARAMETER Verbose
    If specified, provides more detailed logging information.

.EXAMPLE
    .\ForwardEmails.ps1 -SourceMailbox "user@domain.com" -TargetMailbox "external@example.com" -StartDate "06/01/2024" -EndDate "06/30/2024" -TestMode -Verbose

.NOTES
    This script requires the ExchangeOnlineManagement module and appropriate permissions to use the Search-Mailbox cmdlet.
    It may attempt to use Connect-IPPSSession if necessary permissions are not immediately available.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SourceMailbox,
    
    [Parameter(Mandatory=$true)]
    [string]$TargetMailbox,
    
    [Parameter(Mandatory=$true)]
    [datetime]$StartDate,
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),  # Defaults to current date if not specified
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({Test-Path (Split-Path $_) -PathType 'Container'})]
    [string]$LogFilePath = "EmailForwardLog_$(Get-Date -Format 'yyyyMMddHHmmss').txt",

    [Parameter(Mandatory=$false)]
    [switch]$TestMode
)

# Function to write to log file
function Write-Log {
    param(
        [string]$Message,
        [switch]$Verbose
    )
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $LogFilePath -Value $LogMessage
    if ($Verbose) {
        Write-Verbose $LogMessage
    } else {
        Write-Host $LogMessage
    }
}

# Function to check environment
function Test-Environment {
    try {
        # Check if running with admin privileges
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            throw "This script requires administrator privileges."
        }

        # Check for Exchange Online PowerShell module
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "ExchangeOnlineManagement module not found. Do you want to install it? (Y/N)"
            $installModule = Read-Host
            if ($installModule -eq 'Y') {
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
                Write-Log "ExchangeOnlineManagement module installed."
            } else {
                throw "ExchangeOnlineManagement module is required. Exiting script."
            }
        }

        # Check if Search-Mailbox cmdlet is available
        if (!(Get-Command Search-Mailbox -ErrorAction SilentlyContinue)) {
            Write-Log "Search-Mailbox cmdlet is not available. Attempting to connect using Connect-IPPSSession."
            
            # Attempt to connect using Connect-IPPSSession
            $UserCredential = Get-Credential -Message "Enter your Exchange Online credentials for IPPSSession"
            Connect-IPPSSession -Credential $UserCredential
            
            # Check again for Search-Mailbox cmdlet
            if (!(Get-Command Search-Mailbox -ErrorAction SilentlyContinue)) {
                Write-Log "Search-Mailbox cmdlet is still not available after Connect-IPPSSession."
                Write-Log "Please ensure you have the necessary roles assigned to your account."
                throw "Search-Mailbox cmdlet not available."
            } else {
                Write-Log "Successfully connected using Connect-IPPSSession. Search-Mailbox cmdlet is now available."
            }
        }

        Write-Log "Environment check passed."
        return $true
    }
    catch {
        Write-Log "Environment check failed: $_"
        return $false
    }
}

# Function to connect to Exchange Online
function Connect-ToExchangeOnline {
    try {
        Import-Module ExchangeOnlineManagement
        $UserCredential = Get-Credential -Message "Enter your Exchange Online credentials"
        Connect-ExchangeOnline -Credential $UserCredential -ShowProgress $true
        Write-Log "Successfully connected to Exchange Online."
    }
    catch {
        Write-Log "Error connecting to Exchange Online: $_"
        throw
    }
}

# Function to forward emails individually
function Forward-Emails {
    param (
        [string]$SourceMailbox,
        [string]$TargetMailbox,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$TestMode
    )

    try {
        $searchQuery = "Received:>=$($StartDate.ToString('MM/dd/yyyy')) AND Received:<=$($EndDate.ToString('MM/dd/yyyy'))"

        Write-Log "Preparing to forward emails from $SourceMailbox to $TargetMailbox" -Verbose
        Write-Log "Date range: $StartDate to $EndDate" -Verbose

        if (-not $TestMode) {
            # Use Search-Mailbox to forward emails
            $result = Search-Mailbox -Identity $SourceMailbox -SearchQuery $searchQuery -TargetMailbox $TargetMailbox -TargetFolder "Inbox" -LogLevel Full -LogOnly:$false
            
            Write-Log "Email forwarding completed." -Verbose
            Write-Log "Number of messages forwarded: $($result.ResultItemsCount)" -Verbose
            Write-Log "Total size of forwarded messages: $($result.ResultItemsSize)" -Verbose
        } else {
            Write-Log "[TEST MODE] Would forward emails with the following parameters:" -Verbose
            Write-Log "[TEST MODE] Source Mailbox: $SourceMailbox" -Verbose
            Write-Log "[TEST MODE] Target Mailbox: $TargetMailbox" -Verbose
            Write-Log "[TEST MODE] Search Query: $searchQuery" -Verbose
        }
    }
    catch {
        Write-Log "Error in email forwarding process: $_"
        throw
    }
}

# Function to get user confirmation
function Get-UserConfirmation {
    param(
        [string]$SourceMailbox,
        [string]$TargetMailbox,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$TestMode
    )
    $modeString = if ($TestMode) { "TEST" } else { "LIVE" }
    $confirmMessage = "Are you sure you want to run the script in $modeString mode to forward emails from $SourceMailbox to $TargetMailbox for the period $StartDate to $EndDate?"
    do {
        $confirmation = Read-Host "$confirmMessage (Y/N)"
        if ($confirmation -eq 'Y') {
            return $true
        }
        elseif ($confirmation -eq 'N') {
            return $false
        }
        else {
            Write-Log "Invalid input: $confirmation. Please enter 'Y' or 'N'."
        }
    } while ($true)
}

# Main script execution
try {
    Write-Log "Script started in $(if ($TestMode) { 'TEST' } else { 'LIVE' }) mode"

    # Check environment
    if (-not (Test-Environment)) {
        throw "Environment check failed. Exiting script."
    }

    Connect-ToExchangeOnline

    if (Get-UserConfirmation -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -StartDate $StartDate -EndDate $EndDate -TestMode:$TestMode) {
        Forward-Emails -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -StartDate $StartDate -EndDate $EndDate -TestMode:$TestMode
    }
    else {
        Write-Log "Operation cancelled by user."
    }
}
catch {
    Write-Log "Unexpected error: $_"
}
finally {
    # Ensure disconnection even if an error occurs
    if (Get-PSSession | Where-Object {$_.Name -like "ExchangeOnlineInternalSession*"}) {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Log "Disconnected from Exchange Online."
    }
    Write-Log "Script completed"
}