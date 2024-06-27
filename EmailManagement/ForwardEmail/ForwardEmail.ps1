<#
.SYNOPSIS
    Forwards emails from a specified Microsoft 365 mailbox to another mailbox based on a date range.

.DESCRIPTION
    This script uses Exchange Online PowerShell to forward emails from a source mailbox to a target mailbox
    within a specified date range. It uses the Compliance Search feature for email discovery and forwarding.
    This script is designed for forwarding emails to external (third-party) mailboxes.

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
    This script requires the ExchangeOnlineManagement module and appropriate permissions to perform compliance searches.
    You may need to run Connect-IPPSSession before running this script to access compliance cmdlets.
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

        # Check if compliance cmdlets are available
        if (!(Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue)) {
            Write-Log "Compliance cmdlets are not available. You may need to run Connect-IPPSSession first."
            Write-Log "Please run the following command and then re-run this script:"
            Write-Log "Connect-IPPSSession -UserPrincipalName your_admin@yourdomain.com"
            throw "Compliance cmdlets not available."
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

# ... [Rest of the script remains the same] ...

# Main script execution
try {
    Write-Log "Script started in $(if ($TestMode) { 'TEST' } else { 'LIVE' }) mode"

    # Check environment
    if (-not (Test-Environment)) {
        throw "Environment check failed. Exiting script."
    }

    Connect-ToExchangeOnline

    if (Get-UserConfirmation -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -StartDate $StartDate -EndDate $EndDate -TestMode:$TestMode) {
        Forward-EmailsUsingComplianceSearch -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -StartDate $StartDate -EndDate $EndDate -TestMode:$TestMode
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