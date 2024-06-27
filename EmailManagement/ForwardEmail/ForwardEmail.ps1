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

# Function to forward emails using Compliance Search
function Forward-EmailsUsingComplianceSearch {
    param (
        [string]$SourceMailbox,
        [string]$TargetMailbox,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$TestMode
    )

    try {
        $searchName = "ForwardEmails_$(Get-Date -Format 'yyyyMMddHHmmss')"
        $searchQuery = "Received:>=$($StartDate.ToString('MM/dd/yyyy')) AND Received:<=$($EndDate.ToString('MM/dd/yyyy'))"

        Write-Log "Creating compliance search: $searchName" -Verbose
        if (-not $TestMode) {
            New-ComplianceSearch -Name $searchName -ExchangeLocation $SourceMailbox -ContentMatchQuery $searchQuery
        } else {
            Write-Log "[TEST MODE] Would create compliance search: $searchName" -Verbose
        }
        
        Write-Log "Starting compliance search" -Verbose
        if (-not $TestMode) {
            Start-ComplianceSearch -Identity $searchName
        } else {
            Write-Log "[TEST MODE] Would start compliance search" -Verbose
        }

        # Simulate search completion and status checks
        $sleepTime = 5
        $maxSleepTime = 60
        $timeout = [DateTime]::Now.AddMinutes(5) # Shortened for test mode
        do {
            Start-Sleep -Seconds $sleepTime
            if (-not $TestMode) {
                $searchStatus = Get-ComplianceSearch -Identity $searchName
                Write-Log "Search status: $($searchStatus.Status)" -Verbose
            } else {
                Write-Log "[TEST MODE] Simulating search status check" -Verbose
            }
            $sleepTime = [Math]::Min($sleepTime * 2, $maxSleepTime)

            if ([DateTime]::Now -gt $timeout) {
                Write-Log "Search simulation completed after 5 minutes."
                break
            }
        } while ($TestMode -or $searchStatus.Status -ne "Completed")

        Write-Log "Exporting search results to $TargetMailbox" -Verbose
        if (-not $TestMode) {
            New-ComplianceSearchAction -SearchName $searchName -Action Export -ExchangeLocation $TargetMailbox
            Write-Log "Email forwarding completed successfully."
        } else {
            Write-Log "[TEST MODE] Would export search results to $TargetMailbox" -Verbose
            Write-Log "[TEST MODE] Email forwarding simulation completed successfully."
        }
    }
    catch {
        Write-Log "Error in compliance search process: $_"
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
    Write-Log "Script started in $($TestMode ? 'TEST' : 'LIVE') mode"

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