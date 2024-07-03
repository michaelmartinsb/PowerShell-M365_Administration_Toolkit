<#
.SYNOPSIS
    Forwards emails from a specified Microsoft 365 mailbox to another mailbox based on a date range.

.DESCRIPTION
    This script uses Exchange Online PowerShell to forward emails from a source mailbox to a target mailbox
    within a specified date range. It uses Compliance Search features to ensure each email is forwarded individually.

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

# Function to connect to Exchange Online and Security & Compliance Center
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

function Connect-ToSecurityComplianceCenter {
    try {
        Import-Module ExchangeOnlineManagement
        $UserCredential = Get-Credential -Message "Enter your Office 365 credentials for Security & Compliance Center"
        Connect-IPPSSession -Credential $UserCredential
        Write-Log "Successfully connected to Security & Compliance Center."
    }
    catch {
        Write-Log "Error connecting to Security & Compliance Center: $_"
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

        # Wait for search to complete
        $timeout = [DateTime]::Now.AddMinutes(30)
        do {
            Start-Sleep -Seconds 30
            if (-not $TestMode) {
                $searchStatus = Get-ComplianceSearch -Identity $searchName
                Write-Log "Search status: $($searchStatus.Status)" -Verbose
            } else {
                Write-Log "[TEST MODE] Simulating search status check" -Verbose
            }

            if ([DateTime]::Now -gt $timeout) {
                throw "Search timed out after 30 minutes."
            }
        } while ($TestMode -or $searchStatus.Status -ne "Completed")

        Write-Log "Starting compliance search action" -Verbose
        if (-not $TestMode) {
            $actionName = "${searchName}_Action"
            New-ComplianceSearchAction -SearchName $searchName -ActionName $actionName -Export
            
            # Wait for the action to complete
            do {
                Start-Sleep -Seconds 30
                $actionStatus = Get-ComplianceSearchAction -Identity $actionName
                Write-Log "Action status: $($actionStatus.Status)" -Verbose
                
                if ([DateTime]::Now -gt $timeout) {
                    throw "Action timed out after 30 minutes."
                }
            } while ($actionStatus.Status -ne "Completed")
            
            # Get the results and forward them
            $results = Get-ComplianceSearchAction -Identity $actionName | Select-Object -ExpandProperty Results
            $emailIds = $results | Select-String -Pattern "Subject:.*?Location:.*?Item:\s*(\S+)" -AllMatches | 
                        ForEach-Object { $_.Matches.Groups[1].Value }
            
            foreach ($emailId in $emailIds) {
                Write-Log "Forwarding email with ID: $emailId" -Verbose
                if (-not $TestMode) {
                    Search-Mailbox -Identity $SourceMailbox -SearchQuery "ItemId:$emailId" -TargetMailbox $TargetMailbox -TargetFolder "Inbox" -LogOnly:$false
                } else {
                    Write-Log "[TEST MODE] Would forward email with ID: $emailId" -Verbose
                }
            }
            
            Write-Log "Email forwarding completed successfully."
        } else {
            Write-Log "[TEST MODE] Would forward search results to $TargetMailbox" -Verbose
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
    Write-Log "Script started in $(if ($TestMode) { 'TEST' } else { 'LIVE' }) mode"

    # Check environment
    if (-not (Test-Environment)) {
        throw "Environment check failed. Exiting script."
    }

    Connect-ToExchangeOnline
    Connect-ToSecurityComplianceCenter

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
