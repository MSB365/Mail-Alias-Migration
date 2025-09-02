<#
.SYNOPSIS
    Imports email aliases from JSON file to Exchange Online mailboxes.

.DESCRIPTION
    This script reads a JSON file created by Export-OnPremiseMailboxes.ps1 and adds
    the email aliases to corresponding mailboxes in Exchange Online. The script only
    adds aliases and never removes existing ones. It matches mailboxes by primary SMTP address.

.PARAMETER JsonPath
    Path to the JSON file containing mailbox data from the export script.

.PARAMETER WhatIf
    Preview changes without applying them.

.PARAMETER Credential
    Credentials for Exchange Online connection.

.PARAMETER TenantId
    Azure AD Tenant ID for Exchange Online connection.

.EXAMPLE
    .\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json

.EXAMPLE
    .\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -WhatIf

.EXAMPLE
    .\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -TenantId "your-tenant-id"

.NOTES
    Author: Exchange Migration Tool
    Version: 1.0
    Requires: ExchangeOnlineManagement PowerShell module
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$JsonPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to connect to Exchange Online
function Connect-ExchangeOnline {
    param(
        [System.Management.Automation.PSCredential]$Cred,
        [string]$Tenant
    )
    
    try {
        Write-ColorOutput "Connecting to Exchange Online..." "Yellow"
        
        # Check if ExchangeOnlineManagement module is available
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-ColorOutput "ExchangeOnlineManagement module not found. Installing..." "Yellow"
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
        }
        
        # Import the module
        Import-Module ExchangeOnlineManagement -Force
        
        # Prepare connection parameters
        $connectParams = @{}
        
        if ($Cred) {
            $connectParams.Credential = $Cred
        }
        
        if ($Tenant) {
            $connectParams.DomainName = $Tenant
        }
        
        # Connect to Exchange Online
        Connect-ExchangeOnline @connectParams -ShowProgress $true
        
        Write-ColorOutput "Connected to Exchange Online successfully." "Green"
        return $true
    }
    catch {
        Write-ColorOutput "Error connecting to Exchange Online: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to load and validate JSON data
function Import-MailboxData {
    param(
        [string]$FilePath
    )
    
    try {
        Write-ColorOutput "Loading mailbox data from JSON file..." "Yellow"
        
        $jsonContent = Get-Content -Path $FilePath -Raw | ConvertFrom-Json
        
        # Validate JSON structure
        if (-not $jsonContent.Mailboxes) {
            throw "Invalid JSON format: Missing 'Mailboxes' property."
        }
        
        Write-ColorOutput "Loaded data for $($jsonContent.Mailboxes.Count) mailboxes." "Green"
        Write-ColorOutput "Export Date: $($jsonContent.ExportInfo.ExportDate)" "Cyan"
        Write-ColorOutput "Exported By: $($jsonContent.ExportInfo.ExportedBy)" "Cyan"
        
        return $jsonContent
    }
    catch {
        Write-ColorOutput "Error loading JSON file: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to process mailbox aliases
function Import-Aliases {
    param(
        [object]$MailboxData,
        [bool]$PreviewOnly = $false
    )
    
    try {
        $totalMailboxes = $MailboxData.Mailboxes.Count
        $processedCount = 0
        $successCount = 0
        $skippedCount = 0
        $errorCount = 0
        $aliasesAdded = 0
        
        Write-ColorOutput "Processing $totalMailboxes mailboxes..." "Yellow"
        
        foreach ($mailboxInfo in $MailboxData.Mailboxes) {
            $processedCount++
            $percentComplete = [math]::Round(($processedCount / $totalMailboxes) * 100, 2)
            Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailboxInfo.DisplayName)" -PercentComplete $percentComplete
            
            try {
                # Skip if no aliases to add
                if ($mailboxInfo.EmailAliases.Count -eq 0) {
                    Write-Verbose "Skipping $($mailboxInfo.DisplayName) - no aliases to add."
                    $skippedCount++
                    continue
                }
                
                # Find mailbox in Exchange Online by primary SMTP address
                $onlineMailbox = $null
                try {
                    $onlineMailbox = Get-Mailbox -Identity $mailboxInfo.PrimarySMTPAddress -ErrorAction Stop
                }
                catch {
                    Write-ColorOutput "Warning: Mailbox not found in Exchange Online: $($mailboxInfo.PrimarySMTPAddress)" "Yellow"
                    $skippedCount++
                    continue
                }
                
                # Get current email addresses
                $currentAddresses = $onlineMailbox.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" -or $_.PrefixString -eq "smtp" }
                $currentSMTPAddresses = $currentAddresses.SmtpAddress
                
                # Determine which aliases need to be added
                $aliasesToAdd = @()
                foreach ($alias in $mailboxInfo.EmailAliases) {
                    if ($alias -notin $currentSMTPAddresses) {
                        $aliasesToAdd += $alias
                    }
                }
                
                if ($aliasesToAdd.Count -eq 0) {
                    Write-Verbose "Skipping $($mailboxInfo.DisplayName) - all aliases already exist."
                    $skippedCount++
                    continue
                }
                
                # Display what will be added
                Write-ColorOutput "Mailbox: $($mailboxInfo.DisplayName) ($($mailboxInfo.PrimarySMTPAddress))" "White"
                Write-ColorOutput "  Aliases to add: $($aliasesToAdd -join ', ')" "Cyan"
                
                if (-not $PreviewOnly) {
                    # Add the aliases
                    $newAddresses = $onlineMailbox.EmailAddresses
                    foreach ($aliasToAdd in $aliasesToAdd) {
                        $newAddresses += "smtp:$aliasToAdd"
                    }
                    
                    Set-Mailbox -Identity $onlineMailbox.Identity -EmailAddresses $newAddresses
                    Write-ColorOutput "  Successfully added $($aliasesToAdd.Count) aliases." "Green"
                    $aliasesAdded += $aliasesToAdd.Count
                }
                else {
                    Write-ColorOutput "  [PREVIEW] Would add $($aliasesToAdd.Count) aliases." "Yellow"
                    $aliasesAdded += $aliasesToAdd.Count
                }
                
                $successCount++
            }
            catch {
                Write-ColorOutput "Error processing $($mailboxInfo.DisplayName): $($_.Exception.Message)" "Red"
                $errorCount++
            }
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
        
        # Display summary
        Write-ColorOutput "=== Import Summary ===" "Cyan"
        Write-ColorOutput "Total Mailboxes Processed: $processedCount" "White"
        Write-ColorOutput "Successfully Processed: $successCount" "Green"
        Write-ColorOutput "Skipped: $skippedCount" "Yellow"
        Write-ColorOutput "Errors: $errorCount" "Red"
        Write-ColorOutput "Total Aliases Added: $aliasesAdded" "Green"
        
        if ($PreviewOnly) {
            Write-ColorOutput "*** PREVIEW MODE - No changes were made ***" "Yellow"
        }
        
        return @{
            ProcessedCount = $processedCount
            SuccessCount = $successCount
            SkippedCount = $skippedCount
            ErrorCount = $errorCount
            AliasesAdded = $aliasesAdded
        }
    }
    catch {
        Write-ColorOutput "Error during import process: $($_.Exception.Message)" "Red"
        throw
    }
}

# Main execution
try {
    Write-ColorOutput "=== Exchange Online Alias Import Tool ===" "Cyan"
    
    if ($WhatIf) {
        Write-ColorOutput "Running in PREVIEW mode - no changes will be made." "Yellow"
    }
    
    # Load mailbox data from JSON
    $mailboxData = Import-MailboxData -FilePath $JsonPath
    
    # Connect to Exchange Online
    if (-not (Connect-ExchangeOnline -Cred $Credential -Tenant $TenantId)) {
        throw "Failed to connect to Exchange Online."
    }
    
    # Process the aliases
    $result = Import-Aliases -MailboxData $mailboxData -PreviewOnly:$WhatIf
    
    if ($result.ErrorCount -eq 0) {
        Write-ColorOutput "Import completed successfully!" "Green"
    }
    else {
        Write-ColorOutput "Import completed with $($result.ErrorCount) errors. Check the output above for details." "Yellow"
    }
}
catch {
    Write-ColorOutput "Import failed: $($_.Exception.Message)" "Red"
    exit 1
}
finally {
    # Disconnect from Exchange Online
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-ColorOutput "Disconnected from Exchange Online." "Green"
    }
    catch {
        # Ignore disconnect errors
    }
}
