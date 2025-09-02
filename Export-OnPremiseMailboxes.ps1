<#
.SYNOPSIS
    Exports email addresses from on-premise Exchange mailboxes to JSON format.

.DESCRIPTION
    This script connects to an on-premise Exchange server and exports all email addresses
    (primary SMTP and aliases) for specified mailboxes to a JSON file. The script supports
    various scoping options including database, organizational unit, and custom filters.

.PARAMETER Server
    Exchange server to connect to. If not specified, auto-discovery will be used.

.PARAMETER Database
    Scope export to mailboxes in a specific database.

.PARAMETER OrganizationalUnit
    Scope export to mailboxes in a specific organizational unit.

.PARAMETER Filter
    Custom filter for selecting mailboxes (e.g., "Department -eq 'Sales'").

.PARAMETER OutputPath
    Path for the output JSON file. Default is ".\mailbox-addresses.json".

.PARAMETER Credential
    Credentials for Exchange connection. If not provided, current user credentials will be used.

.EXAMPLE
    .\Export-OnPremiseMailboxes.ps1 -Database "Mailbox Database 01"
    
.EXAMPLE
    .\Export-OnPremiseMailboxes.ps1 -OrganizationalUnit "OU=Users,OU=Corporate,DC=contoso,DC=com"

.EXAMPLE
    .\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales' -and Office -eq 'New York'"

.NOTES
    Author: Exchange Migration Tool
    Version: 1.0
    Requires: Exchange Management Shell or Remote PowerShell to Exchange
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Server,
    
    [Parameter(Mandatory = $false)]
    [string]$Database,
    
    [Parameter(Mandatory = $false)]
    [string]$OrganizationalUnit,
    
    [Parameter(Mandatory = $false)]
    [string]$Filter,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\mailbox-addresses.json",
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to establish Exchange connection
function Connect-ExchangeServer {
    param(
        [string]$ServerName,
        [System.Management.Automation.PSCredential]$Cred
    )
    
    try {
        Write-ColorOutput "Connecting to Exchange server..." "Yellow"
        
        # Check if Exchange Management Shell is already loaded
        if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
            Write-ColorOutput "Exchange Management Shell already loaded." "Green"
            return $true
        }
        
        # Try to load Exchange Management Shell
        if (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) {
            Write-ColorOutput "Exchange Management Shell loaded successfully." "Green"
            return $true
        }
        
        # Try to add Exchange snapin
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
            Write-ColorOutput "Exchange Management Shell snapin loaded." "Green"
            return $true
        }
        catch {
            Write-ColorOutput "Failed to load Exchange snapin. Trying remote connection..." "Yellow"
        }
        
        # Try remote PowerShell connection
        if ($ServerName) {
            $sessionParams = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri = "http://$ServerName/PowerShell/"
                Authentication = 'Kerberos'
            }
            
            if ($Cred) {
                $sessionParams.Credential = $Cred
            }
            
            $session = New-PSSession @sessionParams
            Import-PSSession $session -DisableNameChecking
            Write-ColorOutput "Connected to Exchange server via remote PowerShell." "Green"
            return $true
        }
        
        throw "Unable to establish Exchange connection."
    }
    catch {
        Write-ColorOutput "Error connecting to Exchange: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to get mailboxes based on parameters
function Get-TargetMailboxes {
    param(
        [string]$DatabaseName,
        [string]$OrgUnit,
        [string]$FilterString
    )
    
    try {
        Write-ColorOutput "Retrieving target mailboxes..." "Yellow"
        
        $getMailboxParams = @{
            ResultSize = 'Unlimited'
        }
        
        # Add scoping parameters
        if ($DatabaseName) {
            $getMailboxParams.Database = $DatabaseName
            Write-ColorOutput "Scoping to database: $DatabaseName" "Cyan"
        }
        
        if ($OrgUnit) {
            $getMailboxParams.OrganizationalUnit = $OrgUnit
            Write-ColorOutput "Scoping to OU: $OrgUnit" "Cyan"
        }
        
        if ($FilterString) {
            $getMailboxParams.Filter = $FilterString
            Write-ColorOutput "Applying filter: $FilterString" "Cyan"
        }
        
        $mailboxes = Get-Mailbox @getMailboxParams
        Write-ColorOutput "Found $($mailboxes.Count) mailboxes matching criteria." "Green"
        
        return $mailboxes
    }
    catch {
        Write-ColorOutput "Error retrieving mailboxes: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to export mailbox data
function Export-MailboxData {
    param(
        [array]$Mailboxes,
        [string]$OutputFile
    )
    
    try {
        Write-ColorOutput "Processing mailbox data..." "Yellow"
        
        $exportData = @{
            ExportInfo = @{
                ExportDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                ExportedBy = $env:USERNAME
                Server = $env:COMPUTERNAME
                TotalMailboxes = $Mailboxes.Count
                ScopingCriteria = @{
                    Database = $Database
                    OrganizationalUnit = $OrganizationalUnit
                    Filter = $Filter
                }
            }
            Mailboxes = @()
        }
        
        $counter = 0
        foreach ($mailbox in $Mailboxes) {
            $counter++
            $percentComplete = [math]::Round(($counter / $Mailboxes.Count) * 100, 2)
            Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.DisplayName)" -PercentComplete $percentComplete
            
            try {
                # Get all email addresses for the mailbox
                $emailAddresses = $mailbox.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" -or $_.PrefixString -eq "smtp" }
                $primarySMTP = ($emailAddresses | Where-Object { $_.IsPrimaryAddress -eq $true }).SmtpAddress
                $aliases = ($emailAddresses | Where-Object { $_.IsPrimaryAddress -eq $false }).SmtpAddress
                
                $mailboxData = @{
                    DisplayName = $mailbox.DisplayName
                    Alias = $mailbox.Alias
                    PrimarySMTPAddress = $primarySMTP
                    EmailAliases = @($aliases)
                    Database = $mailbox.Database.Name
                    OrganizationalUnit = $mailbox.OrganizationalUnit
                    ExportDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                $exportData.Mailboxes += $mailboxData
                
                Write-Verbose "Processed: $($mailbox.DisplayName) - Primary: $primarySMTP - Aliases: $($aliases.Count)"
            }
            catch {
                Write-ColorOutput "Warning: Failed to process mailbox $($mailbox.DisplayName): $($_.Exception.Message)" "Yellow"
            }
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
        
        # Convert to JSON and save
        Write-ColorOutput "Saving data to JSON file..." "Yellow"
        $jsonData = $exportData | ConvertTo-Json -Depth 10
        $jsonData | Out-File -FilePath $OutputFile -Encoding UTF8
        
        Write-ColorOutput "Export completed successfully!" "Green"
        Write-ColorOutput "File saved: $OutputFile" "Green"
        Write-ColorOutput "Total mailboxes exported: $($exportData.Mailboxes.Count)" "Green"
        
        # Display summary
        $aliasCount = ($exportData.Mailboxes | Measure-Object -Property @{Expression={$_.EmailAliases.Count}} -Sum).Sum
        Write-ColorOutput "Total aliases exported: $aliasCount" "Green"
        
        return $exportData
    }
    catch {
        Write-ColorOutput "Error during export: $($_.Exception.Message)" "Red"
        throw
    }
}

# Main execution
try {
    Write-ColorOutput "=== Exchange Mailbox Address Export Tool ===" "Cyan"
    Write-ColorOutput "Starting export process..." "White"
    
    # Validate parameters
    if ($Database -and $OrganizationalUnit) {
        Write-ColorOutput "Warning: Both Database and OrganizationalUnit specified. Database will take precedence." "Yellow"
    }
    
    # Connect to Exchange
    if (-not (Connect-ExchangeServer -ServerName $Server -Cred $Credential)) {
        throw "Failed to connect to Exchange server."
    }
    
    # Get target mailboxes
    $targetMailboxes = Get-TargetMailboxes -DatabaseName $Database -OrgUnit $OrganizationalUnit -FilterString $Filter
    
    if ($targetMailboxes.Count -eq 0) {
        Write-ColorOutput "No mailboxes found matching the specified criteria." "Yellow"
        exit 0
    }
    
    # Export mailbox data
    $exportResult = Export-MailboxData -Mailboxes $targetMailboxes -OutputFile $OutputPath
    
    Write-ColorOutput "=== Export Summary ===" "Cyan"
    Write-ColorOutput "Export Date: $($exportResult.ExportInfo.ExportDate)" "White"
    Write-ColorOutput "Total Mailboxes: $($exportResult.ExportInfo.TotalMailboxes)" "White"
    Write-ColorOutput "Output File: $OutputPath" "White"
    
    if ($Database) { Write-ColorOutput "Database: $Database" "White" }
    if ($OrganizationalUnit) { Write-ColorOutput "Organizational Unit: $OrganizationalUnit" "White" }
    if ($Filter) { Write-ColorOutput "Filter: $Filter" "White" }
    
    Write-ColorOutput "Export completed successfully!" "Green"
}
catch {
    Write-ColorOutput "Export failed: $($_.Exception.Message)" "Red"
    exit 1
}
finally {
    # Clean up any remote sessions
    Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
}
