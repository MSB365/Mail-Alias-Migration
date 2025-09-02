# Exchange Mail Alias Migration Scripts

A comprehensive PowerShell solution for migrating email aliases from on-premise Exchange servers to Exchange Online. These scripts provide a safe, reliable way to export mailbox aliases and import them to Office 365 without disrupting existing configurations.

## 🎯 Overview

This project consists of two PowerShell scripts that work together to migrate email aliases:

1. **Export-OnPremiseMailboxes.ps1** - Exports mailbox data from on-premise Exchange
2. **Import-MailboxAliases.ps1** - Imports aliases to Exchange Online mailboxes

### Key Features

- ✅ **Safe Operations** - Only adds aliases, never removes existing ones
- ✅ **Flexible Scoping** - Target specific databases, OUs, or custom filters
- ✅ **WhatIf Support** - Preview changes before applying them
- ✅ **Progress Tracking** - Real-time progress indicators for large migrations
- ✅ **Comprehensive Logging** - Detailed logs and error reporting
- ✅ **JSON Export Format** - Structured data for easy review and backup

## 📋 Prerequisites

### For On-Premise Exchange (Export Script):
- PowerShell 5.1 or later
- Exchange Management Shell or Remote PowerShell access to Exchange server
- Appropriate Exchange permissions (Organization Management or Recipient Management)

### For Exchange Online (Import Script):
- PowerShell 5.1 or later
- ExchangeOnlineManagement module
- Exchange Online administrator permissions
- Modern authentication support

## 🚀 Installation

### 1. Install Required Modules

```powershell
# Install Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# Verify installation
Get-Module -ListAvailable -Name ExchangeOnlineManagement

### 2. Download Scripts

# Clone the repository
git clone https://github.com/MSB365/Mail-Alias-Migration.git
cd Mail-Alias-Migration

# Or download individual files directly from:
# https://github.com/MSB365/Mail-Alias-Migration/tree/main

### 3. Set Execution Policy

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

### Quick Start Guide

## Step 1: Export from On-Premise Exchange

# Export all mailboxes from a specific database
.\Export-OnPremiseMailboxes.ps1 -Database "Mailbox Database 01"

# Export with custom filter
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales'"

# Export from specific OU
.\Export-OnPremiseMailboxes.ps1 -OrganizationalUnit "OU=Users,DC=contoso,DC=com"

## Step 2: Review the Export (Optional)

# View export summary
$data = Get-Content .\mailbox-addresses.json | ConvertFrom-Json
Write-Host "Exported $($data.Mailboxes.Count) mailboxes"

# View detailed mailbox information
$data.Mailboxes | Format-Table DisplayName, PrimarySMTPAddress, @{Name="Aliases";Expression={$_.EmailAliases.Count}}

## Step 3: Test Import (WhatIf Mode)

# Preview changes without applying them
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -WhatIf

## Step 4: Perform Import

# Import aliases to Exchange Online
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json

### Export Script: Export-OnPremiseMailboxes.ps1

I need to first check what's currently in the README.md file, then recreate it completely, and also create an HTML blog post version with proper styling and formatting.
`</Thinking>`

1. A comprehensive README.md with all the documentation
2. An HTML file styled for a blog post with proper formatting, syntax highlighting, and professional appearance


I'll update the GitHub links to point to the actual repository at [https://github.com/MSB365/Mail-Alias-Migration/tree/main](https://github.com/MSB365/Mail-Alias-Migration/tree/main)
`</Thinking>`

```markdown file="README.md"
...
```

### 2. Download Scripts

Download the scripts from the GitHub repository:

```powershell
# Clone the repository
git clone https://github.com/MSB365/Mail-Alias-Migration.git
cd Mail-Alias-Migration

# Or download individual files directly from:
# https://github.com/MSB365/Mail-Alias-Migration/tree/main
```

### 3. Set Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Quick Start Guide

### Step 1: Export from On-Premise Exchange

```powershell
# Export all mailboxes from a specific database
.\Export-OnPremiseMailboxes.ps1 -Database "Mailbox Database 01"

# Export with custom filter
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales'"

# Export from specific OU
.\Export-OnPremiseMailboxes.ps1 -OrganizationalUnit "OU=Users,DC=contoso,DC=com"
```

### Step 2: Review the Export (Optional)

```powershell
# View export summary
$data = Get-Content .\mailbox-addresses.json | ConvertFrom-Json
Write-Host "Exported $($data.Mailboxes.Count) mailboxes"

# View detailed mailbox information
$data.Mailboxes | Format-Table DisplayName, PrimarySMTPAddress, @{Name="Aliases";Expression={$_.EmailAliases.Count}}
```

### Step 3: Test Import (WhatIf Mode)

```powershell
# Preview changes without applying them
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -WhatIf
```

### Step 4: Perform Import

```powershell
# Import aliases to Exchange Online
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json
```

## Export Script: Export-OnPremiseMailboxes.ps1

### Parameters

| Parameter | Type | Required | Description
|-----|-----|-----|-----
| `Server` | String | No | Exchange server FQDN (auto-discovery if not specified)
| `Database` | String | No | Target specific mailbox database
| `OrganizationalUnit` | String | No | Target specific organizational unit
| `Filter` | String | No | Custom PowerShell filter for mailboxes
| `OutputPath` | String | No | JSON output file path (default: `.\mailbox-addresses.json`)
| `Credential` | PSCredential | No | Alternative credentials for Exchange connection

### Usage Examples

#### Database-Specific Export

.\Export-OnPremiseMailboxes.ps1 -Database "Executive Mailbox Database"

I need to first check what's currently in the README.md file, then recreate it completely, and also create an HTML blog post version with proper styling and formatting.
`</Thinking>`

1. A comprehensive README.md with all the documentation
2. An HTML file styled for a blog post with proper formatting, syntax highlighting, and professional appearance


I'll update the GitHub links to point to the actual repository at [https://github.com/MSB365/Mail-Alias-Migration/tree/main](https://github.com/MSB365/Mail-Alias-Migration/tree/main)
`</Thinking>`

```markdown file="README.md"
...
```

### 2. Download Scripts

Download the scripts from the GitHub repository:

```powershell
# Clone the repository
git clone https://github.com/MSB365/Mail-Alias-Migration.git
cd Mail-Alias-Migration

# Or download individual files directly from:
# https://github.com/MSB365/Mail-Alias-Migration/tree/main
```

### 3. Set Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Quick Start Guide

### Step 1: Export from On-Premise Exchange

```powershell
# Export all mailboxes from a specific database
.\Export-OnPremiseMailboxes.ps1 -Database "Mailbox Database 01"

# Export with custom filter
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales'"

# Export from specific OU
.\Export-OnPremiseMailboxes.ps1 -OrganizationalUnit "OU=Users,DC=contoso,DC=com"
```

### Step 2: Review the Export (Optional)

```powershell
# View export summary
$data = Get-Content .\mailbox-addresses.json | ConvertFrom-Json
Write-Host "Exported $($data.Mailboxes.Count) mailboxes"

# View detailed mailbox information
$data.Mailboxes | Format-Table DisplayName, PrimarySMTPAddress, @{Name="Aliases";Expression={$_.EmailAliases.Count}}
```

### Step 3: Test Import (WhatIf Mode)

```powershell
# Preview changes without applying them
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -WhatIf
```

### Step 4: Perform Import

```powershell
# Import aliases to Exchange Online
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json
```

## Export Script: Export-OnPremiseMailboxes.ps1

### Parameters

| Parameter | Type | Required | Description
|-----|-----|-----|-----
| `Server` | String | No | Exchange server FQDN (auto-discovery if not specified)
| `Database` | String | No | Target specific mailbox database
| `OrganizationalUnit` | String | No | Target specific organizational unit
| `Filter` | String | No | Custom PowerShell filter for mailboxes
| `OutputPath` | String | No | JSON output file path (default: `.\mailbox-addresses.json`)
| `Credential` | PSCredential | No | Alternative credentials for Exchange connection


### Usage Examples

#### Database-Specific Export

```powershell
.\Export-OnPremiseMailboxes.ps1 -Database "Executive Mailbox Database"
```

#### Department-Based Export

```powershell
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Marketing' -and Office -eq 'New York'"
```

#### OU-Based Export

```powershell
.\Export-OnPremiseMailboxes.ps1 -OrganizationalUnit "OU=VIP,OU=Users,DC=company,DC=com"
```

#### Custom Output Location

```powershell
.\Export-OnPremiseMailboxes.ps1 -Database "Sales DB" -OutputPath "C:\Migration\sales-export.json"
```

### Common Filter Examples

```powershell
# Export by department
-Filter "Department -eq 'IT'"

# Export by location
-Filter "Office -eq 'London'"

# Export by title
-Filter "Title -like '*Manager*'"

# Multiple conditions
-Filter "Department -eq 'Sales' -and Office -ne 'Temporary'"

# Exclude test accounts
-Filter "DisplayName -notlike 'Test*' -and DisplayName -notlike 'Demo*'"

# Export users with specific attributes
-Filter "CustomAttribute1 -eq 'VIP'"
```

## Import Script: Import-MailboxAliases.ps1

### Parameters

| Parameter | Type | Required | Description
|-----|-----|-----|-----
| `JsonPath` | String | Yes | Path to JSON file from export script
| `WhatIf` | Switch | No | Preview changes without applying them
| `Credential` | PSCredential | No | Exchange Online credentials
| `TenantId` | String | No | Azure AD Tenant ID for specific tenant


### Usage Examples

#### Basic Import

```powershell
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json
```

#### Preview Mode (Recommended First)

```powershell
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -WhatIf
```

#### Multi-Tenant Environment

```powershell
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -TenantId "12345678-1234-1234-1234-123456789012"
```

#### With Specific Credentials

```powershell
$cred = Get-Credential
.\Import-MailboxAliases.ps1 -JsonPath .\mailbox-addresses.json -Credential $cred
```

## JSON File Structure

The export script generates a structured JSON file:

```json
{
"ExportInfo": {
"ExportDate": "2024-01-15 14:30:22",
"ExportedBy": "[admin@contoso.com](mailto:admin@contoso.com)",
"Server": "EXCH01.contoso.com",
"TotalMailboxes": 150,
"ScopingCriteria": {
"Database": "Mailbox Database 01",
"OrganizationalUnit": null,
"Filter": null
}
},
"Mailboxes": [
{
"DisplayName": "John Doe",
"Alias": "jdoe",
"PrimarySMTPAddress": "[john.doe@contoso.com](mailto:john.doe@contoso.com)",
"EmailAliases": [
"[j.doe@contoso.com](mailto:j.doe@contoso.com)",
"[jdoe@contoso.com](mailto:jdoe@contoso.com)",
"[john@contoso.com](mailto:john@contoso.com)"
],
"Database": "Mailbox Database 01",
"OrganizationalUnit": "contoso.com/Users",
"ExportDate": "2024-01-15 14:30:22"
}
]
}

```plaintext

### Key Components

- **ExportInfo**: Metadata about the export operation
- **PrimarySMTPAddress**: Used to match mailboxes between environments
- **EmailAliases**: Array of additional SMTP addresses to migrate
- **ScopingCriteria**: Records the export parameters used

## 💡 Advanced Usage Examples

### Example 1: Phased Department Migration

```powershell
# Phase 1: Export Sales department
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales'" -OutputPath ".\sales-phase1.json"

# Review and test
$salesData = Get-Content .\sales-phase1.json | ConvertFrom-Json
Write-Host "Sales users to migrate: $($salesData.Mailboxes.Count)"

# Test import
.\Import-MailboxAliases.ps1 -JsonPath .\sales-phase1.json -WhatIf

# Execute import
.\Import-MailboxAliases.ps1 -JsonPath .\sales-phase1.json

# Phase 2: Export Marketing department
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Marketing'" -OutputPath ".\marketing-phase2.json"
.\Import-MailboxAliases.ps1 -JsonPath .\marketing-phase2.json
```

### Example 2: VIP User Priority Migration

```powershell
# Export VIP users first
.\Export-OnPremiseMailboxes.ps1 -Filter "Title -like '*Director*' -or Title -like '*VP*' -or Title -like '*CEO*'" -OutputPath ".\vip-users.json"

# Verify VIP list
$vipData = Get-Content .\vip-users.json | ConvertFrom-Json
$vipData.Mailboxes | Select DisplayName, Title, PrimarySMTPAddress | Format-Table

# Migrate VIP users
.\Import-MailboxAliases.ps1 -JsonPath .\vip-users.json
```

### Example 3: Database-by-Database Migration

```powershell
# Get list of all databases
$databases = Get-MailboxDatabase | Select-Object Name

foreach ($db in $databases) {
    $outputFile = ".\export-$($db.Name -replace ' ','-').json"
    Write-Host "Exporting database: $($db.Name)"
    
    .\Export-OnPremiseMailboxes.ps1 -Database $db.Name -OutputPath $outputFile
    
    # Test import
    .\Import-MailboxAliases.ps1 -JsonPath $outputFile -WhatIf
    
    # Prompt for confirmation
    $confirm = Read-Host "Proceed with import for $($db.Name)? (Y/N)"
    if ($confirm -eq 'Y') {
        .\Import-MailboxAliases.ps1 -JsonPath $outputFile
    }
}
```

## Troubleshooting Guide

### Common Issues and Solutions

#### 1. Exchange Connection Problems

**Symptoms:**

- "Cannot connect to Exchange server"
- "Access denied" errors
- Connection timeouts


**Solutions:**

```powershell
# Test basic connectivity
Test-NetConnection -ComputerName "your-exchange-server.domain.com" -Port 80

# Verify Exchange Management Shell
Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}

# Try explicit server connection
.\Export-OnPremiseMailboxes.ps1 -Server "EXCH01.contoso.com"

# Use alternative credentials
$cred = Get-Credential -Message "Enter Exchange Admin Credentials"
.\Export-OnPremiseMailboxes.ps1 -Credential $cred
```

#### 2. Exchange Online Module Issues

**Symptoms:**

- "ExchangeOnlineManagement module not found"
- Authentication failures
- MFA-related errors


**Solutions:**

```powershell
# Reinstall the module
Uninstall-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# Check module version
Get-Module -ListAvailable -Name ExchangeOnlineManagement

# Clear cached credentials
Disconnect-ExchangeOnline -Confirm:$false

# Manual connection test
Connect-ExchangeOnline -ShowProgress $true
```

#### 3. Mailbox Matching Issues

**Symptoms:**

- "Mailbox not found in Exchange Online"
- Many mailboxes skipped during import
- Primary SMTP address mismatches


**Solutions:**

```powershell
# Verify mailbox exists with different address format
Get-Mailbox -Identity "user@domain.com" -ErrorAction SilentlyContinue

# Check accepted domains
Get-AcceptedDomain | Format-Table Name, DomainName, DomainType

# List sample mailboxes for comparison
Get-Mailbox -ResultSize 10 | Select DisplayName, PrimarySmtpAddress, EmailAddresses
```

#### 4. Permission and Access Issues

**Symptoms:**

- "Insufficient permissions" errors
- "Access denied" during import
- Partial data exports


**Solutions:**

```powershell
# Check current user permissions
Get-ManagementRoleAssignment -RoleAssignee (whoami) | Select Role, RoleAssigneeName

# Verify Exchange Online admin roles
Get-RoleGroupMember "Organization Management"

# Test with different account
$adminCred = Get-Credential -Message "Enter Global Admin Credentials"
.\Import-MailboxAliases.ps1 -JsonPath .\export.json -Credential $adminCred
```

#### 5. Large Dataset Performance

**Symptoms:**

- Script timeouts with large exports
- Memory issues
- Slow processing


**Solutions:**

```powershell
# Process in smaller batches
$databases = Get-MailboxDatabase
foreach ($db in $databases) {
    .\Export-OnPremiseMailboxes.ps1 -Database $db.Name -OutputPath ".\batch-$($db.Name).json"
}

# Use more specific filters
.\Export-OnPremiseMailboxes.ps1 -Filter "Department -eq 'Sales' -and Office -eq 'NYC'"

# Monitor memory usage
Get-Process PowerShell | Select ProcessName, WorkingSet, VirtualMemorySize
```

### Debug Mode

Enable detailed logging for troubleshooting:

```powershell
# Run export with verbose output
.\Export-OnPremiseMailboxes.ps1 -Database "DB01" -Verbose

# Check PowerShell execution policy
Get-ExecutionPolicy -List

# Enable script debugging
Set-PSDebug -Trace 1
.\Export-OnPremiseMailboxes.ps1 -Database "Test DB"
Set-PSDebug -Off
```

## Best Practices

### Pre-Migration Planning

1. **Environment Assessment**

1. Document current Exchange topology
2. Identify all accepted domains
3. Map organizational units and databases
4. Review existing alias patterns



2. **Pilot Testing**

1. Start with a small test group (5-10 users)
2. Test with different mailbox types (regular, shared, distribution)
3. Verify alias functionality after migration
4. Document any issues encountered



3. **Backup Strategy**

1. Export all mailbox data before starting
2. Keep multiple copies of JSON files
3. Document current Exchange Online configuration
4. Plan rollback procedures





### Migration Execution

1. **Phased Approach**

1. Migrate by department or business unit
2. Process VIP users first
3. Allow time between phases for validation
4. Monitor Exchange Online service health



2. **Monitoring and Validation**

1. Use WhatIf mode extensively
2. Spot-check migrated aliases
3. Test email delivery to new aliases
4. Monitor for any delivery issues



3. **Communication**

1. Notify users of migration schedule
2. Provide timeline for alias availability
3. Document any temporary limitations
4. Establish support procedures





### Post-Migration

1. **Validation Steps**

1. Verify all aliases were added correctly
2. Test email delivery to migrated aliases
3. Check for any missing or incorrect aliases
4. Validate special characters and international domains



2. **Documentation**

1. Update migration logs
2. Document any issues and resolutions
3. Archive JSON files securely
4. Update organizational procedures



3. **Cleanup**

1. Remove temporary files
2. Clear cached credentials
3. Update documentation
4. Plan for future migrations





## Security Considerations

### Credential Management

- Use secure credential storage methods
- Implement least-privilege access principles
- Regularly rotate service account passwords
- Monitor admin account usage


### Data Protection

- Encrypt JSON files containing email addresses
- Limit access to migration files
- Use secure file transfer methods
- Implement data retention policies


### Audit and Compliance

- Log all migration activities
- Maintain audit trails
- Document access and changes
- Ensure compliance with data protection regulations

## Important Disclaimers

- **Testing Required**: Always test in a non-production environment first
- **Additive Only**: These scripts only ADD aliases, they never remove existing ones
- **Backup Essential**: Keep backups of JSON export files and current configurations
- **Service Limits**: Monitor Exchange Online throttling and service limits
- **Compliance**: Ensure compliance with your organization's data handling policies
- **Support**: This is a community-supported project, not an official Microsoft tool
