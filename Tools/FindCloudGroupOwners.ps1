#Requires -Version 7.0
<#
.SYNOPSIS
Lists Entra ID owners and members for cloud-only groups sourced from Excel.

.DESCRIPTION
Given an Excel workbook that includes group metadata columns (DisplayName, Id,
Source, GroupType), this script filters to groups marked as CloudOnly, looks
them up in Microsoft Entra ID via Microsoft Graph PowerShell, and prints the
owners and members that are currently assigned. No destructive actions are
performed; the script is intended as a validation aid before any clean-up.

.PARAMETER ExcelPath
Path to the .xlsx file that contains the group metadata.

.PARAMETER WorksheetName
Optional worksheet name to read when the workbook has multiple sheets.

.PARAMETER SkipModuleInstall
Prevents the script from attempting to install missing modules automatically.

.PARAMETER GraphScopes
Microsoft Graph permission scopes to request during Connect-MgGraph.

.NOTES
- Requires the ImportExcel community module for reading .xlsx files.
- Requires Microsoft.Graph PowerShell modules (Authentication + Groups).
- Set $DryRun to $false ONLY when you are ready for scripts that make changes.
- Example: .\FindCloudGroupOwners.ps1 -ExcelPath C:\Temp\Groups.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExcelPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WorksheetName,

    [Parameter()]
    [switch]$SkipModuleInstall,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]]$GraphScopes = @('Group.Read.All', 'GroupMember.Read.All')
)

# Toggle this switch if/when you introduce commands that modify resources.
# In dry-run mode the script only enumerates owners and members.
$DryRun = $true

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($SkipModuleInstall.IsPresent) {
            throw "Required module '$Name' is missing. Install it manually or rerun without -SkipModuleInstall."
        }

        Write-Host "Installing module '$Name' for the current user..."
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber | Out-Null
    }

    Import-Module -Name $Name -ErrorAction Stop | Out-Null
}

function Connect-GraphIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Scopes
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -eq $context -or -not $context.Account) {
        Write-Host "Connecting to Microsoft Graph with scopes [$($Scopes -join ', ')]..."
        $connectParams = @{
            Scopes = $Scopes
        }

        $connectCommand = Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue
        if ($connectCommand -and $connectCommand.Parameters.Keys -contains 'NoWelcome') {
            $connectParams['NoWelcome'] = $true
        }

        Connect-MgGraph @connectParams
    }
}

function Get-AdditionalPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Principal,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $propertyNames = $Principal.PSObject.Properties.Name
    if ($propertyNames -contains $PropertyName) {
        $value = $Principal.$PropertyName
        if ($null -ne $value -and $value -ne '') {
            return $value
        }
    }

    $additionalProperties = $null
    if ($propertyNames -contains 'AdditionalProperties') {
        $additionalProperties = $Principal.AdditionalProperties
    }

    if ($additionalProperties -and $additionalProperties.ContainsKey($PropertyName)) {
        $value = $additionalProperties[$PropertyName]
        if ($null -ne $value -and $value -ne '') {
            return $value
        }
    }

    return $null
}

function ConvertTo-FriendlyPrincipalString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Principal
    )

    $propertyNames = $Principal.PSObject.Properties.Name
    $additionalProperties = $null
    if ($propertyNames -contains 'AdditionalProperties') {
        $additionalProperties = $Principal.AdditionalProperties
    }

    $typeName = $null
    if ($propertyNames -contains '@odata.type') {
        $typeName = $Principal.'@odata.type'
    }
    elseif ($additionalProperties -and $additionalProperties.ContainsKey('@odata.type')) {
        $typeName = $additionalProperties['@odata.type']
    }

    if ($typeName) {
        $typeName = $typeName -replace '^#?microsoft\.graph\.', ''
    }
    else {
        $typeName = 'DirectoryObject'
    }

    $displayName = Get-AdditionalPropertyValue -Principal $Principal -PropertyName 'displayName'
    if (-not $displayName) {
        $displayName = '<no display name>'
    }

    $identifierCandidates = @(
        'userPrincipalName',
        'mail',
        'servicePrincipalName',
        'servicePrincipalNames',
        'deviceId',
        'id'
    )

    $identifier = $null
    foreach ($candidate in $identifierCandidates) {
        $value = Get-AdditionalPropertyValue -Principal $Principal -PropertyName $candidate
        if ($value) {
            if ($value -is [System.Array]) {
                $value = $value | Select-Object -First 1
            }
            $identifier = $value
            break
        }
    }

    if (-not $identifier) {
        $identifier = $Principal.Id
    }

    return "{0} ({1}) [{2}]" -f $displayName, $identifier, $typeName
}

function Show-PrincipalList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter()]
        [array]$Principals
    )

    Write-Host "$Label:"
    if (-not $Principals -or $Principals.Count -eq 0) {
        Write-Host '  - None found'
        return
    }

    foreach ($principal in $Principals) {
        Write-Host ("  - {0}" -f (ConvertTo-FriendlyPrincipalString -Principal $principal))
    }
}

function Import-CloudOnlyGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter()]
        [string]$Worksheet
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Excel file not found at '$Path'."
    }

    $resolvedPath = (Resolve-Path -Path $Path).Path

    $importParams = @{
        Path = $resolvedPath
    }

    if ($Worksheet) {
        $importParams['WorksheetName'] = $Worksheet
    }

    $rows = Import-Excel @importParams
    if (-not $rows) {
        Write-Warning "No rows were returned from '$resolvedPath'."
        return @()
    }

    return $rows |
        Where-Object {
            $sourceValue = if ($_.PSObject.Properties.Name -contains 'Source') { $_.Source } else { $null }
            if (-not $sourceValue) {
                return $false
            }

            $sourceValue.ToString().Trim() -ieq 'CloudOnly'
        }
}

# --- Main execution flow ----------------------------------------------------

Write-Host "DryRun mode is set to: $DryRun"
if ($DryRun) {
    Write-Host '[DryRun] The script only enumerates data; no write operations will run.'
}

$requiredModules = @(
    'ImportExcel',
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Groups'
)

foreach ($module in $requiredModules) {
    Ensure-Module -Name $module
}

Connect-GraphIfNeeded -Scopes $GraphScopes

$cloudOnlyGroups = Import-CloudOnlyGroups -Path $ExcelPath -Worksheet $WorksheetName

if (-not $cloudOnlyGroups -or $cloudOnlyGroups.Count -eq 0) {
    Write-Host 'No CloudOnly groups were found in the supplied workbook.'
    return
}

Write-Host ("Processing {0} cloud-only group(s)..." -f $cloudOnlyGroups.Count)

foreach ($group in $cloudOnlyGroups) {
    $displayName = if ($group.PSObject.Properties.Name -contains 'DisplayName' -and $group.DisplayName) {
        $group.DisplayName
    }
    else {
        '<no display name provided>'
    }

    $groupId = if ($group.PSObject.Properties.Name -contains 'Id') { $group.Id } else { $null }
    if (-not $groupId) {
        Write-Warning "Skipping a row because it does not contain an Id. DisplayName: $displayName"
        continue
    }

    $groupType = if ($group.PSObject.Properties.Name -contains 'GroupType' -and $group.GroupType) {
        $group.GroupType
    }
    else {
        '<not provided>'
    }

    Write-Host ''
    Write-Host ('=' * 72)
    Write-Host "Group DisplayName : $displayName"
    Write-Host "Group Id          : $groupId"
    Write-Host "GroupType         : $groupType"

    if ($DryRun) {
        Write-Host '[DryRun] Enumerating owners/members without making changes.'
    }

    try {
        $null = Get-MgGroup -GroupId $groupId -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to read group '$displayName' ($groupId). Error: $($_.Exception.Message)"
        continue
    }

    $ownerParams = @{
        GroupId           = $groupId
        All               = $true
        Property          = 'id,displayName,userPrincipalName,mail,servicePrincipalName,servicePrincipalNames'
        ConsistencyLevel  = 'eventual'
        ErrorAction       = 'Stop'
    }

    $memberParams = @{
        GroupId           = $groupId
        All               = $true
        Property          = 'id,displayName,userPrincipalName,mail,deviceId,servicePrincipalName,servicePrincipalNames'
        ConsistencyLevel  = 'eventual'
        ErrorAction       = 'Stop'
    }

    try {
        $owners = @(Get-MgGroupOwner @ownerParams)
    }
    catch {
        Write-Warning "Failed to retrieve owners for $displayName ($groupId): $($_.Exception.Message)"
        $owners = @()
    }

    try {
        $members = @(Get-MgGroupMember @memberParams)
    }
    catch {
        Write-Warning "Failed to retrieve members for $displayName ($groupId): $($_.Exception.Message)"
        $members = @()
    }

    Show-PrincipalList -Label 'Owners found' -Principals $owners
    Show-PrincipalList -Label 'Members found' -Principals $members
}

Write-Host ''
Write-Host 'Completed owner/member enumeration for all cloud-only groups.'
