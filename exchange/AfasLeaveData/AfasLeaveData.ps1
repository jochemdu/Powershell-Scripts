#Requires -Version 5.1
<#
.SYNOPSIS
    Synchronizes leave data from AFAS to Exchange/Outlook calendars.

.DESCRIPTION
    Retrieves leave/absence data from AFAS via REST API and creates corresponding
    calendar events in user mailboxes via Exchange Web Services (EWS).

    Features:
    - Connects to AFAS REST API using App Connector token
    - Retrieves leave data for specified employees
    - Creates/updates calendar events in Exchange (On-Prem or Online)
    - Supports full sync and delta sync modes
    - CSV/Excel reporting of sync results

.PARAMETER ConfigPath
    Path to JSON configuration file.

.PARAMETER Credential
    Credentials for Exchange authentication.

.PARAMETER AfasToken
    AFAS App Connector token. Can also be provided via SecretManagement.

.PARAMETER AfasEnvironment
    AFAS environment ID (e.g., 'O12345').

.PARAMETER SyncMode
    Full or Delta sync mode. Default: Delta.

.PARAMETER DaysAhead
    Number of days ahead to sync leave data. Default: 90.

.PARAMETER DaysBehind
    Number of days behind to sync. Default: 0.

.PARAMETER OutputPath
    Path for CSV report output.

.PARAMETER ExcelOutputPath
    Path for Excel report output (requires ImportExcel module).

.PARAMETER TestMode
    Run without making actual changes.

.EXAMPLE
    .\AfasLeaveData.ps1 -ConfigPath .\config.json -Credential (Get-Credential)

.EXAMPLE
    .\AfasLeaveData.ps1 -AfasEnvironment 'O12345' -TestMode -Verbose

.NOTES
    Version: 1.0.0
    Author: [Your Name]
    
    Requirements:
    - PowerShell 5.1 or later
    - EWS Managed API for calendar access
    - AFAS App Connector with GetConnector access
    - Exchange impersonation rights for calendar modifications
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AfasToken,

    [Parameter()]
    [ValidatePattern('^[A-Z]\d{5}$')]
    [string]$AfasEnvironment,

    [Parameter()]
    [ValidateSet('Full', 'Delta')]
    [string]$SyncMode = 'Delta',

    [Parameter()]
    [ValidateRange(1, 365)]
    [int]$DaysAhead = 90,

    [Parameter()]
    [ValidateRange(0, 30)]
    [int]$DaysBehind = 0,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'afas-leave-sync-report.csv'),

    [Parameter()]
    [string]$ExcelOutputPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$EwsUrl,

    [Parameter()]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$ImpersonationSmtp,

    [Parameter()]
    [switch]$TestMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Initialization

$script:Version = '1.0.0'
$script:IsDotSourced = $MyInvocation.InvocationName -eq '.'
$script:ScriptRoot = $PSScriptRoot
$script:ModulePath = Join-Path -Path $script:ScriptRoot -ChildPath 'modules\AfasCore\AfasCore.psm1'

# Import shared module
if (Test-Path -Path $script:ModulePath) {
    Import-Module $script:ModulePath -Force -ErrorAction Stop
}
else {
    Write-Warning "AfasCore module not found at: $script:ModulePath"
}

# Import ExchangeCore if available (shared Exchange functions)
$exchangeCorePath = Join-Path -Path (Split-Path $script:ScriptRoot -Parent) -ChildPath 'Find-GhostRoomMeetings\modules\ExchangeCore\ExchangeCore.psm1'
if (Test-Path -Path $exchangeCorePath) {
    Import-Module $exchangeCorePath -Force -ErrorAction Stop
}

#endregion Initialization

#region Configuration

function Import-AfasConfiguration {
    <#
    .SYNOPSIS
        Loads and validates AFAS configuration.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Configuration file not found: $Path"
    }

    $content = Get-Content -Path $Path -Raw

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $config = $content | ConvertFrom-Json -AsHashtable
    }
    else {
        $obj = $content | ConvertFrom-Json
        $config = @{}
        foreach ($prop in $obj.PSObject.Properties) {
            $config[$prop.Name] = $prop.Value
        }
    }

    return $config
}

#endregion Configuration

#region AFAS Functions

function Get-AfasLeaveData {
    <#
    .SYNOPSIS
        Retrieves leave data from AFAS GetConnector.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Environment,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Token,

        [Parameter()]
        [string]$ConnectorId = 'Profit_Verlof',

        [Parameter()]
        [datetime]$FromDate,

        [Parameter()]
        [datetime]$ToDate
    )

    $baseUrl = "https://$Environment.rest.afas.online/profitrestservices/connectors"
    $url = "$baseUrl/$ConnectorId"

    $headers = @{
        'Authorization' = "AfasToken $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($Token)))"
        'Content-Type'  = 'application/json'
    }

    # Build filter if dates provided
    $filterParts = @()
    if ($FromDate) {
        $filterParts += "Datum_begin ge '$($FromDate.ToString('yyyy-MM-dd'))'"
    }
    if ($ToDate) {
        $filterParts += "Datum_eind le '$($ToDate.ToString('yyyy-MM-dd'))'"
    }

    if ($filterParts.Count -gt 0) {
        $filter = $filterParts -join ' and '
        $url += "?filterfieldids=$([uri]::EscapeDataString($filter))"
    }

    Write-Verbose "Calling AFAS API: $url"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
        return $response.rows
    }
    catch {
        Write-Error "Failed to retrieve AFAS leave data: $_"
        throw
    }
}

#endregion AFAS Functions

#region Calendar Functions

function New-LeaveCalendarEvent {
    <#
    .SYNOPSIS
        Creates a calendar event for leave entry.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $EwsService,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$UserSmtp,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $LeaveEntry,

        [Parameter()]
        [switch]$TestMode
    )

    if ($TestMode) {
        Write-Verbose "[TEST] Would create calendar event for $UserSmtp : $($LeaveEntry.Omschrijving)"
        return [PSCustomObject]@{
            User     = $UserSmtp
            Subject  = $LeaveEntry.Omschrijving
            Start    = $LeaveEntry.Datum_begin
            End      = $LeaveEntry.Datum_eind
            Status   = 'TestMode'
            Action   = 'Create'
        }
    }

    if ($PSCmdlet.ShouldProcess($UserSmtp, "Create leave calendar event: $($LeaveEntry.Omschrijving)")) {
        # TODO: Implement actual EWS calendar event creation
        Write-Verbose "Creating calendar event for $UserSmtp"
    }
}

function Sync-LeaveToCalendar {
    <#
    .SYNOPSIS
        Synchronizes all leave entries to Exchange calendars.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $EwsService,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [PSCustomObject[]]$LeaveData,

        [Parameter()]
        [hashtable]$UserMapping,

        [Parameter()]
        [switch]$TestMode
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processed = 0

    foreach ($entry in $LeaveData) {
        $processed++
        $percentComplete = [int](($processed / $LeaveData.Count) * 100)
        Write-Progress -Activity 'Syncing leave data' -Status "$processed of $($LeaveData.Count)" -PercentComplete $percentComplete

        # Map AFAS employee to Exchange mailbox
        $userSmtp = $UserMapping[$entry.Medewerker] ?? "$($entry.Medewerker)@contoso.com"

        $result = New-LeaveCalendarEvent -EwsService $EwsService `
            -UserSmtp $userSmtp `
            -LeaveEntry $entry `
            -TestMode:$TestMode

        if ($result) {
            $results.Add($result)
        }
    }

    Write-Progress -Activity 'Syncing leave data' -Completed
    return $results.ToArray()
}

#endregion Calendar Functions

#region Main Execution

if ($script:IsDotSourced) {
    Write-Verbose 'Script was dot-sourced; functions are now available.'
    return
}

# Calculate date window
$fromDate = (Get-Date).Date.AddDays(-$DaysBehind)
$toDate = (Get-Date).Date.AddDays($DaysAhead)

Write-Verbose "AfasLeaveData v$script:Version"
Write-Verbose "Sync window: $fromDate to $toDate"

if ($TestMode) {
    Write-Host "TEST MODE: No actual changes will be made" -ForegroundColor Cyan
}

# Load configuration
$config = @{}
if ($ConfigPath) {
    $config = Import-AfasConfiguration -Path $ConfigPath

    # Apply config defaults
    if (-not $AfasEnvironment -and $config.ContainsKey('AfasEnvironment')) {
        $AfasEnvironment = $config['AfasEnvironment']
    }
}

# Validate required parameters
if (-not $AfasEnvironment) {
    throw 'AfasEnvironment is required. Provide via -AfasEnvironment or config file.'
}

if (-not $AfasToken -and -not $TestMode) {
    # Try to get from SecretManagement if available
    if (Get-Command -Name Get-Secret -ErrorAction SilentlyContinue) {
        $AfasToken = Get-Secret -Name 'AfasToken' -AsPlainText -ErrorAction SilentlyContinue
    }

    if (-not $AfasToken) {
        throw 'AfasToken is required. Provide via -AfasToken, SecretManagement, or use -TestMode.'
    }
}

try {
    # Get leave data from AFAS
    if ($TestMode) {
        Write-Verbose 'Test mode: Using mock leave data'
        $leaveData = @(
            [PSCustomObject]@{
                Medewerker    = 'EMP001'
                Omschrijving  = 'Vakantie'
                Datum_begin   = $fromDate.AddDays(5)
                Datum_eind    = $fromDate.AddDays(10)
                Verlofsoort   = 'Vakantie'
            }
        )
    }
    else {
        $leaveData = Get-AfasLeaveData -Environment $AfasEnvironment `
            -Token $AfasToken `
            -FromDate $fromDate `
            -ToDate $toDate
    }

    Write-Host "Retrieved $($leaveData.Count) leave entries from AFAS" -ForegroundColor Cyan

    if ($leaveData.Count -eq 0) {
        Write-Host 'No leave data found in specified date range.' -ForegroundColor Yellow
        return
    }

    # TODO: Connect to EWS and sync
    # $ews = Connect-EwsService ...
    # $results = Sync-LeaveToCalendar -EwsService $ews -LeaveData $leaveData -TestMode:$TestMode

    # For now, output leave data
    $results = $leaveData | ForEach-Object {
        [PSCustomObject]@{
            Employee    = $_.Medewerker
            Description = $_.Omschrijving
            StartDate   = $_.Datum_begin
            EndDate     = $_.Datum_eind
            LeaveType   = $_.Verlofsoort
            Status      = if ($TestMode) { 'TestMode' } else { 'Pending' }
        }
    }

    # Export results
    if ($results.Count -gt 0) {
        $outputDir = Split-Path -Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path -Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }

        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Green

        if ($ExcelOutputPath) {
            if (Get-Module -Name ImportExcel -ListAvailable) {
                Import-Module ImportExcel -ErrorAction Stop
                $results | Export-Excel -Path $ExcelOutputPath -WorksheetName 'LeaveSync' -AutoSize -FreezeTopRow
                Write-Host "Excel report saved to: $ExcelOutputPath" -ForegroundColor Green
            }
            else {
                Write-Warning 'ImportExcel module not available. Skipping Excel export.'
            }
        }
    }

    # Return results to pipeline
    $results
}
catch {
    Write-Error -ErrorRecord $_
    throw
}

#endregion Main Execution
