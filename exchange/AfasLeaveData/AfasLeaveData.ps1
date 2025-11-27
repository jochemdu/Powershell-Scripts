#Requires -Version 5.1
<#
.SYNOPSIS
    Synchronizes leave data from AFAS Insite to Exchange calendars via Enterprise Service Bus.

.DESCRIPTION
    Retrieves leave/absence data from AFAS Insite via the Enterprise Service Bus REST API
    and creates/removes corresponding calendar events in user mailboxes via EWS.

    This script replicates the functionality of the legacy scripts:
    - Import-CalendarCSV.ps1 (create leave appointments)
    - Remove-CalendarItemsCSV.ps1 (remove canceled appointments)

    Features:
    - Retrieves leave data via Enterprise Service Bus REST API
    - Creates calendar items for new leave bookings
    - Removes calendar items for canceled leave
    - Uses Get-Mailbox for ITCode to email mapping
    - Password file authentication (same as legacy)
    - Tab-separated logging (compatible with legacy)
    - CSV file processing with Processed folder archival

.PARAMETER ConfigPath
    Path to JSON or PSD1 configuration file.

.PARAMETER Credential
    Credentials for Exchange and API authentication. If not provided,
    will be loaded from password file specified in config.

.PARAMETER Mode
    Operation mode: 'Import' (create appointments), 'Remove' (delete canceled), or 'Both'.
    Default: Both

.PARAMETER TestMode
    Run without making actual changes (dry run).

.PARAMETER Verbose
    Enable verbose output.

.EXAMPLE
    .\AfasLeaveData.ps1 -ConfigPath .\config.json
    
    Runs both import and remove operations using config file.

.EXAMPLE
    .\AfasLeaveData.ps1 -ConfigPath .\config.json -Mode Import
    
    Only imports new leave data to calendars.

.EXAMPLE
    .\AfasLeaveData.ps1 -ConfigPath .\config.json -TestMode -Verbose
    
    Dry run with verbose output.

.NOTES
    Version: 1.0.0
    Author: Jochem. Based on original scripts by Erwin Rook
    
    Requirements:
    - PowerShell 5.1 or later
    - EWS Managed API 2.2
    - Exchange Management Shell (for Get-Mailbox)
    - ApplicationImpersonation rights
    - Access to Enterprise Service Bus endpoints
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path $_ })]
    [string]$ConfigPath,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [ValidateSet('Import', 'Remove', 'Both')]
    [string]$Mode = 'Both',

    [Parameter()]
    [switch]$TestMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Script Variables

$script:Version = '1.0.0'
$script:LogFile = $null
$script:Config = $null
$script:EwsService = $null
$script:ExchangeSession = $null

#endregion Script Variables

#region Logging Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes a log entry in legacy-compatible format.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Context,
        
        [Parameter(Mandatory)]
        [ValidateSet('[START]', '[SUCCESS]', '[ERROR]', '[WARNING]', '[STOP]', '[INFO]')]
        [string]$Status,
        
        [Parameter(Mandatory)]
        [string]$Message
    )
    
    # Pad context to 20 characters (legacy format)
    if ($Context.Length -lt 20) {
        $Context = $Context.PadRight(20)
    }
    
    $logEntry = "$Context`t$Status`t$Message"
    
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8
    }
    
    # Also write to console based on status
    switch ($Status) {
        '[ERROR]'   { Write-Warning $Message }
        '[WARNING]' { Write-Warning $Message }
        '[SUCCESS]' { Write-Verbose $Message }
        default     { Write-Verbose $Message }
    }
}

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Creates a new log file with timestamp.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,
        
        [Parameter(Mandatory)]
        [string]$Prefix
    )
    
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format 'ddMMyyyy-HHmm'
    $script:LogFile = Join-Path -Path $LogPath -ChildPath "$Prefix-$timestamp.log"
    
    return $script:LogFile
}

#endregion Logging Functions

#region Configuration Functions

function Import-Configuration {
    <#
    .SYNOPSIS
        Loads configuration from JSON or PSD1 file.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    
    switch ($extension) {
        '.json' {
            $content = Get-Content -Path $Path -Raw -Encoding UTF8
            
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                return ($content | ConvertFrom-Json -AsHashtable)
            }
            else {
                # PowerShell 5.1 - manual conversion to hashtable
                $obj = $content | ConvertFrom-Json
                return (ConvertTo-Hashtable -InputObject $obj)
            }
        }
        '.psd1' {
            return (Import-PowerShellDataFile -Path $Path)
        }
        default {
            throw "Unsupported configuration file format: $extension"
        }
    }
}

function ConvertTo-Hashtable {
    <#
    .SYNOPSIS
        Recursively converts PSCustomObject to hashtable.
    #>
    param($InputObject)
    
    if ($null -eq $InputObject) { return $null }
    
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $collection = @(foreach ($item in $InputObject) { ConvertTo-Hashtable -InputObject $item })
        return $collection
    }
    
    if ($InputObject -is [PSCustomObject]) {
        $hash = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
        }
        return $hash
    }
    
    return $InputObject
}

function Get-CredentialFromFile {
    <#
    .SYNOPSIS
        Reads credentials from password file.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Username,
        
        [Parameter(Mandatory)]
        [string]$PasswordFilePath
    )
    
    if (-not (Test-Path -Path $PasswordFilePath)) {
        throw "Password file not found: $PasswordFilePath"
    }
    
    $securePassword = Get-Content -Path $PasswordFilePath | ConvertTo-SecureString
    return New-Object System.Management.Automation.PSCredential($Username, $securePassword)
}

#endregion Configuration Functions

#region Exchange Functions

function Connect-ExchangeSession {
    <#
    .SYNOPSIS
        Connects to Exchange PowerShell for Get-Mailbox access.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeUri,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $context = "Load Exchange CmdLets"
    
    # Check if session already exists
    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" }
    
    if ($existingSession) {
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Exchange session already exists"
        return $existingSession
    }
    
    try {
        $session = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ExchangeUri `
            -Authentication Kerberos `
            -Credential $Credential `
            -ErrorAction Stop
        
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Exchange CmdLets loaded successfully"
        
        $script:ExchangeSession = $session
        return $session
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "Error loading Exchange CmdLets: $_"
        throw
    }
}

function Initialize-EwsService {
    <#
    .SYNOPSIS
        Initializes EWS service object.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$EwsAssemblyPath,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $context = "Check EWS Managed API"
    
    if (-not (Test-Path -Path $EwsAssemblyPath)) {
        Write-Log -Context $context -Status '[ERROR]' -Message "EWS Managed API could not be found at $EwsAssemblyPath"
        throw "EWS Managed API not found at: $EwsAssemblyPath"
    }
    
    [void][Reflection.Assembly]::LoadFile($EwsAssemblyPath)
    Write-Log -Context $context -Status '[SUCCESS]' -Message "EWS Managed API loaded"
    
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential)
    
    $script:EwsService = $service
    return $service
}

function Get-UserEmail {
    <#
    .SYNOPSIS
        Resolves ITCode to email address via Get-Mailbox.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ITCode
    )
    
    $context = "Get EmailAddress"
    
    try {
        $mailbox = Get-Mailbox -Identity $ITCode -ErrorAction Stop
        $email = $mailbox.PrimarySmtpAddress.ToString()
        Write-Verbose "Resolved $ITCode to $email"
        return $email
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "Could not find email of $ITCode"
        return $null
    }
}

#endregion Exchange Functions

#region API Functions

function Test-ApiEndpoint {
    <#
    .SYNOPSIS
        Tests if API endpoint is reachable.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Uri,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter()]
        [string]$ProxyUrl
    )
    
    $context = "Test Website"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    try {
        $request = [System.Net.WebRequest]::Create($Uri)
        $request.Credentials = $Credential.GetNetworkCredential()
        
        if ($ProxyUrl) {
            $request.Proxy = [System.Net.WebProxy]::new($ProxyUrl)
        }
        
        $response = $request.GetResponse()
        $statusCode = [int]$response.StatusCode
        $response.Close()
        
        if ($statusCode -eq 200) {
            Write-Log -Context $context -Status '[SUCCESS]' -Message "Site is OK!"
            return $true
        }
        else {
            Write-Log -Context $context -Status '[ERROR]' -Message "Site returned status $statusCode"
            return $false
        }
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "The Site may be down, please check! $_"
        return $false
    }
}

function Get-LeaveDataFromApi {
    <#
    .SYNOPSIS
        Downloads leave data from API to CSV file.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Uri,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory)]
        [string]$OutputPath,
        
        [Parameter()]
        [string]$ProxyUrl
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $invokeParams = @{
        Uri        = $Uri
        Credential = $Credential
        OutFile    = $OutputPath
        Method     = 'Get'
    }
    
    if ($ProxyUrl) {
        $invokeParams.Proxy = $ProxyUrl
    }
    
    Invoke-RestMethod @invokeParams
    
    return $OutputPath
}

#endregion API Functions

#region CSV Functions

function Move-ExistingCsvFiles {
    <#
    .SYNOPSIS
        Moves existing CSV files to processed folder.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        
        [Parameter(Mandatory)]
        [string]$Pattern,
        
        [Parameter(Mandatory)]
        [string]$ProcessedPath
    )
    
    $context = "Create CSV file"
    
    if (-not (Test-Path -Path $ProcessedPath)) {
        New-Item -Path $ProcessedPath -ItemType Directory -Force | Out-Null
    }
    
    $files = Get-ChildItem -Path $ScriptPath -Filter $Pattern -ErrorAction SilentlyContinue
    
    if (-not $files) {
        Write-Log -Context $context -Status '[SUCCESS]' -Message "CSV does not exist in $ScriptPath"
        return
    }
    
    foreach ($file in $files) {
        Move-Item -Path $file.FullName -Destination $ProcessedPath -Force
        Write-Log -Context $context -Status '[WARNING]' -Message "$($file.Name) already exists in $ScriptPath file moved to $ProcessedPath"
    }
}

function Import-LeaveDataCsv {
    <#
    .SYNOPSIS
        Imports and validates CSV file.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    $context = "Import CSV"
    
    try {
        $csvData = Import-Csv -Path $Path -ErrorAction Stop
        Write-Log -Context $context -Status '[SUCCESS]' -Message "CSV file imported successfully"
        
        # Validate required fields
        $requiredFields = @('ITCode', 'StartDate', 'StartTime', 'EndDate', 'EndTime')
        
        if ($csvData.Count -gt 0) {
            foreach ($field in $requiredFields) {
                if (-not ($csvData[0].PSObject.Properties.Name -contains $field)) {
                    Write-Log -Context $context -Status '[ERROR]' -Message "Import file is missing required field: $field"
                }
            }
        }
        
        return $csvData
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "CSV file not found or invalid"
        throw
    }
}

#endregion CSV Functions

#region Calendar Functions

function New-LeaveAppointment {
    <#
    .SYNOPSIS
        Creates a leave calendar appointment.
    #>
    param(
        [Parameter(Mandatory)]
        $EwsService,
        
        [Parameter(Mandatory)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory)]
        $LeaveItem,
        
        [Parameter(Mandatory)]
        [string]$Subject,
        
        [Parameter(Mandatory)]
        [string]$Body,
        
        [Parameter()]
        [switch]$TestMode
    )
    
    $context = "Create calendar item"
    $mailboxUser = $LeaveItem.ITCode
    
    # Set impersonation
    $EwsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, 
        $EmailAddress
    )
    
    # Autodiscover
    try {
        $EwsService.AutodiscoverUrl($EmailAddress)
        Write-Log -Context "Autodiscover" -Status '[SUCCESS]' -Message "Performing autodiscover for $EmailAddress"
    }
    catch {
        Write-Log -Context "Autodiscover" -Status '[ERROR]' -Message "Autodiscover for $EmailAddress not successful"
        return $false
    }
    
    # Bind to calendar
    try {
        $calendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind(
            $EwsService, 
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
        )
        Write-Log -Context "Open user calendar" -Status '[SUCCESS]' -Message "Calendar for user $mailboxUser opened successfully"
    }
    catch {
        Write-Log -Context "Open user calendar" -Status '[ERROR]' -Message "Cannot open calendar for user $mailboxUser"
        return $false
    }
    
    # Create appointment
    try {
        $startDate = [DateTime]($LeaveItem.StartDate + " " + $LeaveItem.StartTime)
        $endDate = [DateTime]($LeaveItem.EndDate + " " + $LeaveItem.EndTime)
        
        if ($TestMode) {
            Write-Log -Context $context -Status '[INFO]' -Message "[TEST] Would create: $Subject $startDate - $endDate for $EmailAddress"
            return $true
        }
        
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($EwsService)
        $appointment.Subject = $Subject
        $appointment.Start = $startDate
        $appointment.End = $endDate
        $appointment.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::OOF
        $appointment.IsReminderSet = $false
        $appointment.Body = $Body
        
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Required fields set successfully"
        
        $appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Created $Subject $startDate $endDate for user $EmailAddress"
        
        return $true
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "Failed to create appointment: $Subject - $_"
        return $false
    }
}

function Remove-LeaveAppointment {
    <#
    .SYNOPSIS
        Removes a canceled leave calendar appointment.
    #>
    param(
        [Parameter(Mandatory)]
        $EwsService,
        
        [Parameter(Mandatory)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory)]
        $LeaveItem,
        
        [Parameter(Mandatory)]
        [string]$Subject,
        
        [Parameter()]
        [switch]$TestMode
    )
    
    $context = "Find calendar item"
    $mailboxUser = $LeaveItem.ITCode
    
    # Set impersonation
    $EwsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, 
        $EmailAddress
    )
    
    # Autodiscover
    try {
        $EwsService.AutodiscoverUrl($EmailAddress)
        Write-Log -Context "Autodiscover" -Status '[SUCCESS]' -Message "Performing autodiscover for $EmailAddress"
    }
    catch {
        Write-Log -Context "Autodiscover" -Status '[ERROR]' -Message "Autodiscover for $EmailAddress not successful"
        return $false
    }
    
    # Bind to calendar
    try {
        $calendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind(
            $EwsService, 
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
        )
        Write-Log -Context "Open user calendar" -Status '[SUCCESS]' -Message "Calendar for user $mailboxUser opened successfully"
    }
    catch {
        Write-Log -Context "Open user calendar" -Status '[ERROR]' -Message "Cannot open calendar for user $mailboxUser"
        return $false
    }
    
    # Find and remove appointment
    try {
        $startDate = [DateTime]($LeaveItem.StartDate + " " + $LeaveItem.StartTime)
        $endDate = [DateTime]($LeaveItem.EndDate + " " + $LeaveItem.EndTime)
        
        $calendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, 10)
        $appointments = $EwsService.FindAppointments($calendarFolder.Id, $calendarView)
        
        # Filter by subject and exact times
        $matchingAppointments = $appointments | Where-Object { 
            $_.Subject -eq $Subject -and 
            $_.Start -eq $startDate -and 
            $_.End -eq $endDate 
        }
        
        if (-not $matchingAppointments) {
            Write-Log -Context $context -Status '[ERROR]' -Message "Calendar Item not found"
            return $false
        }
        
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Calendar Item found. Required fields set successfully"
        
        if ($TestMode) {
            Write-Log -Context $context -Status '[INFO]' -Message "[TEST] Would delete: $Subject $startDate - $endDate for $EmailAddress"
            return $true
        }
        
        foreach ($apt in $matchingAppointments) {
            $apt.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
        }
        
        Write-Log -Context $context -Status '[SUCCESS]' -Message "Successfully deleted item $Subject $startDate $endDate"
        return $true
    }
    catch {
        Write-Log -Context $context -Status '[ERROR]' -Message "Failed to delete appointment: $Subject - $_"
        return $false
    }
}

#endregion Calendar Functions

#region Main Processing Functions

function Invoke-ImportLeaveData {
    <#
    .SYNOPSIS
        Main import process - creates leave appointments.
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter()]
        [switch]$TestMode
    )
    
    $scriptPath = $Config.Paths.ScriptPath
    $processedPath = $Config.Paths.ProcessedPath
    $leaveEndpoint = $Config.Api.LeaveDataEndpoint
    $proxyUrl = $Config.Api.ProxyUrl
    $subject = $Config.LeaveSettings.Subject
    $body = $Config.LeaveSettings.Body
    
    # Generate CSV filename
    $csvFileName = Join-Path -Path $scriptPath -ChildPath ("LeaveDataExchange" + (Get-Date -Format "ddMMyyyy") + ".csv")
    
    # Move existing CSV files
    Move-ExistingCsvFiles -ScriptPath $scriptPath -Pattern "LeaveDataExchange*" -ProcessedPath $processedPath
    
    # Test API endpoint
    if (-not (Test-ApiEndpoint -Uri $leaveEndpoint -Credential $Credential -ProxyUrl $proxyUrl)) {
        throw "API endpoint is not available"
    }
    
    # Download leave data
    if (-not $TestMode) {
        Get-LeaveDataFromApi -Uri $leaveEndpoint -Credential $Credential -OutputPath $csvFileName -ProxyUrl $proxyUrl
    }
    else {
        Write-Log -Context "Get API Data" -Status '[INFO]' -Message "[TEST] Would download from $leaveEndpoint to $csvFileName"
        return
    }
    
    # Import CSV
    $leaveData = Import-LeaveDataCsv -Path $csvFileName
    
    if (-not $leaveData -or $leaveData.Count -eq 0) {
        Write-Log -Context "Process Data" -Status '[INFO]' -Message "No leave data to process"
        return
    }
    
    Write-Host "Processing $($leaveData.Count) leave entries..." -ForegroundColor Cyan
    
    # Process each leave entry
    $successCount = 0
    $errorCount = 0
    
    foreach ($item in $leaveData) {
        $email = Get-UserEmail -ITCode $item.ITCode
        
        if (-not $email) {
            $errorCount++
            continue
        }
        
        $result = New-LeaveAppointment -EwsService $script:EwsService `
            -EmailAddress $email `
            -LeaveItem $item `
            -Subject $subject `
            -Body $body `
            -TestMode:$TestMode
        
        if ($result) {
            $successCount++
        }
        else {
            $errorCount++
        }
    }
    
    # Move processed file
    if (Test-Path -Path $csvFileName) {
        Move-Item -Path $csvFileName -Destination $processedPath -Force
        Write-Log -Context "Processed" -Status '[SUCCESS]' -Message "All items processed successfully file $csvFileName moved to $processedPath"
    }
    
    Write-Host "Import complete: Success=$successCount, Errors=$errorCount" -ForegroundColor Cyan
}

function Invoke-RemoveLeaveData {
    <#
    .SYNOPSIS
        Main remove process - deletes canceled leave appointments.
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter()]
        [switch]$TestMode
    )
    
    $scriptPath = $Config.Paths.ScriptPath
    $processedPath = $Config.Paths.ProcessedPath
    $canceledEndpoint = $Config.Api.CanceledLeaveEndpoint
    $proxyUrl = $Config.Api.ProxyUrl
    $subject = $Config.LeaveSettings.Subject
    
    # Generate CSV filename
    $csvFileName = Join-Path -Path $scriptPath -ChildPath ("CanceledLeaveDataExchange" + (Get-Date -Format "ddMMyyyy") + ".csv")
    
    # Move existing CSV files
    Move-ExistingCsvFiles -ScriptPath $scriptPath -Pattern "CanceledLeaveDataExchange*" -ProcessedPath $processedPath
    
    # Test API endpoint
    if (-not (Test-ApiEndpoint -Uri $canceledEndpoint -Credential $Credential -ProxyUrl $proxyUrl)) {
        throw "API endpoint is not available"
    }
    
    # Download canceled leave data
    if (-not $TestMode) {
        Get-LeaveDataFromApi -Uri $canceledEndpoint -Credential $Credential -OutputPath $csvFileName -ProxyUrl $proxyUrl
    }
    else {
        Write-Log -Context "Get API Data" -Status '[INFO]' -Message "[TEST] Would download from $canceledEndpoint to $csvFileName"
        return
    }
    
    # Import CSV
    $leaveData = Import-LeaveDataCsv -Path $csvFileName
    
    if (-not $leaveData -or $leaveData.Count -eq 0) {
        Write-Log -Context "Process Data" -Status '[INFO]' -Message "No canceled leave data to process"
        return
    }
    
    Write-Host "Processing $($leaveData.Count) canceled leave entries..." -ForegroundColor Cyan
    
    # Process each canceled leave entry
    $successCount = 0
    $errorCount = 0
    
    foreach ($item in $leaveData) {
        $email = Get-UserEmail -ITCode $item.ITCode
        
        if (-not $email) {
            $errorCount++
            continue
        }
        
        $result = Remove-LeaveAppointment -EwsService $script:EwsService `
            -EmailAddress $email `
            -LeaveItem $item `
            -Subject $subject `
            -TestMode:$TestMode
        
        if ($result) {
            $successCount++
        }
        else {
            $errorCount++
        }
    }
    
    # Move processed file
    if (Test-Path -Path $csvFileName) {
        Move-Item -Path $csvFileName -Destination $processedPath -Force
        Write-Log -Context "Processed" -Status '[SUCCESS]' -Message "All items processed successfully file $csvFileName moved to $processedPath"
    }
    
    Write-Host "Remove complete: Success=$successCount, Errors=$errorCount" -ForegroundColor Cyan
}

#endregion Main Processing Functions

#region Main Execution

# Load configuration
$script:Config = Import-Configuration -Path $ConfigPath

# Initialize logging
$logPrefix = switch ($Mode) {
    'Import' { 'ImportCalendarFromAfas' }
    'Remove' { 'RemoveCalendarItemFromAfas' }
    'Both'   { 'AfasLeaveData' }
}

Initialize-LogFile -LogPath $script:Config.Paths.LogPath -Prefix $logPrefix

# Start logging
$today = Get-Date
Write-Log -Context "AfasLeaveData" -Status '[START]' -Message "*** Start Logging: $today *** Version $script:Version"

if ($TestMode) {
    Write-Host "TEST MODE: No actual changes will be made" -ForegroundColor Yellow
    Write-Log -Context "AfasLeaveData" -Status '[INFO]' -Message "Running in TEST MODE"
}

try {
    # Get credentials
    if (-not $Credential) {
        $Credential = Get-CredentialFromFile `
            -Username $script:Config.Credential.Username `
            -PasswordFilePath $script:Config.Credential.PasswordFile
    }
    
    # Ensure TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    # Connect to Exchange
    Connect-ExchangeSession -ExchangeUri $script:Config.Connection.ExchangeUri -Credential $Credential
    
    # Initialize EWS
    Initialize-EwsService -EwsAssemblyPath $script:Config.EwsAssemblyPath -Credential $Credential
    
    # Execute based on mode
    switch ($Mode) {
        'Import' {
            Invoke-ImportLeaveData -Config $script:Config -Credential $Credential -TestMode:$TestMode
        }
        'Remove' {
            Invoke-RemoveLeaveData -Config $script:Config -Credential $Credential -TestMode:$TestMode
        }
        'Both' {
            Invoke-ImportLeaveData -Config $script:Config -Credential $Credential -TestMode:$TestMode
            Invoke-RemoveLeaveData -Config $script:Config -Credential $Credential -TestMode:$TestMode
        }
    }
}
catch {
    Write-Log -Context "AfasLeaveData" -Status '[ERROR]' -Message "Fatal error: $_"
    throw
}
finally {
    # Cleanup Exchange session
    if ($script:ExchangeSession) {
        Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
    }
    
    # End logging
    $today = Get-Date
    Write-Log -Context "AfasLeaveData" -Status '[STOP]' -Message "*** End Logging: $today ***"
}

#endregion Main Execution
