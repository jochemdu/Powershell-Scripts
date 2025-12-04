#Requires -Version 5.1
<#
.SYNOPSIS
    Audits room mailbox meetings to identify "ghost" meetings with missing or disabled organizers.

.DESCRIPTION
    Connects to Exchange Server (on-premises or Online) via remote PowerShell and EWS to enumerate
    room mailboxes, retrieve calendar items in a specified date window, and validate meeting organizers
    against Active Directory or Exchange Online.

    Produces a report of potential ghost meetings and optionally sends notification emails to remaining
    attendees.

    Features:
    - Supports both Exchange On-Premises and Exchange Online
    - Parallel processing of room mailboxes (PS7+) or sequential with progress (PS5.1)
    - JSON or PSD1 configuration file support
    - CSV and Excel (XLSX) export options
    - Optional email notifications to attendees of ghost meetings

.PARAMETER ConfigPath
    Path to JSON or PSD1 configuration file. Settings from config are used as defaults.
    Command-line parameters override config values.

.PARAMETER ExchangeUri
    Exchange PowerShell endpoint URI for on-premises connections.

.PARAMETER ConnectionType
    Auto, OnPrem, or EXO. Auto detects based on ExchangeUri.

.PARAMETER Credential
    Credentials for Exchange/AD authentication. Will prompt if not provided.

.PARAMETER EwsAssemblyPath
    Path to Microsoft.Exchange.WebServices.dll.

.PARAMETER MonthsAhead
    Number of months ahead to scan for meetings. Default: 12.

.PARAMETER MonthsBehind
    Number of months behind current date to scan. Default: 0.

.PARAMETER OutputPath
    Path for CSV report output.

.PARAMETER ExcelOutputPath
    Optional path for Excel (.xlsx) report output. Requires ImportExcel module.

.PARAMETER OrganizationSmtpSuffix
    Email domain suffix to identify internal users (e.g., 'contoso.com').

.PARAMETER ImpersonationSmtp
    SMTP address for EWS impersonation. Required for Autodiscover.

.PARAMETER SendInquiry
    Send notification emails to attendees of ghost meetings.

.PARAMETER NotificationFrom
    From address for notification emails. Required when SendInquiry is used.

.PARAMETER NotificationTemplate
    Template for notification email body. Use {0} for meeting subject placeholder.

.PARAMETER EwsUrl
    Explicit EWS endpoint URL. Skips Autodiscover if provided.

.PARAMETER LocalSnapin
    Load Exchange Management Shell locally instead of using remote PowerShell.
    Use this when running on a server with Exchange Management Tools installed.
    Credentials are optional - uses current Windows identity if not provided.
    Alias: -Local

.PARAMETER ExchangeServer
    Exchange server FQDN to connect to when using LocalSnapin mode.
    If not specified, auto-discovery is used (which may fail if the auto-discovered server is unreachable).

.PARAMETER TestMode
    Run in test mode without actual Exchange/EWS connections.

.PARAMETER ThrottleLimit
    Maximum parallel room processing threads (PS7+ only). Default: CPU count.

.EXAMPLE
    .\Find-GhostRoomMeetings.ps1 -Credential (Get-Credential) -MonthsAhead 6 -Verbose

.EXAMPLE
    .\Find-GhostRoomMeetings.ps1 -ConfigPath ./config.json -SendInquiry -NotificationFrom admin@contoso.com

.EXAMPLE
    .\Find-GhostRoomMeetings.ps1 -LocalSnapin -ImpersonationSmtp admin@contoso.com -Verbose
    Runs on an Exchange server using the local snap-in with current Windows identity.

.EXAMPLE
    .\Find-GhostRoomMeetings.ps1 -LocalSnapin -Credential (Get-Credential) -ImpersonationSmtp admin@contoso.com
    Runs on an Exchange server using the local snap-in with explicit credentials for EWS.

.NOTES
    - Requires PowerShell 5.1 or later (7.0+ recommended for parallel processing)
    - Requires service account with ApplicationImpersonation rights for EWS
    - Requires EWS Managed API assembly
    - ActiveDirectory module required for on-prem organizer state lookup
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeUri = 'http://exchange.contoso.com/PowerShell/',

    [Parameter()]
    [ValidateSet('Auto', 'OnPrem', 'EXO')]
    [string]$ConnectionType = 'Auto',

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    [Parameter()]
    [ValidateRange(0, 36)]
    [int]$MonthsAhead = 12,

    [Parameter()]
    [ValidateRange(0, 12)]
    [int]$MonthsBehind = 0,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter()]
    [string]$ExcelOutputPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OrganizationSmtpSuffix = 'contoso.com',

    [Parameter()]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$ImpersonationSmtp,

    [Parameter()]
    [switch]$SendInquiry,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$NotificationFrom,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$NotificationTemplate = 'Please confirm if this meeting is still required for {0}.',

    [Parameter()]
    [string]$EwsUrl,

    [Parameter()]
    [string]$ProxyUrl,

    [Parameter()]
    [ValidateSet('Kerberos', 'Negotiate', 'Basic', 'Default')]
    [string]$Authentication = 'Kerberos',

    [Parameter()]
    [switch]$SkipCertificateCheck,

    [Parameter()]
    [Alias('Local')]
    [switch]$LocalSnapin,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeServer,

    [Parameter()]
    [switch]$TestMode,

    [Parameter()]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = [Math]::Max(1, [Environment]::ProcessorCount)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Initialization

$script:IsDotSourced = $MyInvocation.InvocationName -eq '.'
$script:ScriptRoot = $PSScriptRoot
$script:ModulePath = Join-Path -Path $script:ScriptRoot -ChildPath 'modules\ExchangeCore\ExchangeCore.psm1'

# Import shared module
if (Test-Path -Path $script:ModulePath) {
    Import-Module $script:ModulePath -Force -ErrorAction Stop
}
else {
    throw "Required module not found: $script:ModulePath"
}

#endregion Initialization

#region Configuration Loading

# Load config file if specified
$config = @{}
if ($ConfigPath) {
    $config = Import-ConfigurationFile -Path $ConfigPath

    # Support nested Connection object per AGENTS.md spec
    if ($config.ContainsKey('Connection')) {
        $conn = $config['Connection']
        if (-not $PSBoundParameters.ContainsKey('ConnectionType') -and $conn.ContainsKey('Type')) {
            $ConnectionType = $conn['Type']
        }
        if (-not $PSBoundParameters.ContainsKey('EwsUrl') -and $conn.ContainsKey('EwsUrl') -and $conn['EwsUrl']) {
            $EwsUrl = $conn['EwsUrl']
        }
        if (-not $PSBoundParameters.ContainsKey('ExchangeUri') -and $conn.ContainsKey('ExchangeUri')) {
            $ExchangeUri = $conn['ExchangeUri']
        }
        if (-not $PSBoundParameters.ContainsKey('ProxyUrl') -and $conn.ContainsKey('ProxyUrl') -and $conn['ProxyUrl']) {
            $ProxyUrl = $conn['ProxyUrl']
        }
        if (-not $PSBoundParameters.ContainsKey('Authentication') -and $conn.ContainsKey('Authentication') -and $conn['Authentication']) {
            $Authentication = $conn['Authentication']
        }
        if (-not $PSBoundParameters.ContainsKey('SkipCertificateCheck') -and $conn.ContainsKey('SkipCertificateCheck') -and $conn['SkipCertificateCheck']) {
            $SkipCertificateCheck = [switch]$true
        }
    }

    # Support nested Impersonation object per AGENTS.md spec
    if ($config.ContainsKey('Impersonation')) {
        $imp = $config['Impersonation']
        if (-not $PSBoundParameters.ContainsKey('ImpersonationSmtp') -and $imp.ContainsKey('SmtpAddress')) {
            $ImpersonationSmtp = $imp['SmtpAddress']
        }
    }

    # Apply flat config values as defaults
    $configMappings = @{
        ExchangeUri            = 'ExchangeUri'
        ConnectionType         = 'ConnectionType'
        EwsAssemblyPath        = 'EwsAssemblyPath'
        MonthsAhead            = 'MonthsAhead'
        MonthsBehind           = 'MonthsBehind'
        OutputPath             = 'OutputPath'
        ExcelOutputPath        = 'ExcelOutputPath'
        OrganizationSmtpSuffix = 'OrganizationSmtpSuffix'
        ImpersonationSmtp      = 'ImpersonationSmtp'
        NotificationFrom       = 'NotificationFrom'
        NotificationTemplate   = 'NotificationTemplate'
        EwsUrl                 = 'EwsUrl'
        ThrottleLimit          = 'ThrottleLimit'
    }

    foreach ($key in $configMappings.Keys) {
        if (-not $PSBoundParameters.ContainsKey($key) -and $config.ContainsKey($key) -and $null -ne $config[$key]) {
            Set-Variable -Name $key -Value $config[$key] -Scope Local
        }
    }

    # Handle switch parameter
    if (-not $PSBoundParameters.ContainsKey('SendInquiry') -and $config.ContainsKey('SendInquiry')) {
        $SendInquiry = [bool]$config['SendInquiry']
    }
}

# Resolve connection type
$script:ExchangeConnectionType = Get-ResolvedConnectionType -ConnectionType $ConnectionType -ExchangeUri $ExchangeUri

# Generate timestamp for output filenames
$script:ReportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# Set default output paths with timestamp if not specified
if (-not $OutputPath) {
    $OutputPath = Join-Path -Path $PWD -ChildPath "ghost-meetings-report_$($script:ReportTimestamp).csv"
}
elseif ($OutputPath -notmatch '\d{8}-\d{6}') {
    # Add timestamp to user-specified path if it doesn't already have one
    $outputDir = Split-Path -Path $OutputPath -Parent
    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
    $outputExt = [System.IO.Path]::GetExtension($OutputPath)
    if (-not $outputDir) { $outputDir = $PWD }
    $OutputPath = Join-Path -Path $outputDir -ChildPath "${outputName}_$($script:ReportTimestamp)${outputExt}"
}

if ($ExcelOutputPath -and $ExcelOutputPath -notmatch '\d{8}-\d{6}') {
    # Add timestamp to Excel path if it doesn't already have one
    $excelDir = Split-Path -Path $ExcelOutputPath -Parent
    $excelName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelOutputPath)
    $excelExt = [System.IO.Path]::GetExtension($ExcelOutputPath)
    if (-not $excelDir) { $excelDir = $PWD }
    $ExcelOutputPath = Join-Path -Path $excelDir -ChildPath "${excelName}_$($script:ReportTimestamp)${excelExt}"
}

#endregion Configuration Loading

#region Notification Function

function Send-GhostMeetingInquiry {
    <#
    .SYNOPSIS
        Sends notification email to meeting attendees.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$From,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$To,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Body
    )

    $recipientList = $To -join ', '
    if ($PSCmdlet.ShouldProcess($recipientList, 'Send inquiry email')) {
        Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -BodyAsHtml
    }
}

#endregion Notification Function

#region Main Processing

function Find-GhostMeetings {
    <#
    .SYNOPSIS
        Scans room mailboxes for ghost meetings (organizer missing/disabled).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OrganizationSuffix,

        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$ConnectionType,

        [Parameter()]
        [switch]$SendInquiry,

        [Parameter()]
        [string]$NotificationFrom,

        [Parameter()]
        [string]$NotificationTemplate,

        [Parameter(Mandatory)]
        [datetime]$WindowStart,

        [Parameter(Mandatory)]
        [datetime]$WindowEnd,

        [Parameter()]
        [int]$ThrottleLimit = 4
    )

    $rooms = Get-RoomMailboxes -ConnectionType $ConnectionType
    $roomCount = @($rooms).Count

    Write-Verbose "Found $roomCount room mailboxes to process"

    if ($roomCount -eq 0) {
        Write-Warning 'No room mailboxes found'
        return @()
    }

    # Thread-safe collection for parallel processing
    $report = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()

    # Cache organizer states to avoid redundant lookups
    $organizerCache = [System.Collections.Concurrent.ConcurrentDictionary[string, PSCustomObject]]::new()

    # Process rooms - EWS service impersonation isn't thread-safe, so we process sequentially
    # but cache organizer lookups for efficiency
    $processedRooms = 0

    foreach ($room in $rooms) {
        $processedRooms++
        $roomSmtp = $room.PrimarySmtpAddress.ToString()

        $percentComplete = [int](($processedRooms / $roomCount) * 100)
        Write-Progress -Activity 'Scanning room calendars' -Status $roomSmtp -PercentComplete $percentComplete
        Write-Verbose "[$processedRooms/$roomCount] Processing room: $roomSmtp"

        $meetings = Get-RoomCalendarItems -Service $Service -RoomSmtp $roomSmtp -WindowStart $WindowStart -WindowEnd $WindowEnd

        foreach ($meeting in $meetings) {
            if (-not $meeting.Organizer) {
                Write-Verbose "Skipping meeting without organizer: $($meeting.Subject)"
                continue
            }

            # Check cache first, add if not present
            $organizerState = $organizerCache.GetOrAdd(
                $meeting.Organizer,
                {
                    param($key)
                    Get-OrganizerState -SmtpAddress $key -OrganizationSuffix $OrganizationSuffix -ConnectionType $ConnectionType
                }
            )

            $attendees = @($meeting.RequiredAttendees + $meeting.OptionalAttendees) |
                Where-Object { $_ -and $_ -ne $meeting.Organizer }

            $entry = [PSCustomObject]@{
                Room              = $meeting.Room
                Subject           = $meeting.Subject
                Start             = $meeting.Start
                End               = $meeting.End
                Organizer         = $meeting.Organizer
                OrganizerStatus   = $organizerState.Status
                OrganizerType     = $organizerState.MailboxType
                MatchedInternal   = $organizerState.ResolvedSmtp
                IsRecurring       = $meeting.IsRecurring
                Attendees         = $attendees -join ';'
                UniqueId          = $meeting.UniqueId
            }

            $report.Add($entry)

            # Send notification for ghost meetings if requested
            $isGhost = $organizerState.Status -notin @('Active', 'External')
            if ($isGhost -and $SendInquiry -and $NotificationFrom -and $attendees.Count -gt 0) {
                $body = [string]::Format($NotificationTemplate, $meeting.Subject)
                Send-GhostMeetingInquiry -From $NotificationFrom `
                    -To $attendees `
                    -Subject "Room booking confirmation: $($meeting.Subject)" `
                    -Body $body
            }
        }
    }

    Write-Progress -Activity 'Scanning room calendars' -Completed
    return @($report)
}

#endregion Main Processing

#region Script Execution

if ($script:IsDotSourced) {
    Write-Verbose 'Script was dot-sourced; functions are now available.'
    return
}

# Calculate date window
$startWindow = (Get-Date).Date.AddMonths(-$MonthsBehind)
$endWindow = (Get-Date).Date.AddMonths($MonthsAhead)

Write-Verbose "Scanning meetings from $startWindow to $endWindow"

if ($TestMode) {
    Write-Verbose 'Test mode enabled; skipping Exchange/EWS connections and mailbox scan.'
    Write-Host "Test mode: Would scan meetings from $startWindow to $endWindow" -ForegroundColor Cyan
    return
}

# Validate required parameters - credential not needed for local snap-in
if (-not $Credential -and -not $LocalSnapin) {
    $Credential = Get-Credential -Message 'Enter Exchange/AD credentials'
}

if ($SendInquiry -and -not $NotificationFrom) {
    throw '-NotificationFrom is required when -SendInquiry is specified.'
}

# Ensure output directories exist
$outputDir = Split-Path -Path $OutputPath -Parent
if ($outputDir -and -not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

if ($ExcelOutputPath) {
    $excelDir = Split-Path -Path $ExcelOutputPath -Parent
    if ($excelDir -and -not (Test-Path -Path $excelDir)) {
        New-Item -Path $excelDir -ItemType Directory -Force | Out-Null
    }
}

$exchangeSession = $null

try {
    # Display connection parameters for debugging
    Write-Verbose "=== Connection Parameters ==="
    Write-Verbose "  ExchangeUri: $ExchangeUri"
    Write-Verbose "  ConnectionType: $script:ExchangeConnectionType"
    Write-Verbose "  Authentication: $Authentication"
    Write-Verbose "  ProxyUrl: $(if ($ProxyUrl) { $ProxyUrl } else { '(none)' })"
    Write-Verbose "  SkipCertificateCheck: $SkipCertificateCheck"
    Write-Verbose "  LocalSnapin: $LocalSnapin"
    Write-Verbose "  ExchangeServer: $(if ($ExchangeServer) { $ExchangeServer } else { '(auto-discover)' })"
    Write-Verbose "  ImpersonationSmtp: $(if ($ImpersonationSmtp) { $ImpersonationSmtp } else { '(will resolve via Get-Mailbox)' })"
    Write-Verbose "  Credential User: $(if ($Credential) { $Credential.UserName } else { '(current user)' })"
    Write-Verbose "============================="

    # Connect to Exchange
    $exchangeSession = Connect-ExchangeSession -ConnectionUri $ExchangeUri `
        -Credential $Credential `
        -Type $script:ExchangeConnectionType `
        -Authentication $Authentication `
        -ProxyUrl $ProxyUrl `
        -SkipCertificateCheck:$SkipCertificateCheck `
        -LocalSnapin:$LocalSnapin `
        -ExchangeServer $ExchangeServer `
        -TestMode:$TestMode

    Write-Verbose "Exchange session established successfully"

    # Resolve ImpersonationSmtp via Get-Mailbox if not provided
    if (-not $ImpersonationSmtp) {
        if ($Credential -and $Credential.UserName -match '@') {
            $ImpersonationSmtp = $Credential.UserName
        }
        elseif ($Credential) {
            # Extract username from DOMAIN\username format
            $username = $Credential.UserName -replace '^.*\\', ''
            try {
                $mailbox = Get-Mailbox -Identity $username -ErrorAction Stop
                $ImpersonationSmtp = $mailbox.PrimarySmtpAddress.ToString()
                Write-Verbose "Resolved ImpersonationSmtp via Get-Mailbox: $ImpersonationSmtp"
            }
            catch {
                throw "Could not resolve SMTP address for '$username' via Get-Mailbox. Provide -ImpersonationSmtp explicitly."
            }
        }
        elseif ($LocalSnapin) {
            # LocalSnapin without credential - try to resolve current user's mailbox
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $username = $currentUser -replace '^.*\\', ''
            Write-Verbose "LocalSnapin mode without credential, resolving mailbox for current user: $username"
            try {
                $mailbox = Get-Mailbox -Identity $username -ErrorAction Stop
                $ImpersonationSmtp = $mailbox.PrimarySmtpAddress.ToString()
                Write-Verbose "Resolved ImpersonationSmtp via Get-Mailbox for current user: $ImpersonationSmtp"
            }
            catch {
                throw "Could not resolve SMTP address for current user '$username' via Get-Mailbox. Provide -ImpersonationSmtp explicitly."
            }
        }
    }

    # Connect to EWS
    $ewsParams = @{
        EwsAssemblyPath = $EwsAssemblyPath
    }
    # Credential is optional for local snap-in (uses current Windows identity)
    # When no credential is passed, Connect-EwsService uses UseDefaultCredentials internally
    if ($Credential) {
        $ewsParams['Credential'] = $Credential
    }
    elseif ($LocalSnapin) {
        # Don't pass Credential - Connect-EwsService will use default Windows credentials
        Write-Verbose "Using default Windows credentials for EWS (LocalSnapin mode)"
    }
    if ($EwsUrl) {
        $ewsParams['ExplicitUrl'] = $EwsUrl
    }
    if ($ImpersonationSmtp) {
        $ewsParams['ImpersonationSmtp'] = $ImpersonationSmtp
    }
    if ($ProxyUrl) {
        $ewsParams['ProxyUrl'] = $ProxyUrl
    }
    if ($SkipCertificateCheck) {
        $ewsParams['SkipCertificateCheck'] = $true
    }

    $ews = Connect-EwsService @ewsParams

    # Find ghost meetings
    $findParams = @{
        Service              = $ews
        OrganizationSuffix   = $OrganizationSmtpSuffix
        ConnectionType       = $script:ExchangeConnectionType
        SendInquiry          = $SendInquiry
        NotificationFrom     = $NotificationFrom
        NotificationTemplate = $NotificationTemplate
        WindowStart          = $startWindow
        WindowEnd            = $endWindow
        ThrottleLimit        = $ThrottleLimit
    }

    $results = @(Find-GhostMeetings @findParams)

    # Output results
    if ($results.Count -eq 0) {
        Write-Host 'No meetings found in the specified date range.' -ForegroundColor Yellow
    }
    else {
        $ghostCount = @($results | Where-Object { $_.OrganizerStatus -notin @('Active', 'External') }).Count
        Write-Host "Found $($results.Count) meetings, $ghostCount potential ghost meetings." -ForegroundColor Cyan

        # Export CSV
        $results | Export-Csv -NoTypeInformation -Path $OutputPath -Encoding UTF8
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Green

        # Export Excel if requested
        if ($ExcelOutputPath) {
            $importExcelAvailable = Get-Module -Name ImportExcel -ListAvailable
            if (-not $importExcelAvailable) {
                Write-Warning 'ImportExcel module not installed. Skipping Excel export. Install with: Install-Module ImportExcel'
            }
            else {
                Import-Module ImportExcel -ErrorAction Stop
                $excelParams = @{
                    Path          = $ExcelOutputPath
                    WorksheetName = 'GhostMeetings'
                    AutoSize      = $true
                    FreezeTopRow  = $true
                    AutoFilter    = $true
                }
                $results | Export-Excel @excelParams
                Write-Host "Excel report saved to: $ExcelOutputPath" -ForegroundColor Green
            }
        }
    }
}
catch {
    Write-Error -ErrorRecord $_
    throw
}
finally {
    Disconnect-ExchangeSession -Type $script:ExchangeConnectionType -Session $exchangeSession
}

#endregion Script Execution
