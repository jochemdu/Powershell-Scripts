#Requires -Version 5.1
<#
.SYNOPSIS
    Detects room mailbox meetings that underutilize room capacity.

.DESCRIPTION
    Connects to Exchange (on-premises or Online) via remote PowerShell and EWS to analyze
    room mailbox calendars. Identifies meetings where the participant count is below a
    configurable threshold for rooms above a minimum capacity.

    Use cases:
    - Finding 6-seat rooms booked for 1-2 attendees
    - Identifying inefficient room utilization patterns
    - Optimizing meeting room allocation

    Features:
    - Supports both Exchange On-Premises and Exchange Online
    - JSON or PSD1 configuration file support
    - CSV export with detailed meeting information
    - Configurable capacity and participant thresholds

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
    Number of months ahead to scan for meetings. Default: 1.

.PARAMETER MonthsBehind
    Number of months behind current date to scan. Default: 0.

.PARAMETER ImpersonationSmtp
    SMTP address for EWS impersonation. Required for Autodiscover.

.PARAMETER MinimumCapacity
    Only inspect rooms with capacity >= this value. Default: 6.

.PARAMETER MaxParticipants
    Flag meetings with participant count <= this value. Default: 2.

.PARAMETER OutputPath
    Path for CSV report output.

.PARAMETER EwsUrl
    Explicit EWS endpoint URL. Skips Autodiscover if provided.

.PARAMETER LocalSnapin
    Use local Exchange Management Shell snap-in instead of remote PowerShell.
    Run this on the Exchange server directly.
    Credentials are optional - uses current Windows identity if not provided.

.PARAMETER TestMode
    Run in test mode without actual Exchange/EWS connections.

.EXAMPLE
    .\Find-UnderutilizedRoomBookings.ps1 -Credential (Get-Credential) -MinimumCapacity 6 -MaxParticipants 2

.EXAMPLE
    .\Find-UnderutilizedRoomBookings.ps1 -ConfigPath ./config.json -Verbose

.EXAMPLE
    .\Find-UnderutilizedRoomBookings.ps1 -LocalSnapin -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx

    Runs on Exchange server using local snap-in with current Windows credentials.

.NOTES
    - Requires PowerShell 5.1 or later
    - Requires service account with ApplicationImpersonation rights for EWS
    - Requires EWS Managed API assembly
#>

[CmdletBinding(SupportsShouldProcess)]
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
    [int]$MonthsAhead = 1,

    [Parameter()]
    [ValidateRange(0, 12)]
    [int]$MonthsBehind = 0,

    [Parameter()]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$ImpersonationSmtp,

    [Parameter()]
    [ValidateRange(1, 500)]
    [int]$MinimumCapacity = 6,

    [Parameter()]
    [ValidateRange(1, 500)]
    [int]$MaxParticipants = 2,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'underutilized-room-bookings.csv'),

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
    [switch]$TestMode
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
        ExchangeUri     = 'ExchangeUri'
        ConnectionType  = 'ConnectionType'
        EwsAssemblyPath = 'EwsAssemblyPath'
        MonthsAhead     = 'MonthsAhead'
        MonthsBehind    = 'MonthsBehind'
        OutputPath      = 'OutputPath'
        ImpersonationSmtp = 'ImpersonationSmtp'
        MinimumCapacity = 'MinimumCapacity'
        MaxParticipants = 'MaxParticipants'
        EwsUrl          = 'EwsUrl'
    }

    foreach ($key in $configMappings.Keys) {
        if (-not $PSBoundParameters.ContainsKey($key) -and $config.ContainsKey($key) -and $null -ne $config[$key]) {
            Set-Variable -Name $key -Value $config[$key] -Scope Local
        }
    }
}

# Resolve connection type
$script:ExchangeConnectionType = Get-ResolvedConnectionType -ConnectionType $ConnectionType -ExchangeUri $ExchangeUri

#endregion Configuration Loading

#region Room Capacity Functions

function Get-RoomMailboxesWithCapacity {
    <#
    .SYNOPSIS
        Retrieves room mailboxes with capacity information.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$ConnectionType,

        [Parameter()]
        [int]$MinimumCapacity = 1
    )

    Write-Verbose 'Retrieving room mailboxes with capacity information'

    $mailboxes = if ($ConnectionType -eq 'EXO') {
        Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -PropertySets All
    }
    else {
        Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
    }

    $mailboxes |
        Where-Object { $_.ResourceCapacity -ge $MinimumCapacity } |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity, ResourceCapacity
}

function Get-MeetingParticipantInfo {
    <#
    .SYNOPSIS
        Extracts distinct participant count from a meeting.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Meeting,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RoomSmtp
    )

    $participants = [System.Collections.Generic.List[string]]::new()

    if ($Meeting.Organizer) {
        $participants.Add($Meeting.Organizer)
    }

    if ($Meeting.RequiredAttendees) {
        foreach ($attendee in $Meeting.RequiredAttendees) {
            if ($attendee) { $participants.Add($attendee) }
        }
    }

    if ($Meeting.OptionalAttendees) {
        foreach ($attendee in $Meeting.OptionalAttendees) {
            if ($attendee) { $participants.Add($attendee) }
        }
    }

    # Get distinct participants, excluding the room itself
    $distinct = $participants |
        Where-Object { $_ -and $_ -ne $RoomSmtp } |
        Sort-Object -Unique

    [PSCustomObject]@{
        Count        = @($distinct).Count
        Participants = @($distinct)
    }
}

#endregion Room Capacity Functions

#region Main Processing

function Find-UnderutilizedMeetings {
    <#
    .SYNOPSIS
        Scans room mailboxes for underutilized meetings.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$ConnectionType,

        [Parameter(Mandatory)]
        [ValidateRange(1, 500)]
        [int]$MinimumCapacity,

        [Parameter(Mandatory)]
        [ValidateRange(1, 500)]
        [int]$MaxParticipants,

        [Parameter(Mandatory)]
        [datetime]$WindowStart,

        [Parameter(Mandatory)]
        [datetime]$WindowEnd
    )

    $rooms = Get-RoomMailboxesWithCapacity -ConnectionType $ConnectionType -MinimumCapacity $MinimumCapacity
    $roomCount = @($rooms).Count

    Write-Verbose "Found $roomCount room mailboxes with capacity >= $MinimumCapacity"

    if ($roomCount -eq 0) {
        Write-Warning "No room mailboxes found with capacity >= $MinimumCapacity"
        return @()
    }

    $report = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processedRooms = 0

    foreach ($room in $rooms) {
        $processedRooms++
        $roomSmtp = $room.PrimarySmtpAddress.ToString()

        $percentComplete = [int](($processedRooms / $roomCount) * 100)
        Write-Progress -Activity 'Scanning room calendars' -Status "$roomSmtp (Capacity: $($room.ResourceCapacity))" -PercentComplete $percentComplete
        Write-Verbose "[$processedRooms/$roomCount] Processing room: $roomSmtp (Capacity: $($room.ResourceCapacity))"

        $meetings = Get-RoomCalendarItems -Service $Service -RoomSmtp $roomSmtp -WindowStart $WindowStart -WindowEnd $WindowEnd

        foreach ($meeting in $meetings) {
            $participantInfo = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp $roomSmtp

            if ($participantInfo.Count -le $MaxParticipants) {
                $entry = [PSCustomObject]@{
                    Room             = $roomSmtp
                    DisplayName      = $room.DisplayName
                    Capacity         = $room.ResourceCapacity
                    Subject          = $meeting.Subject
                    Start            = $meeting.Start
                    End              = $meeting.End
                    Organizer        = $meeting.Organizer
                    ParticipantCount = $participantInfo.Count
                    Participants     = $participantInfo.Participants -join ';'
                    UniqueId         = $meeting.UniqueId
                }
                $report.Add($entry)
            }
        }
    }

    Write-Progress -Activity 'Scanning room calendars' -Completed
    return $report.ToArray()
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
Write-Verbose "Looking for rooms with capacity >= $MinimumCapacity and meetings with <= $MaxParticipants participants"
Write-Verbose "LocalSnapin mode: $LocalSnapin"

if ($TestMode) {
    Write-Verbose 'Test mode enabled; skipping Exchange/EWS connections.'
    Write-Host "Test mode: Would scan meetings from $startWindow to $endWindow" -ForegroundColor Cyan
    Write-Host "  Minimum room capacity: $MinimumCapacity" -ForegroundColor Cyan
    Write-Host "  Max participants threshold: $MaxParticipants" -ForegroundColor Cyan
    return
}

# Validate required parameters - credential only required for non-LocalSnapin mode
if (-not $LocalSnapin -and -not $Credential) {
    $Credential = Get-Credential -Message 'Enter Exchange/AD credentials'
}

# Ensure output directory exists
$outputDir = Split-Path -Path $OutputPath -Parent
if ($outputDir -and -not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$exchangeSession = $null

try {
    # Verbose output for connection parameters
    Write-Verbose "=== Connection Parameters ==="
    Write-Verbose "  ExchangeUri: $ExchangeUri"
    Write-Verbose "  ConnectionType: $script:ExchangeConnectionType"
    Write-Verbose "  LocalSnapin: $LocalSnapin"
    Write-Verbose "  Authentication: $Authentication"
    Write-Verbose "  ProxyUrl: $(if ($ProxyUrl) { $ProxyUrl } else { '(none)' })"
    Write-Verbose "  SkipCertificateCheck: $SkipCertificateCheck"
    Write-Verbose "  ImpersonationSmtp: $(if ($ImpersonationSmtp) { $ImpersonationSmtp } else { '(will resolve via Get-Mailbox)' })"
    Write-Verbose "  Credential User: $(if ($Credential) { $Credential.UserName } else { '(using current Windows identity)' })"
    Write-Verbose "=============================="

    # Connect to Exchange
    $exchangeSession = Connect-ExchangeSession -ConnectionUri $ExchangeUri `
        -Credential $Credential `
        -Type $script:ExchangeConnectionType `
        -Authentication $Authentication `
        -ProxyUrl $ProxyUrl `
        -SkipCertificateCheck:$SkipCertificateCheck `
        -LocalSnapin:$LocalSnapin `
        -TestMode:$TestMode

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
            # For local snap-in, try to resolve from current user
            try {
                $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $username = $currentUser -replace '^.*\\', ''
                $mailbox = Get-Mailbox -Identity $username -ErrorAction Stop
                $ImpersonationSmtp = $mailbox.PrimarySmtpAddress.ToString()
                Write-Verbose "Resolved ImpersonationSmtp via Get-Mailbox for current user: $ImpersonationSmtp"
            }
            catch {
                throw "Could not resolve SMTP address for current user. Provide -ImpersonationSmtp explicitly."
            }
        }
    }

    # Connect to EWS
    $ewsParams = @{
        EwsAssemblyPath = $EwsAssemblyPath
    }
    # Credential is optional for local snap-in (uses current Windows identity)
    if ($Credential) {
        $ewsParams['Credential'] = $Credential
    }
    elseif ($LocalSnapin) {
        # Use current Windows identity for EWS when running locally
        Write-Verbose "Using default Windows credentials for EWS (LocalSnapin mode)"
        $ewsParams['UseDefaultCredentials'] = $true
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

    $ews = Connect-EwsService @ewsParams

    # Find underutilized meetings
    $findParams = @{
        Service         = $ews
        ConnectionType  = $script:ExchangeConnectionType
        MinimumCapacity = $MinimumCapacity
        MaxParticipants = $MaxParticipants
        WindowStart     = $startWindow
        WindowEnd       = $endWindow
    }

    $results = Find-UnderutilizedMeetings @findParams

    # Output results
    if ($results.Count -eq 0) {
        Write-Host 'No underutilized meetings found in the specified date range.' -ForegroundColor Yellow
    }
    else {
        Write-Host "Found $($results.Count) underutilized room bookings." -ForegroundColor Cyan

        # Export CSV
        if ($PSCmdlet.ShouldProcess($OutputPath, 'Export underutilized room bookings')) {
            $results | Export-Csv -NoTypeInformation -Path $OutputPath -Encoding UTF8
            Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
        }

        # Return results to pipeline
        $results
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
