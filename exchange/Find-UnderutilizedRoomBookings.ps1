<#!
.SYNOPSIS
    Detects room mailbox meetings that underutilize room capacity.

.DESCRIPTION
    Connects to Exchange on-premises or Exchange Online, reads room mailbox capacity,
    and audits calendar items via EWS to find meetings where the participant count is
    below a configurable threshold for rooms above a minimum capacity.
    Outputs a CSV report to highlight bookings like a room of 6 seats with only 1 attendee,
    or a room of 8 seats with 2 attendees.

.PARAMETER ConnectionType
    Determines which Exchange cmdlets to use: Auto (default), OnPrem or EXO.

.PARAMETER ExchangeUri
    Remote PowerShell URI for on-premises Exchange (ignored for EXO).

.PARAMETER Credential
    Credentials for Exchange PowerShell and EWS; prompted when omitted for OnPrem.

.PARAMETER EwsAssemblyPath
    Path to the EWS Managed API assembly used for calendar queries.

.PARAMETER MonthsAhead
    How many months ahead from today to include in the scan window.

.PARAMETER MonthsBehind
    How many months back from today to include in the scan window.

.PARAMETER ImpersonationSmtp
    SMTP address used for EWS Autodiscover and impersonation; defaults to the credential username when possible.

.PARAMETER MinimumCapacity
    Only inspect rooms whose ResourceCapacity is greater than or equal to this value.

.PARAMETER MaxParticipants
    Flag meetings where the distinct participant count (organizer + attendees) is less than or equal to this value.

.PARAMETER OutputPath
    CSV export path for the underutilized booking report.

.PARAMETER EwsUrl
    Explicit EWS URL; when omitted Autodiscover is used.

.PARAMETER ConfigPath
    Optional JSON file to pre-populate parameters (credentials excluded).

.PARAMETER TestMode
    Skips live connectivity for automated tests and fills in dummy credentials when needed.

.EXAMPLE
    pwsh -NoProfile -File ./exchange/Find-UnderutilizedRoomBookings.ps1 \ 
        -ConnectionType Auto \ 
        -ExchangeUri 'http://exchange.contoso.com/PowerShell/' \ 
        -ImpersonationSmtp 'service@contoso.com' \ 
        -MinimumCapacity 6 \ 
        -MaxParticipants 2 \ 
        -OutputPath './reports/underutilized.csv'

.NOTES
    Requires Exchange Management Shell or ExchangeOnlineManagement, plus the EWS Managed API.
    Run with an account that has ApplicationImpersonation (EXO) or equivalent impersonation rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()][ValidateNotNullOrEmpty()][string]$ExchangeUri = 'http://exchange.contoso.com/PowerShell/',

    [Parameter()][ValidateSet('Auto','OnPrem','EXO')][string]$ConnectionType = 'Auto',

    [Parameter()][System.Management.Automation.PSCredential]$Credential,

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsAssemblyPath = 'C:\\Program Files\\Microsoft\\Exchange\\Web Services\\2.2\\Microsoft.Exchange.WebServices.dll',

    [Parameter()][ValidateRange(0,36)][int]$MonthsAhead = 1,

    [Parameter()][ValidateRange(0,12)][int]$MonthsBehind = 0,

    [Parameter()][ValidateNotNullOrEmpty()][string]$ImpersonationSmtp,

    [Parameter()][ValidateRange(1,500)][int]$MinimumCapacity = 6,

    [Parameter()][ValidateRange(1,500)][int]$MaxParticipants = 2,

    [Parameter()][ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'underutilized-room-bookings.csv'),

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsUrl,

    [Parameter()][ValidateNotNullOrEmpty()][string]$ConfigPath,

    [Parameter()][switch]$TestMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:IsDotSourced = $MyInvocation.InvocationName -eq '.'

function Import-ConfigurationFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Configuration file not found at '$Path'"
    }

    $raw = Get-Content -Path $Path -Raw
    try {
        return $raw | ConvertFrom-Json
    } catch {
        throw "Configuration file '$Path' is not valid JSON: $($_.Exception.Message)"
    }
}

function Set-ConfigDefault {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name,
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][hashtable]$BoundParameters,
        [Parameter(Mandatory)][ref]$Variable,
        [Parameter()][switch]$IsSwitch
    )

    if (-not $Config) { return }
    if ($BoundParameters.ContainsKey($Name)) { return }
    if (-not ($Config.PSObject.Properties.Name -contains $Name)) { return }

    if ($IsSwitch) {
        $Variable.Value = [bool]$Config.$Name
    } else {
        $Variable.Value = $Config.$Name
    }
}

function New-TemporaryCredential {
    [CmdletBinding()]
    param(
        [Parameter()][ValidateNotNullOrEmpty()][string]$UserName = 'test.user@contoso.com'
    )

    $secureString = New-Object System.Security.SecureString
    foreach ($character in ([guid]::NewGuid().ToString('N')).ToCharArray()) {
        $secureString.AppendChar($character)
    }

    $secureString.MakeReadOnly()
    return [pscredential]::new($UserName, $secureString)
}

function Connect-ExchangeSession {
    [CmdletBinding()]
    param(
        [Parameter()][ValidateNotNullOrEmpty()][string]$ConnectionUri,
        [Parameter()][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory)][ValidateSet('OnPrem','EXO')][string]$Type,
        [Parameter()][switch]$TestMode
    )

    if ($Type -eq 'EXO') {
        if ($TestMode) {
            Write-Verbose 'Test mode enabled; skipping Connect-ExchangeOnline.'
            return $null
        }

        if (-not (Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
            throw 'ExchangeOnlineManagement module is required to connect to Exchange Online. Install-Module ExchangeOnlineManagement and try again.'
        }

        Write-Verbose 'Connecting to Exchange Online with modern authentication.'
        $connectParams = @{ ShowBanner = $false; CommandName = 'Get-ExoMailbox','Get-ExoRecipient' }

        if ($Credential) {
            $connectParams['Credential'] = $Credential
            $connectParams['UserPrincipalName'] = $Credential.UserName
        }

        Connect-ExchangeOnline @connectParams | Out-Null
        return $null
    }

    if ($TestMode) {
        Write-Verbose 'Test mode enabled; skipping on-prem Exchange session creation.'
        return $null
    }

    if (-not $Credential) {
        throw 'Provide -Credential for on-premises connections.'
    }

    Write-Verbose "Opening remote Exchange PowerShell session to $ConnectionUri"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $Credential
    Import-PSSession $session -DisableNameChecking | Out-Null
    return $session
}

function Disconnect-ExchangeSession {
    [CmdletBinding()]
    param(
        [Parameter()][ValidateSet('OnPrem','EXO')][string]$Type,
        [Parameter()][System.Management.Automation.Runspaces.PSSession]$Session
    )

    if ($Type -eq 'EXO') {
        if (Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
            Write-Verbose 'Disconnecting Exchange Online session'
            Disconnect-ExchangeOnline -Confirm:$false
        }
        return
    }

    if ($Session) {
        Write-Verbose 'Removing Exchange PowerShell session'
        Remove-PSSession $Session
    }
}

function Connect-EwsService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][System.Management.Automation.PSCredential]$Credential,
        [Parameter()][ValidateNotNullOrEmpty()][string]$EwsAssemblyPath,
        [Parameter()][ValidateNotNullOrEmpty()][string]$ImpersonationSmtp,
        [Parameter()][ValidateNotNullOrEmpty()][string]$ExplicitUrl
    )

    if (-not (Test-Path -Path $EwsAssemblyPath)) {
        throw "EWS assembly not found at '$EwsAssemblyPath'"
    }

    Add-Type -Path $EwsAssemblyPath

    $service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential.UserName, $Credential.GetNetworkCredential().Password)
    if ($PSBoundParameters.ContainsKey('ExplicitUrl')) {
        $service.Url = $ExplicitUrl
    } else {
        $service.AutodiscoverUrl($ImpersonationSmtp, { param($url) return $url -like 'https://*' })
    }

    if ($ImpersonationSmtp) {
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $ImpersonationSmtp)
    }

    return $service
}

function Get-RoomMailboxes {
    [CmdletBinding()]
    param()

    Write-Verbose 'Retrieving room mailboxes'
    if ($script:ExchangeConnectionType -eq 'EXO') {
        Get-ExoMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -PropertySets All |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity, ResourceCapacity
    } else {
        Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity, ResourceCapacity
    }
}

function Get-RoomMeetings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RoomSmtp,
        [Parameter(Mandatory)][datetime]$WindowStart,
        [Parameter(Mandatory)][datetime]$WindowEnd
    )

    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $RoomSmtp)

    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $RoomSmtp)
    $calendar = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service, $folderId)

    $view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($WindowStart, $WindowEnd, 200)
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    $moreAvailable = $true
    $offset = 0
    while ($moreAvailable) {
        $view.Offset = $offset
        $items = $calendar.FindAppointments($view)
        foreach ($item in $items.Items) {
            $item.Load()
            [pscustomobject]@{
                Room              = $RoomSmtp
                Subject           = $item.Subject
                Start             = $item.Start
                End               = $item.End
                Organizer         = $item.Organizer.Address
                RequiredAttendees = $item.RequiredAttendees | ForEach-Object { $_.Address }
                OptionalAttendees = $item.OptionalAttendees | ForEach-Object { $_.Address }
                UniqueId          = $item.Id.UniqueId
            }
        }

        $moreAvailable = $items.MoreAvailable
        $offset += $items.Items.Count
    }
}

function Get-MeetingParticipantInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()]$Meeting,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RoomSmtp
    )

    $participants = @()
    if ($Meeting.Organizer) { $participants += $Meeting.Organizer }
    if ($Meeting.RequiredAttendees) { $participants += $Meeting.RequiredAttendees }
    if ($Meeting.OptionalAttendees) { $participants += $Meeting.OptionalAttendees }

    $distinct = $participants |
        Where-Object { $_ } |
        Where-Object { $_ -ne $RoomSmtp } |
        Sort-Object -Unique

    [pscustomobject]@{
        Count        = $distinct.Count
        Participants = $distinct
    }
}

function Find-UnderutilizedMeetings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory)][ValidateRange(1,500)][int]$MinimumCapacity,
        [Parameter(Mandatory)][ValidateRange(1,500)][int]$MaxParticipants,
        [Parameter(Mandatory)][datetime]$WindowStart,
        [Parameter(Mandatory)][datetime]$WindowEnd
    )

    $rooms = Get-RoomMailboxes | Where-Object { $_.ResourceCapacity -ge $MinimumCapacity }
    $report = @()

    foreach ($room in $rooms) {
        $meetings = Get-RoomMeetings -Service $Service -RoomSmtp $room.PrimarySmtpAddress -WindowStart $WindowStart -WindowEnd $WindowEnd
        foreach ($meeting in $meetings) {
            $participantInfo = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp $room.PrimarySmtpAddress
            if ($participantInfo.Count -le $MaxParticipants) {
                $report += [pscustomobject]@{
                    Room              = $room.PrimarySmtpAddress
                    DisplayName       = $room.DisplayName
                    Capacity          = $room.ResourceCapacity
                    Subject           = $meeting.Subject
                    Start             = $meeting.Start
                    End               = $meeting.End
                    Organizer         = $meeting.Organizer
                    ParticipantCount  = $participantInfo.Count
                    Participants      = ($participantInfo.Participants -join ';')
                    UniqueId          = $meeting.UniqueId
                }
            }
        }
    }

    return $report
}

$config = $null
$script:BoundScriptParameters = $PSBoundParameters.Clone()
if ($ConfigPath) {
    $config = Import-ConfigurationFile -Path $ConfigPath
}

Set-ConfigDefault -Name 'ExchangeUri' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$ExchangeUri)
Set-ConfigDefault -Name 'ConnectionType' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$ConnectionType)
Set-ConfigDefault -Name 'EwsAssemblyPath' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$EwsAssemblyPath)
Set-ConfigDefault -Name 'MonthsAhead' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$MonthsAhead)
Set-ConfigDefault -Name 'MonthsBehind' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$MonthsBehind)
Set-ConfigDefault -Name 'OutputPath' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$OutputPath)
Set-ConfigDefault -Name 'ImpersonationSmtp' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$ImpersonationSmtp)
Set-ConfigDefault -Name 'MinimumCapacity' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$MinimumCapacity)
Set-ConfigDefault -Name 'MaxParticipants' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$MaxParticipants)
Set-ConfigDefault -Name 'EwsUrl' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$EwsUrl)
Set-ConfigDefault -Name 'TestMode' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$TestMode) -IsSwitch

$script:ExchangeConnectionType = switch ($ConnectionType) {
    'EXO'   { 'EXO' }
    'Auto'  {
        if ($ExchangeUri -match 'outlook\.office365\.com' -or $ExchangeUri -match 'ps\.outlook\.com' -or $ExchangeUri -match 'office365\.com') { 'EXO' } else { 'OnPrem' }
    }
    default { 'OnPrem' }
}

if (-not $Credential -and $script:ExchangeConnectionType -eq 'OnPrem') {
    if ($TestMode -or $script:IsDotSourced) {
        $Credential = New-TemporaryCredential
    } else {
        $Credential = Get-Credential -Message 'Enter Exchange/AD credentials'
    }
}

if (-not $ImpersonationSmtp) {
    if ($Credential -and $Credential.UserName -match '@') {
        $ImpersonationSmtp = $Credential.UserName
    } else {
        throw 'Provide -ImpersonationSmtp (SMTP address) for EWS Autodiscover and impersonation.'
    }
}

$windowStart = (Get-Date).AddMonths(-1 * $MonthsBehind)
$windowEnd = (Get-Date).AddMonths($MonthsAhead)

$session = $null
$ewsService = $null
try {
    $session = Connect-ExchangeSession -ConnectionUri $ExchangeUri -Credential $Credential -Type $script:ExchangeConnectionType -TestMode:$TestMode

    if ($TestMode) {
        $ewsService = [pscustomobject]@{ Name = 'TestEwsService' }
    } else {
        $ewsService = Connect-EwsService -Credential $Credential -EwsAssemblyPath $EwsAssemblyPath -ImpersonationSmtp $ImpersonationSmtp -ExplicitUrl $EwsUrl
    }

    $report = Find-UnderutilizedMeetings -Service $ewsService -MinimumCapacity $MinimumCapacity -MaxParticipants $MaxParticipants -WindowStart $windowStart -WindowEnd $windowEnd

    if ($PSCmdlet.ShouldProcess($OutputPath, 'Export underutilized room bookings')) {
        $report | Export-Csv -Path $OutputPath -NoTypeInformation
    }

    $report
} finally {
    Disconnect-ExchangeSession -Type $script:ExchangeConnectionType -Session $session
}
