<#!
.SYNOPSIS
    Audits room mailbox meetings to identify "ghost" meetings with missing or disabled organizers.

.DESCRIPTION
    Connects to Exchange Server on-premises via remote PowerShell and EWS to enumerate room mailboxes,
    retrieve calendar items in a specified date window, and validate meeting organizers against Active Directory.
    Produces a report of potential ghost meetings and optionally sends notification emails to remaining attendees.
    A JSON configuration file can be supplied with -ConfigPath to pre-populate parameter values (credentials excluded).

.NOTES
    - Requires a service account with FullAccess/impersonation rights to room mailboxes for EWS queries.
    - Ensure the EWS Managed API assembly is available locally and specify -EwsAssemblyPath accordingly.
    - Run in the Exchange Management Shell or a session with Exchange/AD modules available.
    - Organizer transfer is not supported by Exchange; recreate meetings with a new organizer when needed.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()][ValidateNotNullOrEmpty()][string]$ExchangeUri = 'http://exchange.contoso.com/PowerShell/',

    [Parameter()][ValidateSet('Auto','OnPrem','EXO')][string]$ConnectionType = 'Auto',

    [Parameter()][System.Management.Automation.PSCredential]$Credential,

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    [Parameter()][ValidateRange(0,36)][int]$MonthsAhead = 12,

    [Parameter()][ValidateRange(0,12)][int]$MonthsBehind = 0,

    [Parameter()][ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'ghost-meetings-report.csv'),

    [Parameter()][string]$ExcelOutputPath = $null,

    [Parameter()][ValidateNotNullOrEmpty()][string]$OrganizationSmtpSuffix = 'contoso.com',

    [Parameter()][ValidateNotNullOrEmpty()][string]$ImpersonationSmtp,

    [Parameter()][switch]$SendInquiry,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationFrom,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationTemplate = 'Please confirm if this meeting is still required for {0}.',

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
Set-ConfigDefault -Name 'ExcelOutputPath' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$ExcelOutputPath)
Set-ConfigDefault -Name 'OrganizationSmtpSuffix' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$OrganizationSmtpSuffix)
Set-ConfigDefault -Name 'ImpersonationSmtp' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$ImpersonationSmtp)
Set-ConfigDefault -Name 'SendInquiry' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$SendInquiry) -IsSwitch
Set-ConfigDefault -Name 'NotificationFrom' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$NotificationFrom)
Set-ConfigDefault -Name 'NotificationTemplate' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$NotificationTemplate)
Set-ConfigDefault -Name 'EwsUrl' -Config $config -BoundParameters $BoundScriptParameters -Variable ([ref]$EwsUrl)

$script:ExchangeConnectionType = switch ($ConnectionType) {
    'EXO'   { 'EXO' }
    'Auto'  {
        if ($ExchangeUri -match 'outlook\.office365\.com' -or $ExchangeUri -match 'ps\.outlook\.com' -or $ExchangeUri -match 'office365\.com') { 'EXO' } else { 'OnPrem' }
    }
    default { 'OnPrem' }
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

if (-not $Credential -and $script:ExchangeConnectionType -eq 'OnPrem') {
    if ($TestMode -or $script:IsDotSourced) {
        $Credential = New-TemporaryCredential
    } else {
        $Credential = Get-Credential -Message 'Enter Exchange/AD credentials'
    }
}

if ($SendInquiry -and -not $NotificationFrom) {
    throw "-NotificationFrom is required when -SendInquiry is specified."
}

if (-not $ImpersonationSmtp) {
    if ($Credential -and $Credential.UserName -match '@') {
        $ImpersonationSmtp = $Credential.UserName
    } else {
        throw 'Provide -ImpersonationSmtp (SMTP address) for EWS Autodiscover and impersonation.'
    }
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
        Get-ExoMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity
    } else {
        Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity
    }
}

function Get-RoomMeetings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory)][ValidateNotNull()][string]$RoomSmtp,
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
                IsRecurring       = $item.AppointmentType -ne [Microsoft.Exchange.WebServices.Data.AppointmentType]::Single
                Organizer         = $item.Organizer.Address
                RequiredAttendees = $item.RequiredAttendees | ForEach-Object { $_.Address }
                OptionalAttendees = $item.OptionalAttendees | ForEach-Object { $_.Address }
                UniqueId          = $item.Id.UniqueId
                EwsItem           = $item
            }
        }

        $moreAvailable = $items.MoreAvailable
        $offset += $items.Items.Count
    }
}

function Test-OrganizerState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SmtpAddress,
        [Parameter()][ValidateNotNullOrEmpty()][string]$OrganizationSuffix
    )

    $domainMatchesOrg = $false
    if ($OrganizationSuffix) {
        $domainMatchesOrg = $SmtpAddress.ToLowerInvariant().EndsWith($OrganizationSuffix.ToLowerInvariant())
    }

    $recipient = $null
    if ($domainMatchesOrg) {
        if ($script:ExchangeConnectionType -eq 'EXO') {
            $recipient = Get-ExoRecipient -Identity $SmtpAddress -ErrorAction SilentlyContinue -PropertySets All
        } else {
            $recipient = Get-Recipient -ErrorAction SilentlyContinue -Identity $SmtpAddress
        }
    }

    if (-not $recipient) {
        return [pscustomobject]@{
            Organizer  = $SmtpAddress
            Status     = if ($domainMatchesOrg) { 'NotFound' } else { 'External' }
            Enabled    = $null
            Recipient  = $null
        }
    }

    $enabled = $null
    if ($script:ExchangeConnectionType -eq 'EXO') {
        $exoMailbox = Get-ExoMailbox -Identity $SmtpAddress -ErrorAction SilentlyContinue -PropertySets All
        if ($exoMailbox -and $exoMailbox.PSObject.Properties.Name -contains 'AccountDisabled') {
            $enabled = -not $exoMailbox.AccountDisabled
        }
    } elseif (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Verbose 'ActiveDirectory module not available; skipping enabled-state lookup.'
    } else {
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue | Out-Null
        $user = Get-ADUser -ErrorAction SilentlyContinue -Identity $recipient.SamAccountName -Properties Enabled
        if ($user) {
            $enabled = $user.Enabled
        }
    }

    $status = if ($enabled -eq $false) { 'Disabled' } else { 'Active' }

    [pscustomobject]@{
        Organizer  = $SmtpAddress
        Status     = $status
        Enabled    = $enabled
        Recipient  = $recipient.RecipientType
    }
}

function Send-GhostMeetingInquiry {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$From,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$To,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Subject,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Body
    )

    if (-not $To) { return }
    if ($PSCmdlet.ShouldProcess(($To -join ','), 'Send inquiry email')) {
        Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -BodyAsHtml
    }
}

function Find-GhostMeetings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OrganizationSuffix,
        [Parameter()][switch]$SendInquiry,
        [Parameter()][string]$NotificationFrom,
        [Parameter()][string]$NotificationTemplate,
        [Parameter()][datetime]$WindowStart,
        [Parameter()][datetime]$WindowEnd
    )

    $rooms = Get-RoomMailboxes
    $report = @()

    $roomIndex = 0
    foreach ($room in $rooms) {
        $roomIndex++
        Write-Progress -Activity 'Scanning room calendars' -Status $room.PrimarySmtpAddress -PercentComplete (($roomIndex / $rooms.Count) * 100)
        Write-Verbose "Inspecting room $($room.PrimarySmtpAddress)"
        $meetings = Get-RoomMeetings -Service $Service -RoomSmtp $room.PrimarySmtpAddress.ToString() -WindowStart $WindowStart -WindowEnd $WindowEnd

        $meetingIndex = 0
        foreach ($meeting in $meetings) {
            $meetingIndex++
            if ($meetings.Count -gt 0) {
                Write-Progress -Activity "Scanning $($room.PrimarySmtpAddress)" -Status $meeting.Subject -PercentComplete (($meetingIndex / $meetings.Count) * 100)
            }

            $organizerState = Test-OrganizerState -SmtpAddress $meeting.Organizer -OrganizationSuffix $OrganizationSuffix
            $status = $organizerState.Status
            $attendees = @($meeting.RequiredAttendees + $meeting.OptionalAttendees) | Where-Object { $_ -and $_ -ne $meeting.Organizer }

            $entry = [pscustomobject]@{
                Room             = $meeting.Room
                Subject          = $meeting.Subject
                Start            = $meeting.Start
                End              = $meeting.End
                Organizer        = $meeting.Organizer
                OrganizerStatus  = $status
                IsRecurring      = $meeting.IsRecurring
                Attendees        = $attendees -join ';'
                UniqueId         = $meeting.UniqueId
            }

            $report += $entry

            if ($status -ne 'Active' -and $SendInquiry -and $NotificationFrom -and $attendees.Count -gt 0) {
                $body = [string]::Format($NotificationTemplate, $meeting.Subject)
                Send-GhostMeetingInquiry -From $NotificationFrom -To $attendees -Subject "Room booking confirmation: $($meeting.Subject)" -Body $body
            }
        }
    }

    return $report
}

if ($script:IsDotSourced) {
    return
}

$startWindow = (Get-Date).AddMonths(-$MonthsBehind)
$endWindow = (Get-Date).AddMonths($MonthsAhead)

if ($TestMode) {
    Write-Verbose 'Test mode enabled; skipping Exchange/EWS connections and mailbox scan.'
    return
}

$credentialMissingMessage = 'A credential is required to authenticate to EWS for mailbox scans. Provide -Credential explicitly or configure a secure retrieval method.'
if (-not $Credential) {
    throw $credentialMissingMessage
}

$exchangeSession = Connect-ExchangeSession -ConnectionUri $ExchangeUri -Credential $Credential -Type $script:ExchangeConnectionType -TestMode:$TestMode

try {
    $ews = Connect-EwsService -Credential $Credential -EwsAssemblyPath $EwsAssemblyPath -ImpersonationSmtp $ImpersonationSmtp -ExplicitUrl $EwsUrl
    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -Path $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory | Out-Null
    }

    if ($ExcelOutputPath) {
        $excelDirectory = Split-Path -Path $ExcelOutputPath -Parent
        if (-not (Test-Path -Path $excelDirectory)) {
            New-Item -Path $excelDirectory -ItemType Directory | Out-Null
        }
    }
    $results = Find-GhostMeetings -Service $ews -OrganizationSuffix $OrganizationSmtpSuffix -SendInquiry:$SendInquiry -NotificationFrom $NotificationFrom -NotificationTemplate $NotificationTemplate -WindowStart $startWindow -WindowEnd $endWindow -Verbose:$VerbosePreference
    $results | Export-Csv -NoTypeInformation -Path $OutputPath
    Write-Host "Ghost meeting report saved to $OutputPath" -ForegroundColor Green

    if ($ExcelOutputPath) {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw "ImportExcel module is required to export to Excel. Install-Module ImportExcel and try again."
        }

        Import-Module ImportExcel -ErrorAction Stop
        $results | Export-Excel -Path $ExcelOutputPath -WorksheetName 'GhostMeetings' -AutoSize
        Write-Host "Ghost meeting Excel report saved to $ExcelOutputPath" -ForegroundColor Green
    }
}
finally {
    Disconnect-ExchangeSession -Type $script:ExchangeConnectionType -Session $exchangeSession
}
