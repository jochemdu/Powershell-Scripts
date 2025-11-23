<#!
.SYNOPSIS
    Audits room mailbox meetings to identify "ghost" meetings with missing or disabled organizers.

.DESCRIPTION
    Connects to Exchange Server on-premises via remote PowerShell and EWS to enumerate room mailboxes,
    retrieve calendar items in a specified date window, and validate meeting organizers against Active Directory.
    Produces a report of potential ghost meetings and optionally sends notification emails to remaining attendees.

.NOTES
    - Requires a service account with FullAccess/impersonation rights to room mailboxes for EWS queries.
    - Ensure the EWS Managed API assembly is available locally and specify -EwsAssemblyPath accordingly.
    - Run in the Exchange Management Shell or a session with Exchange/AD modules available.
    - Organizer transfer is not supported by Exchange; recreate meetings with a new organizer when needed.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()][ValidateNotNullOrEmpty()][string]$ExchangeUri = 'http://exchange.contoso.com/PowerShell/',

    [Parameter()][ValidateNotNull()][System.Management.Automation.PSCredential]$Credential = (Get-Credential -Message 'Enter Exchange/AD credentials'),

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    [Parameter()][ValidateRange(0,36)][int]$MonthsAhead = 12,

    [Parameter()][ValidateRange(0,12)][int]$MonthsBehind = 0,

    [Parameter()][ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'ghost-meetings-report.csv'),

    [Parameter()][ValidateNotNullOrEmpty()][string]$OrganizationSmtpSuffix = 'contoso.com',

    [Parameter()][ValidateNotNullOrEmpty()][string]$ImpersonationSmtp,

    [Parameter()][switch]$SendInquiry,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationFrom,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationTemplate = 'Please confirm if this meeting is still required for {0}.',

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($SendInquiry -and -not $NotificationFrom) {
    throw "-NotificationFrom is required when -SendInquiry is specified."
}

if (-not $ImpersonationSmtp) {
    if ($Credential.UserName -match '@') {
        $ImpersonationSmtp = $Credential.UserName
    } else {
        throw 'Provide -ImpersonationSmtp (SMTP address) for EWS Autodiscover and impersonation.'
    }
}

function Connect-ExchangeSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ConnectionUri,
        [Parameter(Mandatory)][ValidateNotNull()][System.Management.Automation.PSCredential]$Credential
    )

    Write-Verbose "Opening remote Exchange PowerShell session to $ConnectionUri"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $Credential
    Import-PSSession $session -DisableNameChecking | Out-Null
    return $session
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
    Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity
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
        $recipient = Get-Recipient -ErrorAction SilentlyContinue -Identity $SmtpAddress
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
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
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

    foreach ($room in $rooms) {
        Write-Verbose "Inspecting room $($room.PrimarySmtpAddress)"
        $meetings = Get-RoomMeetings -Service $Service -RoomSmtp $room.PrimarySmtpAddress.ToString() -WindowStart $WindowStart -WindowEnd $WindowEnd

        foreach ($meeting in $meetings) {
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

            if ($status -ne 'Active' -and $SendInquiry -and $NotificationFrom) {
                $body = [string]::Format($NotificationTemplate, $meeting.Subject)
                Send-GhostMeetingInquiry -From $NotificationFrom -To $attendees -Subject "Room booking confirmation: $($meeting.Subject)" -Body $body
            }
        }
    }

    return $report
}

$startWindow = (Get-Date).AddMonths(-$MonthsBehind)
$endWindow = (Get-Date).AddMonths($MonthsAhead)

$exchangeSession = Connect-ExchangeSession -ConnectionUri $ExchangeUri -Credential $Credential

try {
    $ews = Connect-EwsService -Credential $Credential -EwsAssemblyPath $EwsAssemblyPath -ImpersonationSmtp $ImpersonationSmtp -ExplicitUrl $EwsUrl
    $results = Find-GhostMeetings -Service $ews -OrganizationSuffix $OrganizationSmtpSuffix -SendInquiry:$SendInquiry -NotificationFrom $NotificationFrom -NotificationTemplate $NotificationTemplate -WindowStart $startWindow -WindowEnd $endWindow -Verbose:$VerbosePreference
    $results | Export-Csv -NoTypeInformation -Path $OutputPath
    Write-Host "Ghost meeting report saved to $OutputPath" -ForegroundColor Green
}
finally {
    if ($exchangeSession) {
        Write-Verbose 'Removing Exchange PowerShell session'
        Remove-PSSession $exchangeSession
    }
}
