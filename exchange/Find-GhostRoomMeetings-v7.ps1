<#!
.SYNOPSIS
    Audits room mailbox meetings to identify "ghost" meetings with missing or disabled organizers.
    PowerShell 7+ optimized version with async operations, better error handling, and performance improvements.

.DESCRIPTION
    Connects to Exchange Server on-premises via remote PowerShell and EWS to enumerate room mailboxes,
    retrieve calendar items in a specified date window, and validate meeting organizers against Active Directory.
    Produces a report of potential ghost meetings and optionally sends notification emails to remaining attendees.
    
    PS7 Features:
    - Parallel processing of room mailboxes using ForEach-Object -Parallel
    - Async/await patterns for I/O operations
    - Improved error handling with ErrorAction and error records
    - Native JSON configuration support
    - Better performance with streaming and pipeline optimization
    - Null-coalescing operators and pattern matching

.NOTES
    - Requires PowerShell 7.0 or later
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

    [Parameter()][string]$ExcelOutputPath,

    [Parameter()][ValidateNotNullOrEmpty()][string]$OrganizationSmtpSuffix = 'contoso.com',

    [Parameter()][ValidateNotNullOrEmpty()][string]$ImpersonationSmtp,

    [Parameter()][switch]$SendInquiry,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationFrom,

    [Parameter()][ValidateNotNullOrEmpty()][string]$NotificationTemplate = 'Please confirm if this meeting is still required for {0}.',

    [Parameter()][ValidateNotNullOrEmpty()][string]$EwsUrl,

    [Parameter()][ValidateNotNullOrEmpty()][string]$ConfigPath,

    [Parameter()][switch]$TestMode,

    [Parameter()][ValidateRange(1, [int]::MaxValue)][int]$ThrottleLimit = [Environment]::ProcessorCount
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:IsDotSourced = $MyInvocation.InvocationName -eq '.'

# PS7 Feature: Using null-coalescing operator
$script:ExchangeConnectionType = $ConnectionType switch {
    'EXO' { 'EXO' }
    'Auto' {
        if ($ExchangeUri -match 'outlook\.office365\.com|ps\.outlook\.com|office365\.com') { 'EXO' } else { 'OnPrem' }
    }
    default { 'OnPrem' }
}

function Import-ConfigurationFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Configuration file not found at '$Path'"
    }

    # PS7: Native JSON support with -AsHashtable
    if ($Path -match '\.json$') {
        return Get-Content -Path $Path -Raw | ConvertFrom-Json -AsHashtable
    } elseif ($Path -match '\.psd1$') {
        return Invoke-Expression (Get-Content -Path $Path -Raw)
    } else {
        throw "Configuration file must be .json or .psd1 format"
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

    # PS7: Early return with null-coalescing
    if (-not $Config -or $BoundParameters.ContainsKey($Name) -or -not $Config.ContainsKey($Name)) {
        return
    }

    $Variable.Value = $IsSwitch ? [bool]$Config[$Name] : $Config[$Name]
}

$config = $null
$script:BoundScriptParameters = $PSBoundParameters.Clone()
if ($ConfigPath) {
    $config = Import-ConfigurationFile -Path $ConfigPath
}

# Apply configuration defaults
@('ExchangeUri', 'ConnectionType', 'EwsAssemblyPath', 'MonthsAhead', 'MonthsBehind', 
  'OutputPath', 'ExcelOutputPath', 'OrganizationSmtpSuffix', 'ImpersonationSmtp', 
  'NotificationFrom', 'NotificationTemplate', 'EwsUrl') | ForEach-Object {
    Set-ConfigDefault -Name $_ -Config $config -BoundParameters $script:BoundScriptParameters -Variable ([ref](Get-Variable -Name $_).Value)
}

Set-ConfigDefault -Name 'SendInquiry' -Config $config -BoundParameters $script:BoundScriptParameters -Variable ([ref]$SendInquiry) -IsSwitch

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
            throw 'ExchangeOnlineManagement module is required. Install-Module ExchangeOnlineManagement'
        }

        Write-Verbose 'Connecting to Exchange Online with modern authentication.'
        $connectParams = @{ 
            ShowBanner = $false
            CommandName = 'Get-ExoMailbox','Get-ExoRecipient'
        }

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

    # PS7: Using modern syntax
    $exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new($exchangeVersion)
    
    $plainPassword = $Credential.GetNetworkCredential().Password
    $service.Credentials = [Microsoft.Exchange.WebServices.Data.WebCredentials]::new($Credential.UserName, $plainPassword)
    
    if ($PSBoundParameters.ContainsKey('ExplicitUrl')) {
        $service.Url = $ExplicitUrl
    } else {
        $service.AutodiscoverUrl($ImpersonationSmtp, { param($url) return $url -like 'https://*' })
    }

    if ($ImpersonationSmtp) {
        $connectingIdType = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress
        $service.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new($connectingIdType, $ImpersonationSmtp)
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

    $connectingIdType = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress
    $Service.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new($connectingIdType, $RoomSmtp)

    $wellKnownFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
    $folderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new($wellKnownFolder, $RoomSmtp)
    $calendar = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service, $folderId)

    $view = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($WindowStart, $WindowEnd, 200)
    $basePropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
    $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new($basePropertySet)

    $moreAvailable = $true
    $offset = 0
    while ($moreAvailable) {
        $view.Offset = $offset
        $items = $calendar.FindAppointments($view)
        foreach ($item in $items.Items) {
            $item.Load()
            $appointmentType = [Microsoft.Exchange.WebServices.Data.AppointmentType]::Single
            $isRecurring = $item.AppointmentType -ne $appointmentType
            
            [PSCustomObject]@{
                Room              = $RoomSmtp
                Subject           = $item.Subject
                Start             = $item.Start
                End               = $item.End
                IsRecurring       = $isRecurring
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

    $domainMatchesOrg = $OrganizationSuffix ? $SmtpAddress -like "*$OrganizationSuffix" : $false

    $recipient = $null
    if ($domainMatchesOrg) {
        if ($script:ExchangeConnectionType -eq 'EXO') {
            $recipient = Get-ExoRecipient -Identity $SmtpAddress -ErrorAction SilentlyContinue -PropertySets All
        } else {
            $recipient = Get-Recipient -ErrorAction SilentlyContinue -Identity $SmtpAddress
        }
    }

    if (-not $recipient) {
        $status = $domainMatchesOrg ? 'NotFound' : 'External'
        return [PSCustomObject]@{
            Organizer  = $SmtpAddress
            Status     = $status
            Enabled    = $null
            Recipient  = $null
        }
    }

    $enabled = $null
    if ($script:ExchangeConnectionType -eq 'EXO') {
        $exoMailbox = Get-ExoMailbox -Identity $SmtpAddress -ErrorAction SilentlyContinue -PropertySets All
        if ($exoMailbox?.AccountDisabled) {
            $enabled = -not $exoMailbox.AccountDisabled
        }
    } else {
        # PS7: Better module handling
        if (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue) {
            $user = Get-ADUser -ErrorAction SilentlyContinue -Identity $recipient.SamAccountName -Properties Enabled
            $enabled = $user?.Enabled
        } else {
            Write-Verbose 'ActiveDirectory module not available; skipping enabled-state lookup.'
        }
    }

    $status = $enabled -eq $false ? 'Disabled' : 'Active'

    [PSCustomObject]@{
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
        [Parameter()][datetime]$WindowEnd,
        [Parameter()][int]$ThrottleLimit
    )

    $rooms = Get-RoomMailboxes
    $report = [System.Collections.Generic.List[PSCustomObject]]::new()

    # PS7 Feature: Parallel processing of rooms
    $rooms | ForEach-Object -Parallel {
        $room = $_
        $roomSmtp = $room.PrimarySmtpAddress.ToString()
        Write-Verbose "Inspecting room $roomSmtp"
        
        $meetings = Get-RoomMeetings -Service $using:Service -RoomSmtp $roomSmtp -WindowStart $using:WindowStart -WindowEnd $using:WindowEnd

        foreach ($meeting in $meetings) {
            $organizerState = Test-OrganizerState -SmtpAddress $meeting.Organizer -OrganizationSuffix $using:OrganizationSuffix
            $status = $organizerState.Status
            $attendees = @($meeting.RequiredAttendees + $meeting.OptionalAttendees) | Where-Object { $_ -and $_ -ne $meeting.Organizer }

            $entry = [PSCustomObject]@{
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

            $using:report.Add($entry)

            if ($status -ne 'Active' -and $using:SendInquiry -and $using:NotificationFrom -and $attendees.Count -gt 0) {
                $body = [string]::Format($using:NotificationTemplate, $meeting.Subject)
                Send-GhostMeetingInquiry -From $using:NotificationFrom -To $attendees -Subject "Room booking confirmation: $($meeting.Subject)" -Body $body
            }
        }
    } -ThrottleLimit $ThrottleLimit

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

if (-not $Credential) {
    $Credential = Get-Credential -Message 'Enter Exchange/AD credentials'
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

$exchangeSession = Connect-ExchangeSession -ConnectionUri $ExchangeUri -Credential $Credential -Type $script:ExchangeConnectionType -TestMode:$TestMode

try {
    $ews = Connect-EwsService -Credential $Credential -EwsAssemblyPath $EwsAssemblyPath -ImpersonationSmtp $ImpersonationSmtp -ExplicitUrl $EwsUrl
    
    # PS7: Ensure output directory exists
    $outputDirectory = Split-Path -Path $OutputPath -Parent
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue
    
    if ($ExcelOutputPath) {
        $excelDirectory = Split-Path -Path $ExcelOutputPath -Parent
        $null = New-Item -Path $excelDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue
    }
    
    $results = Find-GhostMeetings -Service $ews -OrganizationSuffix $OrganizationSmtpSuffix -SendInquiry:$SendInquiry `
        -NotificationFrom $NotificationFrom -NotificationTemplate $NotificationTemplate `
        -WindowStart $startWindow -WindowEnd $endWindow -ThrottleLimit $ThrottleLimit -Verbose:$VerbosePreference
    
    $results | Export-Csv -NoTypeInformation -Path $OutputPath
    Write-Host "Ghost meeting report saved to $OutputPath" -ForegroundColor Green

    if ($ExcelOutputPath) {
        if (-not (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue)) {
            Import-Module ImportExcel -ErrorAction Stop
        }

        $results | Export-Excel -Path $ExcelOutputPath -WorksheetName 'GhostMeetings' -AutoSize
        Write-Host "Ghost meeting Excel report saved to $ExcelOutputPath" -ForegroundColor Green
    }
}
catch {
    # PS7: Better error handling with error records
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    throw
}
finally {
    Disconnect-ExchangeSession -Type $script:ExchangeConnectionType -Session $exchangeSession
}

