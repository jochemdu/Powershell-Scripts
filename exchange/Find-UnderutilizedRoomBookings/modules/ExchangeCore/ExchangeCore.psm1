#Requires -Version 5.1
<#
.SYNOPSIS
    Shared Exchange connection and EWS functions for Exchange scripts.

.DESCRIPTION
    This module provides reusable functions for:
    - Connecting/disconnecting Exchange sessions (On-Prem and EXO)
    - Connecting to EWS
    - Retrieving room mailboxes
    - Importing configuration files

.NOTES
    Compatible with PowerShell 5.1+ and PowerShell 7+
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration

function Import-ConfigurationFile {
    <#
    .SYNOPSIS
        Imports configuration from JSON or PSD1 file.
    .PARAMETER Path
        Path to .json or .psd1 configuration file.
    .OUTPUTS
        Hashtable with configuration values.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path -Path $_ })]
        [string]$Path
    )

    switch -Regex ($Path) {
        '\.json$' {
            $content = Get-Content -Path $Path -Raw
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                return $content | ConvertFrom-Json -AsHashtable
            }
            else {
                # PS5.1: Convert PSCustomObject to hashtable
                $obj = $content | ConvertFrom-Json
                return ConvertTo-Hashtable -InputObject $obj
            }
        }
        '\.psd1$' {
            return Import-PowerShellDataFile -Path $Path
        }
        default {
            throw "Configuration file must be .json or .psd1 format. Got: $Path"
        }
    }
}

function ConvertTo-Hashtable {
    <#
    .SYNOPSIS
        Converts PSCustomObject to hashtable (for PS5.1 compatibility).
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]$InputObject
    )

    process {
        $hash = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $value = $prop.Value
            if ($value -is [PSCustomObject]) {
                $value = ConvertTo-Hashtable -InputObject $value
            }
            elseif ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                $value = @($value | ForEach-Object {
                        if ($_ -is [PSCustomObject]) { ConvertTo-Hashtable -InputObject $_ } else { $_ }
                    })
            }
            $hash[$prop.Name] = $value
        }
        return $hash
    }
}

function Get-ResolvedConnectionType {
    <#
    .SYNOPSIS
        Determines connection type (OnPrem or EXO) based on URI.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Auto', 'OnPrem', 'EXO')]
        [string]$ConnectionType,

        [Parameter()]
        [string]$ExchangeUri
    )

    if ($ConnectionType -ne 'Auto') {
        return $ConnectionType
    }

    $exoPatterns = @('outlook\.office365\.com', 'ps\.outlook\.com', 'office365\.com')
    foreach ($pattern in $exoPatterns) {
        if ($ExchangeUri -match $pattern) {
            return 'EXO'
        }
    }
    return 'OnPrem'
}

#endregion Configuration

#region Exchange Connection

function Connect-ExchangeSession {
    <#
    .SYNOPSIS
        Connects to Exchange (On-Premises or Online).
    .PARAMETER ConnectionUri
        Exchange PowerShell endpoint URI for on-prem.
    .PARAMETER Credential
        Credentials for authentication.
    .PARAMETER Type
        OnPrem or EXO connection type.
    .PARAMETER TestMode
        Skip actual connection for testing.
    .OUTPUTS
        PSSession for OnPrem, $null for EXO.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.Runspaces.PSSession])]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ConnectionUri,

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$Type,

        [Parameter()]
        [switch]$TestMode
    )

    if ($TestMode) {
        Write-Verbose "Test mode enabled; skipping $Type connection."
        return $null
    }

    if ($Type -eq 'EXO') {
        $connectCmd = Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue
        if (-not $connectCmd) {
            throw 'ExchangeOnlineManagement module required. Install with: Install-Module ExchangeOnlineManagement'
        }

        Write-Verbose 'Connecting to Exchange Online with modern authentication.'
        $connectParams = @{
            ShowBanner  = $false
            CommandName = @('Get-EXOMailbox', 'Get-EXORecipient')
        }

        if ($Credential) {
            $connectParams['Credential'] = $Credential
            $connectParams['UserPrincipalName'] = $Credential.UserName
        }

        Connect-ExchangeOnline @connectParams | Out-Null
        return $null
    }

    # OnPrem connection
    if (-not $Credential) {
        throw 'Credential required for on-premises Exchange connections.'
    }

    Write-Verbose "Opening remote Exchange PowerShell session to $ConnectionUri"
    $sessionParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $ConnectionUri
        Authentication    = 'Kerberos'
        Credential        = $Credential
    }

    $session = New-PSSession @sessionParams
    Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
    return $session
}

function Disconnect-ExchangeSession {
    <#
    .SYNOPSIS
        Disconnects Exchange session.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$Type,

        [Parameter()]
        [System.Management.Automation.Runspaces.PSSession]$Session
    )

    if ($Type -eq 'EXO') {
        $disconnectCmd = Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue
        if ($disconnectCmd) {
            Write-Verbose 'Disconnecting Exchange Online session'
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        return
    }

    if ($Session -and $Session.State -eq 'Opened') {
        Write-Verbose 'Removing Exchange PowerShell session'
        Remove-PSSession $Session -ErrorAction SilentlyContinue
    }
}

#endregion Exchange Connection

#region EWS Functions

function Connect-EwsService {
    <#
    .SYNOPSIS
        Creates and configures EWS ExchangeService object.
    .PARAMETER Credential
        Credentials for EWS authentication.
    .PARAMETER EwsAssemblyPath
        Path to Microsoft.Exchange.WebServices.dll.
    .PARAMETER ImpersonationSmtp
        SMTP address for impersonation and autodiscover.
    .PARAMETER ExplicitUrl
        Explicit EWS endpoint URL (skips autodiscover).
    .OUTPUTS
        Microsoft.Exchange.WebServices.Data.ExchangeService
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path -Path $_ })]
        [string]$EwsAssemblyPath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ImpersonationSmtp,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ExplicitUrl
    )

    if (-not $ExplicitUrl -and -not $ImpersonationSmtp) {
        throw 'Either -ExplicitUrl or -ImpersonationSmtp required for EWS connection.'
    }

    Add-Type -Path $EwsAssemblyPath

    $exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

    # PS5.1 compatible syntax
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new($exchangeVersion)
    }
    else {
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchangeVersion)
    }

    $networkCred = $Credential.GetNetworkCredential()
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $service.Credentials = [Microsoft.Exchange.WebServices.Data.WebCredentials]::new(
            $networkCred.UserName,
            $networkCred.Password,
            $networkCred.Domain
        )
    }
    else {
        $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(
            $networkCred.UserName,
            $networkCred.Password,
            $networkCred.Domain
        )
    }

    if ($ExplicitUrl) {
        $service.Url = [Uri]$ExplicitUrl
    }
    else {
        $redirectCallback = { param($url) return $url -like 'https://*' }
        $service.AutodiscoverUrl($ImpersonationSmtp, $redirectCallback)
    }

    if ($ImpersonationSmtp) {
        $connectingIdType = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $service.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new(
                $connectingIdType,
                $ImpersonationSmtp
            )
        }
        else {
            $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
                $connectingIdType,
                $ImpersonationSmtp
            )
        }
    }

    return $service
}

function Get-RoomMailboxes {
    <#
    .SYNOPSIS
        Retrieves all room mailboxes from Exchange.
    .PARAMETER ConnectionType
        OnPrem or EXO to determine cmdlet.
    .OUTPUTS
        Array of room mailbox objects.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$ConnectionType
    )

    Write-Verbose 'Retrieving room mailboxes'

    $mailboxes = if ($ConnectionType -eq 'EXO') {
        Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
    }
    else {
        Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
    }

    return $mailboxes | Select-Object DisplayName, PrimarySmtpAddress, Alias, Identity
}

function Get-RoomCalendarItems {
    <#
    .SYNOPSIS
        Retrieves calendar items from a room mailbox.
    .PARAMETER Service
        EWS ExchangeService object.
    .PARAMETER RoomSmtp
        SMTP address of the room mailbox.
    .PARAMETER WindowStart
        Start of date range.
    .PARAMETER WindowEnd
        End of date range.
    .OUTPUTS
        Array of meeting objects.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RoomSmtp,

        [Parameter(Mandatory)]
        [datetime]$WindowStart,

        [Parameter(Mandatory)]
        [datetime]$WindowEnd
    )

    # Set impersonation for this room
    $connectingIdType = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $Service.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new($connectingIdType, $RoomSmtp)
    }
    else {
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId($connectingIdType, $RoomSmtp)
    }

    $wellKnownFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $folderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new($wellKnownFolder, $RoomSmtp)
    }
    else {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($wellKnownFolder, $RoomSmtp)
    }

    try {
        $calendar = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service, $folderId)
    }
    catch {
        Write-Warning "Failed to access calendar for room '$RoomSmtp': $_"
        return @()
    }

    $pageSize = 200
    $basePropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $view = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($WindowStart, $WindowEnd, $pageSize)
        $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new($basePropertySet)
    }
    else {
        $view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($WindowStart, $WindowEnd, $pageSize)
        $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet($basePropertySet)
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $offset = 0

    do {
        $view.Offset = $offset
        $appointments = $calendar.FindAppointments($view)

        foreach ($item in $appointments.Items) {
            try {
                $item.Load()
                $singleType = [Microsoft.Exchange.WebServices.Data.AppointmentType]::Single

                $meeting = [PSCustomObject]@{
                    Room              = $RoomSmtp
                    Subject           = $item.Subject
                    Start             = $item.Start
                    End               = $item.End
                    IsRecurring       = $item.AppointmentType -ne $singleType
                    Organizer         = if ($item.Organizer) { $item.Organizer.Address } else { $null }
                    RequiredAttendees = @($item.RequiredAttendees | ForEach-Object { $_.Address } | Where-Object { $_ })
                    OptionalAttendees = @($item.OptionalAttendees | ForEach-Object { $_.Address } | Where-Object { $_ })
                    UniqueId          = $item.Id.UniqueId
                }
                $results.Add($meeting)
            }
            catch {
                Write-Warning "Failed to load appointment in room '$RoomSmtp': $_"
            }
        }

        $offset += $appointments.Items.Count
    } while ($appointments.MoreAvailable)

    return $results.ToArray()
}

#endregion EWS Functions

#region Organizer Validation

function Get-OrganizerState {
    <#
    .SYNOPSIS
        Checks the state of a meeting organizer (Active, Disabled, NotFound, External).
    .PARAMETER SmtpAddress
        Organizer's SMTP address.
    .PARAMETER OrganizationSuffix
        Organization domain suffix (e.g., 'contoso.com').
    .PARAMETER ConnectionType
        OnPrem or EXO connection type.
    .OUTPUTS
        PSCustomObject with Organizer, Status, Enabled, Recipient properties.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SmtpAddress,

        [Parameter()]
        [string]$OrganizationSuffix,

        [Parameter(Mandatory)]
        [ValidateSet('OnPrem', 'EXO')]
        [string]$ConnectionType
    )

    $isInternal = $OrganizationSuffix -and ($SmtpAddress -like "*@$OrganizationSuffix")

    if (-not $isInternal) {
        return [PSCustomObject]@{
            Organizer = $SmtpAddress
            Status    = 'External'
            Enabled   = $null
            Recipient = $null
        }
    }

    # Lookup recipient
    $recipient = if ($ConnectionType -eq 'EXO') {
        Get-EXORecipient -Identity $SmtpAddress -ErrorAction SilentlyContinue
    }
    else {
        Get-Recipient -Identity $SmtpAddress -ErrorAction SilentlyContinue
    }

    if (-not $recipient) {
        return [PSCustomObject]@{
            Organizer = $SmtpAddress
            Status    = 'NotFound'
            Enabled   = $null
            Recipient = $null
        }
    }

    # Check enabled state
    $enabled = $null
    if ($ConnectionType -eq 'EXO') {
        $exoMailbox = Get-EXOMailbox -Identity $SmtpAddress -ErrorAction SilentlyContinue
        if ($exoMailbox -and ($exoMailbox | Get-Member -Name AccountDisabled -ErrorAction SilentlyContinue)) {
            $enabled = -not $exoMailbox.AccountDisabled
        }
    }
    else {
        # Try ActiveDirectory module for on-prem
        $adLoaded = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue
        if (-not $adLoaded) {
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue | Out-Null
            $adLoaded = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue
        }

        if ($adLoaded -and $recipient.SamAccountName) {
            $adUser = Get-ADUser -Identity $recipient.SamAccountName -Properties Enabled -ErrorAction SilentlyContinue
            if ($adUser) {
                $enabled = $adUser.Enabled
            }
        }
        else {
            Write-Verbose 'ActiveDirectory module not available; skipping enabled-state lookup.'
        }
    }

    $status = if ($enabled -eq $false) { 'Disabled' } else { 'Active' }

    return [PSCustomObject]@{
        Organizer = $SmtpAddress
        Status    = $status
        Enabled   = $enabled
        Recipient = $recipient.RecipientType
    }
}

#endregion Organizer Validation

# Export module members
Export-ModuleMember -Function @(
    'Import-ConfigurationFile'
    'ConvertTo-Hashtable'
    'Get-ResolvedConnectionType'
    'Connect-ExchangeSession'
    'Disconnect-ExchangeSession'
    'Connect-EwsService'
    'Get-RoomMailboxes'
    'Get-RoomCalendarItems'
    'Get-OrganizerState'
)
