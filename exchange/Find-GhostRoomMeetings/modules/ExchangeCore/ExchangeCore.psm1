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
        Exchange PowerShell endpoint URI for on-prem. Use 'Local' to load Exchange snap-in locally.
    .PARAMETER Credential
        Credentials for authentication.
    .PARAMETER Type
        OnPrem or EXO connection type.
    .PARAMETER Authentication
        Authentication method: Kerberos, Negotiate, Basic, Default. Default is Kerberos.
    .PARAMETER ProxyUrl
        Proxy server URL (e.g., http://proxy.contoso.com:8080).
    .PARAMETER SkipCertificateCheck
        Skip SSL certificate validation (for self-signed or mismatched certs).
    .PARAMETER TestMode
        Skip actual connection for testing.
    .PARAMETER LocalSnapin
        Load Exchange Management Shell locally instead of remote PowerShell.
        Uses RemoteExchange.ps1 to connect.
    .PARAMETER ExchangeServer
        Exchange server FQDN to connect to when using LocalSnapin mode.
        If not specified, uses -auto (lets Exchange choose a server).
    .OUTPUTS
        PSSession for OnPrem remote, $null for EXO or local snap-in.
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
        [ValidateSet('Kerberos', 'Negotiate', 'Basic', 'Default')]
        [string]$Authentication = 'Kerberos',

        [Parameter()]
        [string]$ProxyUrl,

        [Parameter()]
        [switch]$SkipCertificateCheck,

        [Parameter()]
        [switch]$TestMode,

        [Parameter()]
        [switch]$LocalSnapin,

        [Parameter()]
        [string]$ExchangeServer
    )

    if ($TestMode) {
        Write-Verbose "Test mode enabled; skipping $Type connection."
        return $null
    }

    # Handle local Exchange snap-in loading (for running on Exchange server)
    if ($LocalSnapin -or $ConnectionUri -eq 'Local') {
        Write-Verbose "=== LocalSnapin Mode ==="
        Write-Verbose "  Loading Exchange Management Shell locally..."
        Write-Verbose "  PowerShell Version: $($PSVersionTable.PSVersion)"
        Write-Verbose "  Current User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        Write-Verbose "========================="
        
        # Check if Exchange commands are already available
        $existingCmd = Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue
        if ($existingCmd) {
            Write-Verbose "Exchange commands already available (Get-Mailbox found)"
            return $null
        }
        
        # Snap-ins only work in Windows PowerShell 5.1, not in PowerShell 7+
        # PSEdition is 'Desktop' for Windows PowerShell, 'Core' for PS6+
        # On older PS versions, PSEdition may not exist, so default to Desktop behavior
        $psEdition = $PSVersionTable.PSEdition
        $isWindowsPowerShell = ($null -eq $psEdition) -or ($psEdition -eq 'Desktop')
        
        Write-Verbose "  PSEdition: $(if ($psEdition) { $psEdition } else { '(not set - assuming Desktop)' })"
        Write-Verbose "  IsWindowsPowerShell: $isWindowsPowerShell"
        
        if ($isWindowsPowerShell) {
            Write-Verbose "Windows PowerShell detected, trying snap-ins..."
            
            # Try Exchange 2016/2019 snap-in
            $snapinName = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
            $snapinLoaded = Get-PSSnapin -Name $snapinName -ErrorAction SilentlyContinue
            
            if ($snapinLoaded) {
                Write-Verbose "Exchange snap-in already loaded"
                return $null
            }
            
            $snapinAvailable = Get-PSSnapin -Registered -Name $snapinName -ErrorAction SilentlyContinue
            if ($snapinAvailable) {
                Write-Verbose "Adding Exchange snap-in: $snapinName"
                Add-PSSnapin $snapinName -ErrorAction Stop
                Write-Verbose "Exchange snap-in loaded successfully"
                return $null
            }
            
            # Try Exchange 2013/2010 snap-in
            $snapinName2013 = 'Microsoft.Exchange.Management.PowerShell.E2010'
            $snapinLoaded2013 = Get-PSSnapin -Name $snapinName2013 -ErrorAction SilentlyContinue
            
            if ($snapinLoaded2013) {
                Write-Verbose "Exchange 2013 snap-in already loaded"
                return $null
            }
            
            $snapinAvailable2013 = Get-PSSnapin -Registered -Name $snapinName2013 -ErrorAction SilentlyContinue
            if ($snapinAvailable2013) {
                Write-Verbose "Adding Exchange 2013 snap-in: $snapinName2013"
                Add-PSSnapin $snapinName2013 -ErrorAction Stop
                Write-Verbose "Exchange 2013 snap-in loaded successfully"
                return $null
            }
            
            Write-Verbose "No snap-ins available, trying RemoteExchange.ps1..."
        }
        else {
            Write-Verbose "PowerShell Core/7+ detected - snap-ins not supported, using RemoteExchange.ps1"
        }
        
        # Try RemoteExchange.ps1 script (works in both PS5.1 and PS7)
        # Prioritize common hardcoded paths first, then env var
        $remoteExchangePaths = @(
            'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1',
            'D:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1',
            'E:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1',
            'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
        )
        
        # Add env var path if it's set
        if ($env:ExchangeInstallPath) {
            $envPath = Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1'
            $remoteExchangePaths = @($envPath) + $remoteExchangePaths
        }
        
        Write-Verbose "Searching for RemoteExchange.ps1 in $($remoteExchangePaths.Count) locations..."
        
        foreach ($path in $remoteExchangePaths) {
            Write-Verbose "  Checking: $path"
            if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
                Write-Verbose "Loading Exchange via RemoteExchange.ps1: $path"
                
                # Temporarily disable StrictMode - RemoteExchange.ps1 uses uninitialized variables
                Set-StrictMode -Off
                try {
                    . $path
                    
                    # Connect to specific server or auto-discover
                    if ($ExchangeServer) {
                        Write-Verbose "Connecting to specified Exchange server: $ExchangeServer"
                        $connectParams = @{
                            ServerFqdn = $ExchangeServer
                            ClientApplication = 'ManagementShell'
                        }
                        if ($Credential) {
                            $connectParams['Credential'] = $Credential
                        }
                        Connect-ExchangeServer @connectParams
                    }
                    else {
                        Write-Verbose "Using auto-discovery to find Exchange server"
                        Connect-ExchangeServer -auto -ClientApplication:ManagementShell
                    }
                }
                finally {
                    # Re-enable StrictMode
                    Set-StrictMode -Version Latest
                }
                Write-Verbose "Exchange Management Shell loaded via RemoteExchange.ps1"
                
                # Verify commands are available
                $testCmd = Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue
                if ($testCmd) {
                    Write-Verbose "Exchange commands verified (Get-Mailbox available)"
                    return $null
                }
            }
        }
        
        throw "Exchange Management Shell not found. Ensure you are running on an Exchange server with management tools installed. For PowerShell 7+, RemoteExchange.ps1 must be available."
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

    Write-Verbose "=== Exchange Session Details ==="
    Write-Verbose "  ConnectionUri: $ConnectionUri"
    Write-Verbose "  Authentication: $Authentication"
    Write-Verbose "  Credential User: $($Credential.UserName)"
    Write-Verbose "  SkipCertificateCheck: $SkipCertificateCheck"
    Write-Verbose "  ProxyUrl: $(if ($ProxyUrl) { $ProxyUrl } else { '(none)' })"
    Write-Verbose "================================="
    
    # Pre-flight connectivity check
    try {
        $uri = [Uri]$ConnectionUri
        $testHost = $uri.Host
        $testPort = if ($uri.Port -gt 0) { $uri.Port } elseif ($uri.Scheme -eq 'https') { 443 } else { 80 }
        
        Write-Verbose "Pre-flight check: Testing TCP connection to ${testHost}:${testPort}..."
        $tcpTest = Test-NetConnection -ComputerName $testHost -Port $testPort -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        if ($tcpTest.TcpTestSucceeded) {
            Write-Verbose "Pre-flight check: TCP connection to ${testHost}:${testPort} succeeded"
        }
        else {
            Write-Warning "Pre-flight check: TCP connection to ${testHost}:${testPort} FAILED - check firewall/network"
        }
        
        # Test WinRM port (5985 for HTTP, 5986 for HTTPS)
        $winrmPort = if ($uri.Scheme -eq 'https') { 5986 } else { 5985 }
        Write-Verbose "Pre-flight check: Testing WinRM port ${testHost}:${winrmPort}..."
        $winrmTest = Test-NetConnection -ComputerName $testHost -Port $winrmPort -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        if ($winrmTest.TcpTestSucceeded) {
            Write-Verbose "Pre-flight check: WinRM port ${testHost}:${winrmPort} is open"
        }
        else {
            Write-Warning "Pre-flight check: WinRM port ${testHost}:${winrmPort} CLOSED - WinRM may not be enabled on server"
        }
    }
    catch {
        Write-Verbose "Pre-flight check: Could not test connectivity - $($_.Exception.Message)"
    }
    
    # Build session options
    $sessionOptionParams = @{}
    if ($SkipCertificateCheck) {
        Write-Verbose 'Adding session option: SkipCACheck, SkipCNCheck, SkipRevocationCheck'
        $sessionOptionParams['SkipCACheck'] = $true
        $sessionOptionParams['SkipCNCheck'] = $true
        $sessionOptionParams['SkipRevocationCheck'] = $true
    }
    
    # Proxy is only supported with HTTPS transport - WinRM limitation
    $isHttps = $ConnectionUri -like 'https://*'
    if ($ProxyUrl -and $isHttps) {
        # Normalize proxy URL
        $normalizedProxy = $ProxyUrl
        if ($normalizedProxy -notmatch '^https?://') {
            $normalizedProxy = "http://$normalizedProxy"
        }
        Write-Verbose "Adding session option: ProxyAccessType = AutoDetect with proxy $normalizedProxy"
        $sessionOptionParams['ProxyAccessType'] = 'NoProxyServer'  # We'll handle proxy differently
        
        # For WinRM over HTTPS with proxy, we need to configure WinRM client settings
        # The PSSessionOption proxy settings are limited, so we try a direct approach
        Write-Verbose "Note: WinRM proxy requires system-level proxy configuration or winhttp proxy settings"
        Write-Verbose "If connection fails, try running: netsh winhttp set proxy $ProxyUrl"
    }
    elseif ($ProxyUrl -and -not $isHttps) {
        Write-Warning "Proxy configuration ignored: WinRM only supports proxy with HTTPS transport. Current URI uses HTTP."
    }
    $sessionOptions = New-PSSessionOption @sessionOptionParams
    
    $sessionParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $ConnectionUri
        Authentication    = $Authentication
        Credential        = $Credential
        SessionOption     = $sessionOptions
    }

    Write-Verbose "Creating PSSession with ConfigurationName: Microsoft.Exchange"
    Write-Verbose "Attempting connection..."
    
    try {
        $session = New-PSSession @sessionParams
        Write-Verbose "PSSession created successfully. Session ID: $($session.Id), State: $($session.State)"
        
        Write-Verbose "Importing PSSession commands..."
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        Write-Verbose "PSSession commands imported successfully"
        
        return $session
    }
    catch {
        Write-Verbose "ERROR creating PSSession!"
        Write-Verbose "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Verbose "Exception Message: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        if ($_.ErrorDetails) {
            Write-Verbose "Error Details: $($_.ErrorDetails)"
        }
        Write-Verbose "Full Error Record: $_"
        Write-Verbose "Category Info: $($_.CategoryInfo)"
        Write-Verbose "Fully Qualified Error ID: $($_.FullyQualifiedErrorId)"
        
        $innerMsg = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { '(none)' }
        $detailsMsg = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { '(none)' }
        Write-Error "Failed to connect to Exchange: $($_.Exception.Message) | Inner: $innerMsg | Details: $detailsMsg"
        throw
    }
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
        Credentials for EWS authentication. Use [PSCredential]::Empty for default Windows credentials.
    .PARAMETER EwsAssemblyPath
        Path to Microsoft.Exchange.WebServices.dll.
    .PARAMETER ImpersonationSmtp
        SMTP address for impersonation and autodiscover.
    .PARAMETER ExplicitUrl
        Explicit EWS endpoint URL (skips autodiscover).
    .PARAMETER ProxyUrl
        Proxy server URL (e.g., http://proxy.contoso.com:8080).
    .PARAMETER SkipCertificateCheck
        Skip SSL certificate validation for EWS connections (for self-signed or untrusted certs).
    .OUTPUTS
        Microsoft.Exchange.WebServices.Data.ExchangeService
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
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
        [string]$ExplicitUrl,

        [Parameter()]
        [string]$ProxyUrl,

        [Parameter()]
        [switch]$SkipCertificateCheck
    )

    if (-not $ExplicitUrl -and -not $ImpersonationSmtp) {
        throw 'Either -ExplicitUrl or -ImpersonationSmtp required for EWS connection.'
    }

    # Check if using default credentials
    $useDefaultCredentials = (-not $Credential) -or ($Credential -eq [System.Management.Automation.PSCredential]::Empty)

    Write-Verbose "=== EWS Connection Details ==="
    Write-Verbose "  EwsAssemblyPath: $EwsAssemblyPath"
    Write-Verbose "  ExplicitUrl: $(if ($ExplicitUrl) { $ExplicitUrl } else { '(using Autodiscover)' })"
    Write-Verbose "  ImpersonationSmtp: $(if ($ImpersonationSmtp) { $ImpersonationSmtp } else { '(none)' })"
    Write-Verbose "  ProxyUrl: $(if ($ProxyUrl) { $ProxyUrl } else { '(none)' })"
    Write-Verbose "  SkipCertificateCheck: $SkipCertificateCheck"
    Write-Verbose "  UseDefaultCredentials: $useDefaultCredentials"
    if (-not $useDefaultCredentials) {
        Write-Verbose "  Credential User: $($Credential.UserName)"
    }
    Write-Verbose "==============================="

    # Configure SSL certificate validation bypass if requested
    if ($SkipCertificateCheck) {
        Write-Verbose "Bypassing SSL certificate validation for EWS connections"
        # For .NET Framework / PowerShell 5.1
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        # Also set TLS 1.2 as minimum
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    }

    Write-Verbose "Loading EWS Managed API assembly..."
    Add-Type -Path $EwsAssemblyPath
    Write-Verbose "EWS assembly loaded successfully"

    # Use Exchange2010_SP2 for maximum compatibility
    # Higher versions (2013_SP1+) may request properties like 'Mentions' that don't exist on older servers
    $exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    Write-Verbose "Using Exchange version: $exchangeVersion"

    # PS5.1 compatible syntax
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new($exchangeVersion)
    }
    else {
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchangeVersion)
    }

    # Configure credentials
    if ($UseDefaultCredentials) {
        Write-Verbose "Using default Windows credentials for EWS (current logged-on user)"
        $service.UseDefaultCredentials = $true
    }
    elseif ($Credential) {
        $networkCred = $Credential.GetNetworkCredential()
        Write-Verbose "Setting EWS credentials for user: $($networkCred.UserName)$(if ($networkCred.Domain) { '@' + $networkCred.Domain })"
        
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
    }
    else {
        Write-Verbose "No credential specified, using default Windows credentials"
        $service.UseDefaultCredentials = $true
    }

    # Configure proxy if specified
    if ($ProxyUrl) {
        Write-Verbose "Configuring EWS proxy: $ProxyUrl"
        $webProxy = New-Object System.Net.WebProxy($ProxyUrl, $true)
        if ($Credential) {
            $webProxy.Credentials = $Credential.GetNetworkCredential()
        }
        else {
            $webProxy.UseDefaultCredentials = $true
        }
        $service.WebProxy = $webProxy
        Write-Verbose "EWS proxy configured"
    }

    if ($ExplicitUrl) {
        Write-Verbose "Setting explicit EWS URL: $ExplicitUrl"
        $service.Url = [Uri]$ExplicitUrl
    }
    else {
        Write-Verbose "Running Autodiscover for: $ImpersonationSmtp"
        $redirectCallback = { param($url) return $url -like 'https://*' }
        try {
            $service.AutodiscoverUrl($ImpersonationSmtp, $redirectCallback)
            Write-Verbose "Autodiscover completed. EWS URL: $($service.Url)"
        }
        catch {
            Write-Verbose "ERROR during Autodiscover: $($_.Exception.Message)"
            throw
        }
    }

    if ($ImpersonationSmtp) {
        Write-Verbose "Setting impersonation for: $ImpersonationSmtp"
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
        Write-Verbose "Impersonation configured"
    }

    Write-Verbose "EWS service configured successfully"
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

    # Use IdOnly for initial search to minimize data transfer, then load full properties per item
    # This avoids issues with properties like 'Mentions' that don't exist on all Exchange versions
    $idOnlyPropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $appointments = $null
    $allAppointments = [System.Collections.Generic.List[object]]::new()

    # Function to process a date range chunk
    $processDateRange = {
        param($chunkStart, $chunkEnd)
        
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $view = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($chunkStart, $chunkEnd)
            $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new($idOnlyPropertySet)
        }
        else {
            $view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($chunkStart, $chunkEnd)
            $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet($idOnlyPropertySet)
        }

        return $calendar.FindAppointments($view)
    }

    try {
        # Try full date range first
        $appointments = & $processDateRange $WindowStart $WindowEnd
    }
    catch {
        # If we hit the maximum objects limit, chunk by month
        if ($_.Exception.Message -like '*maximum number of objects*') {
            Write-Verbose "Large calendar detected for '$RoomSmtp', chunking by month..."
            $appointments = $null
            
            $chunkStart = $WindowStart
            while ($chunkStart -lt $WindowEnd) {
                $chunkEnd = $chunkStart.AddMonths(1)
                if ($chunkEnd -gt $WindowEnd) { $chunkEnd = $WindowEnd }
                
                try {
                    $chunkResults = & $processDateRange $chunkStart $chunkEnd
                    # Use foreach iteration instead of .Items to be StrictMode safe
                    foreach ($apt in $chunkResults) {
                        $allAppointments.Add($apt)
                    }
                }
                catch {
                    Write-Warning "Failed to retrieve appointments for room '$RoomSmtp' (chunk $chunkStart to $chunkEnd): $_"
                }
                
                $chunkStart = $chunkEnd
            }
        }
        else {
            Write-Warning "Failed to retrieve appointments for room '$RoomSmtp': $_"
            return @()
        }
    }

    # Get the items to process (either from single call or chunked)
    # Build a simple array to avoid StrictMode issues with .Items property
    $itemsToProcess = @()
    
    if ($null -ne $appointments) {
        # FindAppointments returns FindItemsResults<Appointment> which has an Items collection
        # Iterate to build array instead of accessing .Items directly (StrictMode safe)
        foreach ($apt in $appointments) {
            $itemsToProcess += $apt
        }
    }
    
    # If no items from direct call, use chunked results
    if ($itemsToProcess.Count -eq 0 -and $allAppointments.Count -gt 0) {
        $itemsToProcess = @($allAppointments)
    }
    
    Write-Verbose "  Found $($itemsToProcess.Count) appointments in room '$RoomSmtp'"

    foreach ($item in $itemsToProcess) {
        try {
            # Load full properties for each appointment individually
            # Request only the properties we need to avoid version compatibility issues
            $propsToLoad = [Microsoft.Exchange.WebServices.Data.PropertySet]::new(
                [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::AppointmentType,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::RequiredAttendees,
                [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::OptionalAttendees
            )
            $item.Load($propsToLoad)
            
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
        PSCustomObject with Organizer, Status, Enabled, RecipientType, MailboxType properties.
        MailboxType: User, SharedMailbox, RoomMailbox, EquipmentMailbox, MailUser, MailContact, etc.
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
    $resolvedSmtp = $null
    $matchedInternal = $false

    # If address looks external, try to find matching internal user by local part
    # e.g., lieke.dewit@external.thalesgroup.com -> look for lieke.dewit@nl.thalesgroup.com
    if (-not $isInternal -and $OrganizationSuffix) {
        # Extract local part (before @)
        $localPart = ($SmtpAddress -split '@')[0]
        $internalAddress = "$localPart@$OrganizationSuffix"
        
        Write-Verbose "External address '$SmtpAddress' - checking for internal match '$internalAddress'"
        
        # Try to find internal recipient with constructed address
        $internalRecipient = $null
        if ($ConnectionType -eq 'EXO') {
            $internalRecipient = Get-EXORecipient -Identity $internalAddress -ErrorAction SilentlyContinue
        }
        else {
            $internalRecipient = Get-Recipient -Identity $internalAddress -ErrorAction SilentlyContinue
        }
        
        if ($internalRecipient) {
            # Found matching internal user!
            $resolvedSmtp = $internalRecipient.PrimarySmtpAddress.ToString()
            $matchedInternal = $true
            $isInternal = $true
            Write-Verbose "Matched external '$SmtpAddress' to internal user '$resolvedSmtp'"
        }
        else {
            # No internal match found - truly external
            return [PSCustomObject]@{
                Organizer     = $SmtpAddress
                Status        = 'External'
                Enabled       = $null
                RecipientType = $null
                MailboxType   = 'External'
                ResolvedSmtp  = $null
                MatchedInternal = $false
            }
        }
    }
    elseif (-not $isInternal) {
        # No OrganizationSuffix configured, can't do internal lookup
        return [PSCustomObject]@{
            Organizer     = $SmtpAddress
            Status        = 'External'
            Enabled       = $null
            RecipientType = $null
            MailboxType   = 'External'
            ResolvedSmtp  = $null
            MatchedInternal = $false
        }
    }

    # Lookup recipient - use Get-Mailbox first to get RecipientTypeDetails
    $mailbox = $null
    $recipient = $null
    $recipientTypeDetails = $null
    
    if ($ConnectionType -eq 'EXO') {
        $mailbox = Get-EXOMailbox -Identity $SmtpAddress -ErrorAction SilentlyContinue
        if ($mailbox) {
            $recipientTypeDetails = $mailbox.RecipientTypeDetails
        }
        else {
            $recipient = Get-EXORecipient -Identity $SmtpAddress -ErrorAction SilentlyContinue
            if ($recipient) {
                $recipientTypeDetails = $recipient.RecipientTypeDetails
            }
        }
    }
    else {
        $mailbox = Get-Mailbox -Identity $SmtpAddress -ErrorAction SilentlyContinue
        if ($mailbox) {
            $recipientTypeDetails = $mailbox.RecipientTypeDetails
        }
        else {
            $recipient = Get-Recipient -Identity $SmtpAddress -ErrorAction SilentlyContinue
            if ($recipient) {
                $recipientTypeDetails = $recipient.RecipientTypeDetails
            }
        }
    }

    if (-not $mailbox -and -not $recipient) {
        return [PSCustomObject]@{
            Organizer       = $SmtpAddress
            Status          = 'NotFound'
            Enabled         = $null
            RecipientType   = $null
            MailboxType     = 'NotFound'
            ResolvedSmtp    = $null
            MatchedInternal = $false
        }
    }

    # Map RecipientTypeDetails to friendly MailboxType
    $mailboxType = switch ($recipientTypeDetails) {
        'UserMailbox'           { 'User' }
        'SharedMailbox'         { 'SharedMailbox' }
        'RoomMailbox'           { 'RoomMailbox' }
        'EquipmentMailbox'      { 'EquipmentMailbox' }
        'SchedulingMailbox'     { 'SchedulingMailbox' }
        'LinkedMailbox'         { 'LinkedMailbox' }
        'MailUser'              { 'MailUser' }
        'MailContact'           { 'MailContact' }
        'MailUniversalDistributionGroup' { 'DistributionGroup' }
        'MailUniversalSecurityGroup'     { 'SecurityGroup' }
        'DynamicDistributionGroup'       { 'DynamicGroup' }
        'PublicFolder'          { 'PublicFolder' }
        'RemoteUserMailbox'     { 'RemoteUser' }
        'RemoteSharedMailbox'   { 'RemoteSharedMailbox' }
        'RemoteRoomMailbox'     { 'RemoteRoomMailbox' }
        'RemoteEquipmentMailbox' { 'RemoteEquipmentMailbox' }
        default                 { $recipientTypeDetails.ToString() }
    }

    # Check enabled state
    $enabled = $null
    $recipientObj = if ($mailbox) { $mailbox } else { $recipient }
    
    if ($ConnectionType -eq 'EXO') {
        if ($mailbox -and ($mailbox | Get-Member -Name AccountDisabled -ErrorAction SilentlyContinue)) {
            $enabled = -not $mailbox.AccountDisabled
        }
    }
    else {
        # Try ActiveDirectory module for on-prem (only for user mailboxes)
        if ($mailboxType -eq 'User' -and $recipientObj.SamAccountName) {
            $adLoaded = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue
            if (-not $adLoaded) {
                Import-Module ActiveDirectory -ErrorAction SilentlyContinue | Out-Null
                $adLoaded = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue
            }

            if ($adLoaded) {
                $adUser = Get-ADUser -Identity $recipientObj.SamAccountName -Properties Enabled -ErrorAction SilentlyContinue
                if ($adUser) {
                    $enabled = $adUser.Enabled
                }
            }
            else {
                Write-Verbose 'ActiveDirectory module not available; skipping enabled-state lookup.'
            }
        }
        elseif ($mailboxType -in @('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')) {
            # Resource mailboxes are typically always "enabled" for scheduling purposes
            $enabled = $true
        }
    }

    $status = if ($enabled -eq $false) { 'Disabled' } else { 'Active' }

    return [PSCustomObject]@{
        Organizer       = $SmtpAddress
        Status          = $status
        Enabled         = $enabled
        RecipientType   = $recipientTypeDetails
        MailboxType     = $mailboxType
        ResolvedSmtp    = $resolvedSmtp
        MatchedInternal = $matchedInternal
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
