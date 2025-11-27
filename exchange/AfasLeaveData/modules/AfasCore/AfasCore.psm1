#Requires -Version 5.1
<#
.SYNOPSIS
    Shared AFAS functions for AfasLeaveData scripts.

.DESCRIPTION
    Provides reusable functions for AFAS data retrieval via Integration Bus REST API,
    configuration management, logging, and user mapping.

.NOTES
    Compatible with PowerShell 5.1+ and PowerShell 7+
    Version: 1.0.0
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration

function Import-ConfigurationFile {
    <#
    .SYNOPSIS
        Loads configuration from JSON or PSD1 file.

    .DESCRIPTION
        Reads a configuration file and returns it as a hashtable.
        Supports both JSON and PowerShell Data (.psd1) formats.

    .PARAMETER Path
        Path to the configuration file.

    .OUTPUTS
        [hashtable] Configuration settings.

    .EXAMPLE
        $config = Import-ConfigurationFile -Path './config.json'
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Configuration file not found: $Path"
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($extension) {
        '.json' {
            $content = Get-Content -Path $Path -Raw -Encoding UTF8

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                return ($content | ConvertFrom-Json -AsHashtable)
            }
            else {
                return (ConvertTo-Hashtable -InputObject ($content | ConvertFrom-Json))
            }
        }
        '.psd1' {
            return (Import-PowerShellDataFile -Path $Path)
        }
        default {
            throw "Unsupported configuration file format: $extension. Use .json or .psd1"
        }
    }
}

function ConvertTo-Hashtable {
    <#
    .SYNOPSIS
        Converts a PSCustomObject to a hashtable (recursive).

    .DESCRIPTION
        Used for PowerShell 5.1 compatibility when parsing JSON.

    .PARAMETER InputObject
        The object to convert.

    .OUTPUTS
        [hashtable] The converted hashtable.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @(
                foreach ($item in $InputObject) {
                    ConvertTo-Hashtable -InputObject $item
                }
            )
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
}

#endregion Configuration

#region Credential Management

function Get-StoredCredential {
    <#
    .SYNOPSIS
        Retrieves credentials from password file.

    .DESCRIPTION
        Reads a SecureString password from a file and creates a PSCredential object.
        Password file should be created with: 
        Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File password.txt

    .PARAMETER Username
        The username for the credential.

    .PARAMETER PasswordFilePath
        Path to the password file containing encrypted password.

    .OUTPUTS
        [PSCredential] The credential object.

    .EXAMPLE
        $cred = Get-StoredCredential -Username 'svc_account' -PasswordFilePath 'C:\password.txt'
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Username,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PasswordFilePath
    )

    if (-not (Test-Path -Path $PasswordFilePath)) {
        throw "Password file not found: $PasswordFilePath"
    }

    try {
        $securePassword = Get-Content -Path $PasswordFilePath | ConvertTo-SecureString
        return New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    }
    catch {
        throw "Failed to read password file: $_"
    }
}

#endregion Credential Management

#region Logging

function Write-LogEntry {
    <#
    .SYNOPSIS
        Writes a log entry to file (compatible with legacy format).

    .DESCRIPTION
        Writes tab-separated log entries matching the legacy script format.

    .PARAMETER LogFile
        Path to the log file.

    .PARAMETER Context
        The task context (padded to 20 chars).

    .PARAMETER Status
        Status indicator: [START], [SUCCESS], [ERROR], [WARNING], [STOP]

    .PARAMETER Message
        The log message.

    .EXAMPLE
        Write-LogEntry -LogFile $logPath -Context 'LoadEWS' -Status '[SUCCESS]' -Message 'EWS loaded'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LogFile,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
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
    Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8
}

function New-LogFile {
    <#
    .SYNOPSIS
        Creates a new log file with timestamp.

    .PARAMETER LogPath
        Directory for log files.

    .PARAMETER Prefix
        Log file name prefix.

    .OUTPUTS
        [string] Full path to the new log file.

    .EXAMPLE
        $logFile = New-LogFile -LogPath 'C:\Logs' -Prefix 'ImportCalendar'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LogPath,

        [Parameter()]
        [string]$Prefix = 'AfasLeaveData'
    )

    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'ddMMyyyy-HHmm'
    $logFile = Join-Path -Path $LogPath -ChildPath "$Prefix-$timestamp.log"
    
    return $logFile
}

#endregion Logging

#region API Functions

function Test-ApiEndpoint {
    <#
    .SYNOPSIS
        Tests if an API endpoint is reachable.

    .DESCRIPTION
        Makes an HTTP request to verify the endpoint is available.
        Supports proxy configuration.

    .PARAMETER Uri
        The API endpoint URL.

    .PARAMETER Credential
        Credentials for authentication.

    .PARAMETER ProxyUrl
        Optional proxy URL.

    .OUTPUTS
        [bool] True if endpoint returns HTTP 200.

    .EXAMPLE
        $available = Test-ApiEndpoint -Uri 'https://api.example.com' -Credential $cred
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [string]$ProxyUrl
    )

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $request = [System.Net.WebRequest]::Create($Uri)
        
        if ($Credential) {
            $request.Credentials = $Credential.GetNetworkCredential()
        }

        if ($ProxyUrl) {
            $request.Proxy = [System.Net.WebProxy]::new($ProxyUrl)
        }

        $response = $request.GetResponse()
        $statusCode = [int]$response.StatusCode
        $response.Close()

        return ($statusCode -eq 200)
    }
    catch {
        Write-Warning "API endpoint test failed: $_"
        return $false
    }
}

function Get-LeaveDataFromApi {
    <#
    .SYNOPSIS
        Retrieves leave data from Integration Bus API.

    .DESCRIPTION
        Calls the REST API endpoint and downloads leave data as CSV.
        Compatible with legacy script's data format.

    .PARAMETER Uri
        The API endpoint URL.

    .PARAMETER Credential
        Credentials for authentication.

    .PARAMETER OutputPath
        Path where to save the CSV file.

    .PARAMETER ProxyUrl
        Optional proxy URL.

    .OUTPUTS
        [PSCustomObject[]] Array of leave data entries.

    .EXAMPLE
        $data = Get-LeaveDataFromApi -Uri $endpoint -Credential $cred -OutputPath 'C:\data.csv'
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [Parameter()]
        [string]$ProxyUrl
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $invokeParams = @{
        Uri         = $Uri
        Credential  = $Credential
        OutFile     = $OutputPath
        Method      = 'Get'
    }

    if ($ProxyUrl) {
        $invokeParams.Proxy = $ProxyUrl
    }

    try {
        Invoke-RestMethod @invokeParams

        if (Test-Path -Path $OutputPath) {
            $csvData = Import-Csv -Path $OutputPath
            
            # Validate required fields
            if ($csvData.Count -gt 0) {
                $requiredFields = @('ITCode', 'StartDate', 'StartTime', 'EndDate', 'EndTime')
                foreach ($field in $requiredFields) {
                    if (-not $csvData[0].PSObject.Properties.Name.Contains($field)) {
                        Write-Warning "CSV missing required field: $field"
                    }
                }
            }

            return $csvData
        }
        else {
            throw "CSV file was not created at: $OutputPath"
        }
    }
    catch {
        Write-Error "Failed to retrieve leave data from API: $_"
        throw
    }
}

#endregion API Functions

#region User Mapping

function Get-UserEmailFromITCode {
    <#
    .SYNOPSIS
        Maps an ITCode to an Exchange email address.

    .DESCRIPTION
        Looks up the email address for an employee using Get-Mailbox
        (like legacy scripts) or alternative strategies.

    .PARAMETER ITCode
        The employee ITCode/identifier.

    .PARAMETER Strategy
        Mapping strategy: 'Mailbox', 'Email', 'MappingTable'

    .PARAMETER MappingTable
        Hashtable of ITCode -> Email mappings.

    .PARAMETER DefaultDomain
        Default email domain suffix.

    .OUTPUTS
        [string] The resolved email address.

    .EXAMPLE
        $email = Get-UserEmailFromITCode -ITCode 'jsmith' -Strategy 'Mailbox'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ITCode,

        [Parameter()]
        [ValidateSet('Mailbox', 'Email', 'MappingTable')]
        [string]$Strategy = 'Mailbox',

        [Parameter()]
        [hashtable]$MappingTable,

        [Parameter()]
        [string]$DefaultDomain
    )

    switch ($Strategy) {
        'Mailbox' {
            # Use Get-Mailbox like legacy scripts
            try {
                $mailbox = Get-Mailbox -Identity $ITCode -ErrorAction Stop
                if ($mailbox.PrimarySmtpAddress) {
                    Write-Verbose "Resolved via Get-Mailbox: $($mailbox.PrimarySmtpAddress)"
                    return $mailbox.PrimarySmtpAddress.ToString()
                }
            }
            catch {
                Write-Warning "Get-Mailbox failed for $ITCode : $_"
            }
        }

        'MappingTable' {
            if ($MappingTable -and $MappingTable.ContainsKey($ITCode)) {
                $mapped = $MappingTable[$ITCode]
                Write-Verbose "Using mapped email: $mapped"
                return $mapped
            }
        }

        'Email' {
            if ($DefaultDomain) {
                $generated = "$ITCode@$DefaultDomain"
                Write-Verbose "Using generated email: $generated"
                return $generated
            }
        }
    }

    # Fallback to default domain
    if ($DefaultDomain) {
        $fallback = "$ITCode@$DefaultDomain"
        Write-Verbose "Fallback to: $fallback"
        return $fallback
    }

    Write-Warning "Could not resolve email for ITCode: $ITCode"
    return $null
}

#endregion User Mapping

#region CSV File Management

function Move-ProcessedFile {
    <#
    .SYNOPSIS
        Moves a processed file to the processed folder.

    .PARAMETER SourcePath
        Path to the source file.

    .PARAMETER ProcessedPath
        Destination directory.

    .EXAMPLE
        Move-ProcessedFile -SourcePath 'C:\data.csv' -ProcessedPath 'C:\Processed'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ProcessedPath
    )

    if (-not (Test-Path -Path $ProcessedPath)) {
        New-Item -Path $ProcessedPath -ItemType Directory -Force | Out-Null
    }

    if (Test-Path -Path $SourcePath) {
        Move-Item -Path $SourcePath -Destination $ProcessedPath -Force
        Write-Verbose "Moved $SourcePath to $ProcessedPath"
    }
}

function Get-PendingCsvFiles {
    <#
    .SYNOPSIS
        Gets pending CSV files that need processing.

    .PARAMETER ScriptPath
        Directory containing CSV files.

    .PARAMETER Pattern
        File name pattern to match.

    .PARAMETER ProcessedPath
        Where to move existing files.

    .OUTPUTS
        [System.IO.FileInfo[]] Array of matching files.

    .EXAMPLE
        $files = Get-PendingCsvFiles -ScriptPath 'C:\Data' -Pattern 'LeaveData*'
    #>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptPath,

        [Parameter()]
        [string]$Pattern = 'LeaveDataExchange*',

        [Parameter()]
        [string]$ProcessedPath
    )

    $files = Get-ChildItem -Path $ScriptPath -Filter $Pattern -File -ErrorAction SilentlyContinue

    if ($files -and $ProcessedPath) {
        foreach ($file in $files) {
            Write-Warning "Existing file found: $($file.Name) - moving to processed"
            Move-ProcessedFile -SourcePath $file.FullName -ProcessedPath $ProcessedPath
        }
        return @()
    }

    return $files
}

#endregion CSV File Management

# Export public functions
Export-ModuleMember -Function @(
    # Configuration
    'Import-ConfigurationFile'
    'ConvertTo-Hashtable'
    # Credentials
    'Get-StoredCredential'
    # Logging
    'Write-LogEntry'
    'New-LogFile'
    # API
    'Test-ApiEndpoint'
    'Get-LeaveDataFromApi'
    # User Mapping
    'Get-UserEmailFromITCode'
    # File Management
    'Move-ProcessedFile'
    'Get-PendingCsvFiles'
)
