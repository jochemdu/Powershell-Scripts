#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for AfasLeaveData.ps1

.DESCRIPTION
    Unit tests for the AFAS leave data to Exchange calendar sync script.
    Tests the main script and the AfasCore module.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'AfasLeaveData.ps1' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/AfasLeaveData/AfasLeaveData.ps1'
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/AfasLeaveData/modules/AfasCore/AfasCore.psm1'

        # Import module for testing
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }

        # Create test config matching the new structure
        $script:testConfigPath = Join-Path -Path $TestDrive -ChildPath 'config.json'
        $testConfig = @{
            Connection    = @{
                Type        = 'OnPrem'
                ExchangeUri = 'http://exchange.test.local/PowerShell/'
            }
            Credential    = @{
                Username     = 'svc_test'
                PasswordFile = 'C:\test\password.txt'
            }
            Api           = @{
                LeaveDataEndpoint     = 'https://integrationbus.test.local/api/v1/getLeaveData'
                CanceledLeaveEndpoint = 'https://integrationbus.test.local/api/v1/getCanceledLeaveData'
                ProxyUrl              = $null
            }
            LeaveSettings = @{
                Subject = 'MyPlace Leave Booking (automatically created)'
                Body    = 'Please do not modify this appointment'
            }
            Paths         = @{
                ScriptPath    = 'C:\ScheduledTasks\AfasLeaveData'
                ProcessedPath = 'C:\ScheduledTasks\AfasLeaveData\Processed'
                LogPath       = 'C:\ScheduledTasks\AfasLeaveData\Logs'
            }
            EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
        }
        $testConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $script:testConfigPath -Encoding UTF8
    }

    Context 'Script File' {
        It 'Script file exists' {
            Test-Path -Path $scriptPath | Should -BeTrue
        }

        It 'Script has valid PowerShell syntax' {
            $errors = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$errors)
            $errors.Count | Should -Be 0
        }
    }

    Context 'Parameter Validation' {
        It 'Has ConfigPath parameter' {
            $script = Get-Command -Name $scriptPath -ErrorAction SilentlyContinue
            if ($script) {
                $script.Parameters.Keys | Should -Contain 'ConfigPath'
            }
        }

        It 'Has TestMode switch parameter' {
            $script = Get-Command -Name $scriptPath -ErrorAction SilentlyContinue
            if ($script) {
                $script.Parameters.Keys | Should -Contain 'TestMode'
            }
        }

        It 'Has Mode parameter with valid values' {
            $script = Get-Command -Name $scriptPath -ErrorAction SilentlyContinue
            if ($script -and $script.Parameters.ContainsKey('Mode')) {
                $validateSet = $script.Parameters['Mode'].Attributes | 
                    Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
                $validateSet.ValidValues | Should -Contain 'Import'
                $validateSet.ValidValues | Should -Contain 'Remove'
                $validateSet.ValidValues | Should -Contain 'Both'
            }
        }

        It 'Has Credential parameter' {
            $script = Get-Command -Name $scriptPath -ErrorAction SilentlyContinue
            if ($script) {
                $script.Parameters.Keys | Should -Contain 'Credential'
            }
        }
    }
}

Describe 'AfasCore Module' {
    BeforeAll {
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/AfasLeaveData/modules/AfasCore/AfasCore.psm1'

        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }
        else {
            Write-Warning "Module not found: $modulePath"
        }
    }

    Context 'Module Loading' {
        It 'Module file exists' {
            $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/AfasLeaveData/modules/AfasCore/AfasCore.psm1'
            Test-Path -Path $modulePath | Should -BeTrue
        }
    }

    Context 'Exported Functions' {
        It 'Exports Import-ConfigurationFile function' {
            Get-Command -Name 'Import-ConfigurationFile' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports ConvertTo-Hashtable function' {
            Get-Command -Name 'ConvertTo-Hashtable' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports Get-StoredCredential function' {
            Get-Command -Name 'Get-StoredCredential' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports Write-LogEntry function' {
            Get-Command -Name 'Write-LogEntry' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports Test-ApiEndpoint function' {
            Get-Command -Name 'Test-ApiEndpoint' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports Get-LeaveDataFromApi function' {
            Get-Command -Name 'Get-LeaveDataFromApi' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exports Get-UserEmailFromITCode function' {
            Get-Command -Name 'Get-UserEmailFromITCode' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Import-ConfigurationFile' {
        BeforeAll {
            $script:jsonConfigPath = Join-Path -Path $TestDrive -ChildPath 'test-config.json'
            $script:psd1ConfigPath = Join-Path -Path $TestDrive -ChildPath 'test-config.psd1'

            @{
                Connection = @{ Type = 'OnPrem' }
                Api        = @{ LeaveDataEndpoint = 'https://api.test/leave' }
            } | ConvertTo-Json -Depth 5 | Set-Content -Path $script:jsonConfigPath

            @"
@{
    Connection = @{ Type = 'OnPrem' }
    Api        = @{ LeaveDataEndpoint = 'https://api.test/leave' }
}
"@ | Set-Content -Path $script:psd1ConfigPath
        }

        It 'Loads JSON configuration file' {
            $config = Import-ConfigurationFile -Path $script:jsonConfigPath
            $config | Should -Not -BeNullOrEmpty
            $config.Connection.Type | Should -Be 'OnPrem'
        }

        It 'Loads PSD1 configuration file' {
            $config = Import-ConfigurationFile -Path $script:psd1ConfigPath
            $config | Should -Not -BeNullOrEmpty
            $config.Connection.Type | Should -Be 'OnPrem'
        }

        It 'Throws for non-existent file' {
            { Import-ConfigurationFile -Path 'C:\nonexistent\file.json' } | Should -Throw
        }

        It 'Throws for unsupported file format' {
            $xmlPath = Join-Path -Path $TestDrive -ChildPath 'config.xml'
            '<config/>' | Set-Content -Path $xmlPath
            { Import-ConfigurationFile -Path $xmlPath } | Should -Throw '*Unsupported*'
        }
    }

    Context 'ConvertTo-Hashtable' {
        It 'Converts PSCustomObject to hashtable' {
            $obj = [PSCustomObject]@{ Name = 'Test'; Value = 123 }
            $result = ConvertTo-Hashtable -InputObject $obj
            $result | Should -BeOfType [hashtable]
            $result.Name | Should -Be 'Test'
            $result.Value | Should -Be 123
        }

        It 'Handles nested objects' {
            $obj = [PSCustomObject]@{
                Level1 = [PSCustomObject]@{
                    Level2 = 'DeepValue'
                }
            }
            $result = ConvertTo-Hashtable -InputObject $obj
            $result.Level1 | Should -BeOfType [hashtable]
            $result.Level1.Level2 | Should -Be 'DeepValue'
        }

        It 'Handles null input' {
            $result = ConvertTo-Hashtable -InputObject $null
            $result | Should -BeNullOrEmpty
        }

        It 'Handles arrays' {
            $obj = @(
                [PSCustomObject]@{ Id = 1 },
                [PSCustomObject]@{ Id = 2 }
            )
            $result = ConvertTo-Hashtable -InputObject $obj
            $result | Should -HaveCount 2
            $result[0].Id | Should -Be 1
        }
    }

    Context 'Get-StoredCredential' {
        BeforeAll {
            # Create test password file
            $script:testPasswordFile = Join-Path -Path $TestDrive -ChildPath 'password.txt'
            $securePass = ConvertTo-SecureString 'TestPassword123' -AsPlainText -Force
            $securePass | ConvertFrom-SecureString | Set-Content -Path $script:testPasswordFile
        }

        It 'Creates credential from password file' {
            $cred = Get-StoredCredential -Username 'testuser' -PasswordFilePath $script:testPasswordFile
            $cred | Should -Not -BeNullOrEmpty
            $cred.UserName | Should -Be 'testuser'
        }

        It 'Throws for non-existent password file' {
            { Get-StoredCredential -Username 'test' -PasswordFilePath 'C:\nonexistent.txt' } | Should -Throw
        }
    }

    Context 'Write-LogEntry' {
        BeforeAll {
            $script:testLogFile = Join-Path -Path $TestDrive -ChildPath 'test.log'
        }

        It 'Writes log entry to file' {
            Write-LogEntry -LogFile $script:testLogFile -Context 'TestContext' -Status '[SUCCESS]' -Message 'Test message'
            Test-Path -Path $script:testLogFile | Should -BeTrue
            $content = Get-Content -Path $script:testLogFile
            $content | Should -Match 'TestContext'
            $content | Should -Match '\[SUCCESS\]'
            $content | Should -Match 'Test message'
        }

        It 'Pads context to 20 characters' {
            $logFile = Join-Path -Path $TestDrive -ChildPath 'pad-test.log'
            Write-LogEntry -LogFile $logFile -Context 'Short' -Status '[INFO]' -Message 'Padding test'
            $content = Get-Content -Path $logFile -Raw
            # Context should be padded with spaces
            $content | Should -Match 'Short\s+'
        }
    }

    Context 'New-LogFile' {
        It 'Creates log directory if not exists' {
            $newLogPath = Join-Path -Path $TestDrive -ChildPath 'NewLogs'
            $logFile = New-LogFile -LogPath $newLogPath -Prefix 'Test'
            Test-Path -Path $newLogPath | Should -BeTrue
            $logFile | Should -Match 'Test-\d{8}-\d{4}\.log$'
        }

        It 'Returns full path with timestamp' {
            $logFile = New-LogFile -LogPath $TestDrive -Prefix 'AfasLeaveData'
            $logFile | Should -Match 'AfasLeaveData-\d{8}-\d{4}\.log$'
        }
    }

    Context 'Get-UserEmailFromITCode' {
        It 'Uses mapping table when available' {
            $mapping = @{ 'EMP001' = 'mapped@test.com' }
            $result = Get-UserEmailFromITCode -ITCode 'EMP001' -Strategy 'MappingTable' -MappingTable $mapping
            $result | Should -Be 'mapped@test.com'
        }

        It 'Falls back to default domain with Email strategy' {
            $result = Get-UserEmailFromITCode -ITCode 'EMP001' -Strategy 'Email' -DefaultDomain 'fallback.com'
            $result | Should -Be 'EMP001@fallback.com'
        }

        It 'Returns null when no resolution possible' {
            $result = Get-UserEmailFromITCode -ITCode 'EMP001' -Strategy 'MappingTable' -WarningAction SilentlyContinue
            $result | Should -BeNullOrEmpty
        }
    }
}
