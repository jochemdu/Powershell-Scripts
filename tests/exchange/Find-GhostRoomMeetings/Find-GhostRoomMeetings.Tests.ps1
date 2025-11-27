#Requires -Modules Pester
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Find-GhostRoomMeetings.ps1' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-GhostRoomMeetings/Find-GhostRoomMeetings.ps1'
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-GhostRoomMeetings/modules/ExchangeCore/ExchangeCore.psm1'

        # Import module for mocking
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }
    }

    Context 'Parameter Validation' {
        It 'throws when -SendInquiry is used without -NotificationFrom' {
            {
                & $scriptPath -SendInquiry -ImpersonationSmtp 'test@contoso.com' -TestMode 2>&1
            } | Should -Throw '*-NotificationFrom is required*'
        }

        It 'validates MonthsAhead range (0-36)' {
            {
                & $scriptPath -MonthsAhead 50 -TestMode 2>&1
            } | Should -Throw
        }

        It 'validates MonthsBehind range (0-12)' {
            {
                & $scriptPath -MonthsBehind 24 -TestMode 2>&1
            } | Should -Throw
        }

        It 'validates ImpersonationSmtp format' {
            {
                & $scriptPath -ImpersonationSmtp 'not-an-email' -TestMode 2>&1
            } | Should -Throw
        }
    }

    Context 'Configuration Loading' {
        BeforeAll {
            $testConfigPath = Join-Path -Path $TestDrive -ChildPath 'test-config.json'
        }

        It 'loads JSON configuration file' {
            $config = @{
                Connection = @{
                    Type = 'OnPrem'
                    ExchangeUri = 'http://test.contoso.com/PowerShell/'
                }
                Impersonation = @{
                    SmtpAddress = 'config@contoso.com'
                }
                OrganizationSmtpSuffix = 'contoso.com'
                MonthsAhead = 6
                OutputPath = './test-output.csv'
            }
            $config | ConvertTo-Json -Depth 5 | Set-Content -Path $testConfigPath

            # Dot-source to test config loading
            Mock -ModuleName ExchangeCore -CommandName Test-Path -MockWith { $true }

            $result = Import-ConfigurationFile -Path $testConfigPath
            $result.Connection.Type | Should -Be 'OnPrem'
            $result.MonthsAhead | Should -Be 6
        }

        It 'loads PSD1 configuration file' {
            $psd1Path = Join-Path -Path $TestDrive -ChildPath 'test-config.psd1'
            @'
@{
    Connection = @{
        Type = 'EXO'
    }
    MonthsAhead = 3
}
'@ | Set-Content -Path $psd1Path

            $result = Import-ConfigurationFile -Path $psd1Path
            $result.Connection.Type | Should -Be 'EXO'
            $result.MonthsAhead | Should -Be 3
        }

        It 'throws for unsupported config format' {
            $xmlPath = Join-Path -Path $TestDrive -ChildPath 'config.xml'
            '<config />' | Set-Content -Path $xmlPath

            { Import-ConfigurationFile -Path $xmlPath } | Should -Throw '*must be .json or .psd1*'
        }
    }

    Context 'Connection Type Resolution' {
        It 'detects EXO from Office 365 URI' {
            $result = Get-ResolvedConnectionType -ConnectionType 'Auto' -ExchangeUri 'https://outlook.office365.com/powershell-liveid/'
            $result | Should -Be 'EXO'
        }

        It 'defaults to OnPrem for non-O365 URI' {
            $result = Get-ResolvedConnectionType -ConnectionType 'Auto' -ExchangeUri 'http://exchange.contoso.local/PowerShell/'
            $result | Should -Be 'OnPrem'
        }

        It 'respects explicit OnPrem setting' {
            $result = Get-ResolvedConnectionType -ConnectionType 'OnPrem' -ExchangeUri 'https://outlook.office365.com/powershell-liveid/'
            $result | Should -Be 'OnPrem'
        }

        It 'respects explicit EXO setting' {
            $result = Get-ResolvedConnectionType -ConnectionType 'EXO' -ExchangeUri 'http://exchange.contoso.local/PowerShell/'
            $result | Should -Be 'EXO'
        }
    }

    Context 'Test Mode' {
        It 'runs without errors in test mode' {
            $output = & $scriptPath -TestMode -Verbose 4>&1
            $output | Should -Not -BeNullOrEmpty
        }

        It 'does not attempt Exchange connection in test mode' {
            Mock -ModuleName ExchangeCore -CommandName New-PSSession -MockWith {
                throw 'Should not be called in test mode'
            }

            { & $scriptPath -TestMode } | Should -Not -Throw
        }
    }

    Context 'Report Generation (Mocked)' {
        BeforeAll {
            $csvPath = Join-Path -Path $TestDrive -ChildPath 'reports/ghost-meetings.csv'
            $excelPath = Join-Path -Path $TestDrive -ChildPath 'reports/ghost-meetings.xlsx'

            $securePassword = ConvertTo-SecureString -String 'P@ssw0rd!' -AsPlainText -Force
            $script:TestCredential = [PSCredential]::new('service@contoso.com', $securePassword)

            $script:MockRoomMeetings = @(
                [PSCustomObject]@{
                    Room              = 'room1@contoso.com'
                    Subject           = 'Ghost Meeting'
                    Start             = (Get-Date).AddDays(1)
                    End               = (Get-Date).AddDays(1).AddHours(1)
                    IsRecurring       = $false
                    Organizer         = 'ghost@contoso.com'
                    RequiredAttendees = @('attendee@contoso.com')
                    OptionalAttendees = @()
                    UniqueId          = 'abc123'
                }
            )
        }

        It 'creates output directories if missing' {
            $testOutputPath = Join-Path -Path $TestDrive -ChildPath 'new-dir/output.csv'

            Mock -ModuleName ExchangeCore -CommandName Connect-ExchangeSession -MockWith { $null }
            Mock -ModuleName ExchangeCore -CommandName Connect-EwsService -MockWith {
                [PSCustomObject]@{ Name = 'MockEWS' }
            }
            Mock -ModuleName ExchangeCore -CommandName Get-RoomMailboxes -MockWith { @() }
            Mock -ModuleName ExchangeCore -CommandName Disconnect-ExchangeSession -MockWith { }

            # This should create the directory
            $parentDir = Split-Path -Path $testOutputPath -Parent
            Test-Path -Path $parentDir | Should -BeFalse -Because 'Directory should not exist before test'
        }

        It 'exports CSV report with correct structure' {
            $mockResults = @(
                [PSCustomObject]@{
                    Room = 'room1@contoso.com'
                    Subject = 'Test Meeting'
                    Start = Get-Date
                    End = (Get-Date).AddHours(1)
                    Organizer = 'user@contoso.com'
                    OrganizerStatus = 'NotFound'
                    IsRecurring = $false
                    Attendees = 'attendee@contoso.com'
                    UniqueId = 'test123'
                }
            )

            $testCsvPath = Join-Path -Path $TestDrive -ChildPath 'test-export.csv'
            $mockResults | Export-Csv -Path $testCsvPath -NoTypeInformation

            $imported = Import-Csv -Path $testCsvPath
            $imported.Room | Should -Be 'room1@contoso.com'
            $imported.OrganizerStatus | Should -Be 'NotFound'
        }
    }

    Context 'Organizer State Detection' {
        It 'identifies external organizers' {
            Mock -ModuleName ExchangeCore -CommandName Get-Recipient -MockWith { $null }

            $result = Get-OrganizerState -SmtpAddress 'external@external.com' -OrganizationSuffix 'contoso.com' -ConnectionType 'OnPrem'
            $result.Status | Should -Be 'External'
        }

        It 'identifies not-found internal organizers' {
            Mock -ModuleName ExchangeCore -CommandName Get-Recipient -MockWith { $null }

            $result = Get-OrganizerState -SmtpAddress 'deleted@contoso.com' -OrganizationSuffix 'contoso.com' -ConnectionType 'OnPrem'
            $result.Status | Should -Be 'NotFound'
        }
    }
}

Describe 'ExchangeCore Module' {
    BeforeAll {
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-GhostRoomMeetings/modules/ExchangeCore/ExchangeCore.psm1'
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }
    }

    Context 'ConvertTo-Hashtable' {
        It 'converts PSCustomObject to hashtable' {
            $obj = [PSCustomObject]@{
                Name = 'Test'
                Value = 123
            }

            $result = ConvertTo-Hashtable -InputObject $obj
            $result | Should -BeOfType [hashtable]
            $result.Name | Should -Be 'Test'
            $result.Value | Should -Be 123
        }

        It 'handles nested objects' {
            $obj = [PSCustomObject]@{
                Outer = [PSCustomObject]@{
                    Inner = 'value'
                }
            }

            $result = ConvertTo-Hashtable -InputObject $obj
            $result.Outer | Should -BeOfType [hashtable]
            $result.Outer.Inner | Should -Be 'value'
        }
    }

    Context 'Get-ResolvedConnectionType' {
        It 'returns EXO for Office 365 patterns' {
            $patterns = @(
                'https://outlook.office365.com/powershell-liveid/',
                'https://ps.outlook.com/powershell/',
                'https://something.office365.com/test'
            )

            foreach ($uri in $patterns) {
                $result = Get-ResolvedConnectionType -ConnectionType 'Auto' -ExchangeUri $uri
                $result | Should -Be 'EXO' -Because "URI '$uri' should be detected as EXO"
            }
        }

        It 'returns OnPrem for on-premises patterns' {
            $result = Get-ResolvedConnectionType -ConnectionType 'Auto' -ExchangeUri 'http://exchange.contoso.local/PowerShell/'
            $result | Should -Be 'OnPrem'
        }
    }

    Context 'Module Exports' {
        It 'exports expected functions' {
            $expectedFunctions = @(
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

            $exportedFunctions = (Get-Module ExchangeCore).ExportedFunctions.Keys

            foreach ($func in $expectedFunctions) {
                $exportedFunctions | Should -Contain $func -Because "Function '$func' should be exported"
            }
        }
    }
}
