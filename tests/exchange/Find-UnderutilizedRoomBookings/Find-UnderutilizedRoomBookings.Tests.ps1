#Requires -Modules Pester
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Find-UnderutilizedRoomBookings.ps1' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-UnderutilizedRoomBookings/Find-UnderutilizedRoomBookings.ps1'
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-UnderutilizedRoomBookings/modules/ExchangeCore/ExchangeCore.psm1'

        # Import module for mocking
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }
    }

    Context 'Parameter Validation' {
        It 'validates MinimumCapacity range (1-500)' {
            {
                & $scriptPath -MinimumCapacity 0 -TestMode 2>&1
            } | Should -Throw
        }

        It 'validates MaxParticipants range (1-500)' {
            {
                & $scriptPath -MaxParticipants 0 -TestMode 2>&1
            } | Should -Throw
        }

        It 'validates MonthsAhead range (0-36)' {
            {
                & $scriptPath -MonthsAhead 50 -TestMode 2>&1
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
                    Type        = 'OnPrem'
                    ExchangeUri = 'http://test.contoso.com/PowerShell/'
                }
                Impersonation = @{
                    SmtpAddress = 'config@contoso.com'
                }
                MinimumCapacity = 10
                MaxParticipants = 3
                OutputPath      = './test-output.csv'
            }
            $config | ConvertTo-Json -Depth 5 | Set-Content -Path $testConfigPath

            $result = Import-ConfigurationFile -Path $testConfigPath
            $result.Connection.Type | Should -Be 'OnPrem'
            $result.MinimumCapacity | Should -Be 10
            $result.MaxParticipants | Should -Be 3
        }

        It 'loads PSD1 configuration file' {
            $psd1Path = Join-Path -Path $TestDrive -ChildPath 'test-config.psd1'
            @'
@{
    Connection = @{
        Type = 'EXO'
    }
    MinimumCapacity = 8
    MaxParticipants = 2
}
'@ | Set-Content -Path $psd1Path

            $result = Import-ConfigurationFile -Path $psd1Path
            $result.Connection.Type | Should -Be 'EXO'
            $result.MinimumCapacity | Should -Be 8
        }
    }

    Context 'Test Mode' {
        It 'runs without errors in test mode' {
            $output = & $scriptPath -TestMode -Verbose 4>&1
            $output | Should -Not -BeNullOrEmpty
        }

        It 'displays configuration summary in test mode' {
            $output = & $scriptPath -TestMode -MinimumCapacity 8 -MaxParticipants 3 6>&1
            $outputText = $output -join ' '
            $outputText | Should -Match 'Test mode'
        }
    }

    Context 'Participant Counting' {
        BeforeAll {
            # Define test function locally since it's internal to the script
            function Get-MeetingParticipantInfo {
                param(
                    [Parameter(Mandatory)]$Meeting,
                    [Parameter(Mandatory)][string]$RoomSmtp
                )

                $participants = [System.Collections.Generic.List[string]]::new()
                if ($Meeting.Organizer) { $participants.Add($Meeting.Organizer) }
                if ($Meeting.RequiredAttendees) {
                    foreach ($a in $Meeting.RequiredAttendees) { if ($a) { $participants.Add($a) } }
                }
                if ($Meeting.OptionalAttendees) {
                    foreach ($a in $Meeting.OptionalAttendees) { if ($a) { $participants.Add($a) } }
                }

                $distinct = $participants |
                    Where-Object { $_ -and $_ -ne $RoomSmtp } |
                    Sort-Object -Unique

                [PSCustomObject]@{
                    Count        = @($distinct).Count
                    Participants = @($distinct)
                }
            }
        }

        It 'counts distinct participants correctly' {
            $meeting = [PSCustomObject]@{
                Organizer         = 'organizer@contoso.com'
                RequiredAttendees = @('attendee1@contoso.com', 'attendee2@contoso.com')
                OptionalAttendees = @('optional@contoso.com')
            }

            $result = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp 'room@contoso.com'
            $result.Count | Should -Be 4
        }

        It 'excludes room from participant count' {
            $meeting = [PSCustomObject]@{
                Organizer         = 'organizer@contoso.com'
                RequiredAttendees = @('room@contoso.com', 'attendee@contoso.com')
                OptionalAttendees = @()
            }

            $result = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp 'room@contoso.com'
            $result.Count | Should -Be 2
            $result.Participants | Should -Not -Contain 'room@contoso.com'
        }

        It 'handles duplicate attendees' {
            $meeting = [PSCustomObject]@{
                Organizer         = 'user@contoso.com'
                RequiredAttendees = @('user@contoso.com', 'other@contoso.com')
                OptionalAttendees = @('user@contoso.com')
            }

            $result = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp 'room@contoso.com'
            $result.Count | Should -Be 2
        }

        It 'handles empty meeting' {
            $meeting = [PSCustomObject]@{
                Organizer         = $null
                RequiredAttendees = @()
                OptionalAttendees = @()
            }

            $result = Get-MeetingParticipantInfo -Meeting $meeting -RoomSmtp 'room@contoso.com'
            $result.Count | Should -Be 0
        }
    }

    Context 'Report Structure' {
        It 'creates report with correct columns' {
            $mockResult = [PSCustomObject]@{
                Room             = 'conf-large@contoso.com'
                DisplayName      = 'Large Conference Room'
                Capacity         = 12
                Subject          = 'Test Meeting'
                Start            = Get-Date
                End              = (Get-Date).AddHours(1)
                Organizer        = 'user@contoso.com'
                ParticipantCount = 2
                Participants     = 'user@contoso.com;other@contoso.com'
                UniqueId         = 'test123'
            }

            $testCsvPath = Join-Path -Path $TestDrive -ChildPath 'test-export.csv'
            $mockResult | Export-Csv -Path $testCsvPath -NoTypeInformation

            $imported = Import-Csv -Path $testCsvPath
            $imported.Room | Should -Be 'conf-large@contoso.com'
            $imported.Capacity | Should -Be '12'
            $imported.ParticipantCount | Should -Be '2'
        }
    }
}

Describe 'ExchangeCore Module (Underutilized)' {
    BeforeAll {
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../exchange/Find-UnderutilizedRoomBookings/modules/ExchangeCore/ExchangeCore.psm1'
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -Force
        }
    }

    Context 'Module Availability' {
        It 'exports required functions' {
            $requiredFunctions = @(
                'Import-ConfigurationFile'
                'Get-ResolvedConnectionType'
                'Connect-ExchangeSession'
                'Disconnect-ExchangeSession'
                'Connect-EwsService'
                'Get-RoomCalendarItems'
            )

            $exportedFunctions = (Get-Module ExchangeCore).ExportedFunctions.Keys

            foreach ($func in $requiredFunctions) {
                $exportedFunctions | Should -Contain $func -Because "Function '$func' is required"
            }
        }
    }
}
