Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Find-UnderutilizedRoomBookings.ps1' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../exchange/Find-UnderutilizedRoomBookings.ps1'
    }

    Context 'parameter validation' {
        It 'throws when impersonation SMTP cannot be derived' {
            $secure = New-Object System.Security.SecureString
            'password'.ToCharArray() | ForEach-Object { $secure.AppendChar($_) }
            $secure.MakeReadOnly()
            $credential = [pscredential]::new('service', $secure)

            { . $scriptPath -ConnectionType EXO -Credential $credential -TestMode } |
                Should -Throw 'Provide -ImpersonationSmtp (SMTP address) for EWS Autodiscover and impersonation.'
        }
    }

    Context 'configuration defaults' {
        It 'applies config values when parameters are not explicitly bound' {
            $config = [pscustomobject]@{
                MinimumCapacity = 10
                MaxParticipants = 1
                OutputPath      = Join-Path -Path $TestDrive -ChildPath 'under-report.csv'
                ImpersonationSmtp = 'config@contoso.com'
            }

            Mock -CommandName Import-ConfigurationFile -MockWith { return $config }
            Mock -CommandName Connect-ExchangeSession -MockWith { }
            Mock -CommandName Disconnect-ExchangeSession -MockWith { }
            Mock -CommandName Connect-EwsService -MockWith { [pscustomobject]@{} }
            Mock -CommandName Get-RoomMailboxes -MockWith { @() }
            Mock -CommandName Export-Csv -MockWith { }

            . $scriptPath -ConfigPath (Join-Path -Path $TestDrive -ChildPath 'under-config.json') -TestMode

            $MinimumCapacity | Should -Be 10
            $MaxParticipants | Should -Be 1
            $OutputPath | Should -Be $config.OutputPath
            $ImpersonationSmtp | Should -Be $config.ImpersonationSmtp
        }
    }

    Context 'report generation' {
        It 'exports meetings that fall below the participant threshold' {
            $csvPath = Join-Path -Path $TestDrive -ChildPath 'reports/underutilized.csv'
            $rooms = @(
                [pscustomobject]@{ DisplayName = 'Room A'; PrimarySmtpAddress = 'roomA@contoso.com'; ResourceCapacity = 8 },
                [pscustomobject]@{ DisplayName = 'Room B'; PrimarySmtpAddress = 'roomB@contoso.com'; ResourceCapacity = 4 }
            )

            $meetings = @{
                'roomA@contoso.com' = @(
                    [pscustomobject]@{ Room = 'roomA@contoso.com'; Subject = 'Solo focus'; Start = (Get-Date); End = (Get-Date).AddHours(1); Organizer = 'user1@contoso.com'; RequiredAttendees = @(); OptionalAttendees = @() },
                    [pscustomobject]@{ Room = 'roomA@contoso.com'; Subject = 'Team sync'; Start = (Get-Date).AddDays(1); End = (Get-Date).AddDays(1).AddHours(1); Organizer = 'lead@contoso.com'; RequiredAttendees = @('one@contoso.com','two@contoso.com'); OptionalAttendees = @('three@contoso.com') }
                )
                'roomB@contoso.com' = @(
                    [pscustomobject]@{ Room = 'roomB@contoso.com'; Subject = 'Small room'; Start = (Get-Date); End = (Get-Date).AddHours(1); Organizer = 'user2@contoso.com'; RequiredAttendees = @('peer@contoso.com'); OptionalAttendees = @() }
                )
            }

            Mock -CommandName Connect-ExchangeSession -MockWith { }
            Mock -CommandName Disconnect-ExchangeSession -MockWith { }
            Mock -CommandName Connect-EwsService -MockWith { [pscustomobject]@{} }
            Mock -CommandName Get-RoomMailboxes -MockWith { $rooms }
            Mock -CommandName Get-RoomMeetings -MockWith { param($Service,$RoomSmtp) $meetings[$RoomSmtp] }
            Mock -CommandName Export-Csv -MockWith {
                param($InputObject, $Path)
                New-Item -Path $Path -ItemType File -Force | Out-Null
            }

            $report = & $scriptPath -ImpersonationSmtp 'service@contoso.com' -MinimumCapacity 6 -MaxParticipants 2 -OutputPath $csvPath -TestMode

            $report | Should -HaveCount 1
            $report[0].Room | Should -Be 'roomA@contoso.com'
            $report[0].ParticipantCount | Should -Be 1
            Test-Path $csvPath | Should -BeTrue

            Assert-MockCalled -CommandName Export-Csv -Times 1 -ParameterFilter { $Path -eq $csvPath }
            Assert-MockCalled -CommandName Get-RoomMeetings -Times 1 -ParameterFilter { $RoomSmtp -eq 'roomA@contoso.com' }
        }
    }
}
