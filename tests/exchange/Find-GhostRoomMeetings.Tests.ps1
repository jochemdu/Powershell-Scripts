Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Find-GhostRoomMeetings.ps1' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../exchange/Find-GhostRoomMeetings.ps1'
    }

    Context 'parameter validation' {
        It 'throws when -SendInquiry is used without -NotificationFrom' {
            { . $scriptPath -SendInquiry } | Should -Throw '-NotificationFrom is required when -SendInquiry is specified.'
        }
    }

    Context 'configuration defaults' {
        It 'applies values from config when parameters are not bound' {
            $config = [pscustomobject]@{
                OutputPath         = Join-Path -Path $TestDrive -ChildPath 'config-report.csv'
                SendInquiry        = $true
                NotificationFrom   = 'config@contoso.com'
                NotificationTemplate = 'Template from config {0}'
                EwsUrl             = 'https://ews.contoso.com/EWS/Exchange.asmx'
            }

            Mock -CommandName Import-ConfigurationFile -MockWith { return $config }

            . $scriptPath -ConfigPath (Join-Path -Path $TestDrive -ChildPath 'ghost-config.json')

            $OutputPath | Should -Be $config.OutputPath
            $SendInquiry | Should -BeTrue
            $NotificationFrom | Should -Be $config.NotificationFrom
            $NotificationTemplate | Should -Be $config.NotificationTemplate
            $EwsUrl | Should -Be $config.EwsUrl
        }
    }

    Context 'report generation' {
        It 'writes report exports to the TestDrive paths and triggers export commands' {
            $csvPath = Join-Path -Path $TestDrive -ChildPath 'reports/ghost-meetings.csv'
            $excelPath = Join-Path -Path $TestDrive -ChildPath 'reports/ghost-meetings.xlsx'
            $credential = [pscredential]::new(
                'service@contoso.com',
                (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            )

            $roomMeetings = @(
                [pscustomobject]@{
                    Room              = 'room1@contoso.com'
                    Subject           = 'Ghost meeting'
                    Start             = (Get-Date).AddDays(1)
                    End               = (Get-Date).AddDays(1).AddHours(1)
                    IsRecurring       = $false
                    Organizer         = 'ghost@contoso.com'
                    RequiredAttendees = @('attendee@contoso.com')
                    OptionalAttendees = @()
                    UniqueId          = 'abc123'
                    EwsItem           = $null
                }
            )

            Mock -CommandName Import-ConfigurationFile -MockWith { @{ } }
            Mock -CommandName New-PSSession -MockWith { 'session' }
            Mock -CommandName Import-PSSession -MockWith { }
            Mock -CommandName Connect-EwsService -MockWith { return [pscustomobject]@{ Name = 'ews' } }
            Mock -CommandName Disconnect-ExchangeSession -MockWith { }
            Mock -CommandName Get-RoomMailboxes -MockWith { @([pscustomobject]@{ PrimarySmtpAddress = 'room1@contoso.com' }) }
            Mock -CommandName Get-RoomMeetings -MockWith { return $roomMeetings }
            Mock -CommandName Test-OrganizerState -MockWith {
                [pscustomobject]@{
                    Organizer = 'ghost@contoso.com'
                    Status    = 'NotFound'
                    Enabled   = $false
                    Recipient = $null
                }
            }
            Mock -CommandName Send-MailMessage -MockWith { }
            Mock -CommandName Export-Csv -MockWith {
                param($Path)
                New-Item -Path $Path -ItemType File -Force | Out-Null
            }
            Mock -CommandName Export-Excel -MockWith {
                param($Path)
                New-Item -Path $Path -ItemType File -Force | Out-Null
            }
            Mock -CommandName Get-Module -MockWith { [pscustomobject]@{ Name = 'ImportExcel' } }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Add-Type -MockWith { }

            & $scriptPath \
                -ExchangeUri 'http://exchange.contoso.com/PowerShell/' \
                -ConnectionType 'OnPrem' \
                -Credential $credential \
                -OutputPath $csvPath \
                -ExcelOutputPath $excelPath \
                -OrganizationSmtpSuffix 'contoso.com' \
                -ImpersonationSmtp 'service@contoso.com' \
                -SendInquiry \
                -NotificationFrom 'notify@contoso.com'

            Test-Path $csvPath | Should -BeTrue
            Test-Path $excelPath | Should -BeTrue

            Assert-MockCalled -CommandName Export-Csv -Times 1 -ParameterFilter { $Path -eq $csvPath }
            Assert-MockCalled -CommandName Export-Excel -Times 1 -ParameterFilter { $Path -eq $excelPath }
            Assert-MockCalled -CommandName Send-MailMessage -Times 1
            Assert-MockCalled -CommandName Get-RoomMeetings -Times 1
        }
    }
}
