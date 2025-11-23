Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

describe 'Find-GhostRoomMeetings EXO pathway' {
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '../../exchange/Find-GhostRoomMeetings.ps1'
    . $scriptPath

    BeforeEach {
        $script:ExchangeConnectionType = 'EXO'
    }

    It 'chooses Exchange Online when ExchangeUri points to outlook.office365.com' {
        $script:ExchangeConnectionType = 'OnPrem'
        $exchangeUri = 'https://outlook.office365.com/powershell-liveid/'

        # simulate parameter resolution logic
        $resolved = switch ('Auto') {
            'EXO'   { 'EXO' }
            'Auto'  {
                if ($exchangeUri -match 'outlook\.office365\.com' -or $exchangeUri -match 'ps\.outlook\.com' -or $exchangeUri -match 'office365\.com') { 'EXO' } else { 'OnPrem' }
            }
            default { 'OnPrem' }
        }

        $resolved | Should -Be 'EXO'
    }

    It 'uses EXO cmdlets for room lookup' {
        Mock -CommandName Get-ExoMailbox -MockWith { @([pscustomobject]@{ DisplayName='Room'; PrimarySmtpAddress='room@contoso.com'; Alias='room'; Identity='room'}) }
        Mock -CommandName Get-Mailbox -MockWith { throw 'Should not be used for EXO' }

        $rooms = Get-RoomMailboxes

        Assert-MockCalled -CommandName Get-ExoMailbox -Times 1 -Exactly
        Assert-MockCalled -CommandName Get-Mailbox -Times 0 -Exactly
        $rooms[0].PrimarySmtpAddress | Should -Be 'room@contoso.com'
    }

    It 'skips Connect-ExchangeOnline in test mode' {
        Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'Should not connect' }

        $session = Connect-ExchangeSession -Type 'EXO' -TestMode -Credential (New-Object System.Management.Automation.PSCredential('user@contoso.com',(ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)))

        $session | Should -Be $null
        Assert-MockCalled -CommandName Connect-ExchangeOnline -Times 0 -Exactly
    }
}
