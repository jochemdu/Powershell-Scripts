@{
    # Exchange connection settings
    Connection         = @{
        Type         = 'OnPrem'  # 'OnPrem' of 'EXO'
        ExchangeUri  = 'http://exchange.contoso.com/PowerShell/'
        EwsUrl       = $null     # Expliciet EWS URL (overslaat Autodiscover)
        Autodiscover = $true
    }

    # Credentials (password file zoals legacy scripts)
    Credential         = @{
        Username     = 'svc_account'
        PasswordFile = 'C:\ScheduledTasks\AfasLeaveData\password.txt'
    }

    # EWS impersonation settings
    Impersonation      = @{
        SmtpAddress = 'service.account@contoso.com'
    }

    # Integration Bus API endpoints (vervangt directe AFAS connectie)
    Api                = @{
        LeaveDataEndpoint     = 'https://integrationbus.contoso.com:7843/api/v1/getLeaveData'
        CanceledLeaveEndpoint = 'https://integrationbus.contoso.com:7843/api/v1/getCanceledLeaveData'
        ProxyUrl              = $null  # 'http://proxy.contoso.com:8080'
    }

    # Kalender item instellingen (zoals legacy scripts)
    LeaveSettings      = @{
        Subject = 'MyPlace Leave Booking (automatically created)'
        Body    = 'Please do not modify this appointment'
    }

    # Mapping van ITCode/medewerker naar Exchange mailbox
    UserMapping        = @{
        Strategy           = 'Mailbox'  # 'Mailbox' (Get-Mailbox), 'Email', 'MappingTable'
        DefaultDomain      = 'contoso.com'
        UseActiveDirectory = $false
        MappingTable       = @{
            'EMP001' = 'jan.janssen@contoso.com'
            'EMP002' = 'piet.pietersen@contoso.com'
        }
    }

    # Calendar event settings
    Calendar           = @{
        ShowAs   = 'OOF'       # 'Free', 'Tentative', 'Busy', 'OOF', 'WorkingElsewhere'
        Reminder = $false
    }

    # Sync behavior
    Sync               = @{
        Mode      = 'Delta'    # 'Full' of 'Delta'
        DaysAhead = 90
        DaysBehind = 0
    }

    # File paths (zoals legacy scripts)
    Paths              = @{
        ScriptPath      = 'C:\ScheduledTasks\AfasLeaveData'
        ProcessedPath   = 'C:\ScheduledTasks\AfasLeaveData\Processed'
        LogPath         = 'C:\ScheduledTasks\AfasLeaveData\Logs'
        OutputPath      = './reports/afas-leave-sync-report.csv'
        ExcelOutputPath = $null
    }

    # EWS assembly path
    EwsAssemblyPath    = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    # Verbose output
    Verbose            = $false
}
