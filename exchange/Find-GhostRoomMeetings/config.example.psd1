# Find-GhostRoomMeetings Configuration
# Copy this file to config.psd1 and update values for your environment.
# NOTE: Do not store credentials or secrets in this file.

@{
    # Connection settings
    Connection         = @{
        # Type: Auto, OnPrem, or EXO
        Type        = 'OnPrem'

        # Exchange PowerShell endpoint (On-Prem only)
        ExchangeUri = 'http://exchange.contoso.com/PowerShell/'

        # Explicit EWS URL (optional, skips Autodiscover if set)
        EwsUrl      = $null

        # Use Autodiscover for EWS endpoint
        Autodiscover = $true

        # Skip SSL certificate validation (for self-signed or untrusted certs)
        SkipCertificateCheck = $false

        # Use local Exchange snap-in instead of remote PowerShell
        LocalSnapin = $false
    }

    # Impersonation settings
    Impersonation      = @{
        # SMTP address for EWS impersonation
        SmtpAddress = 'service.account@contoso.com'
    }

    # Path to EWS Managed API assembly
    EwsAssemblyPath    = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    # Organization email domain suffix
    OrganizationSmtpSuffix = 'contoso.com'

    # Date range settings
    MonthsAhead        = 12
    MonthsBehind       = 0

    # Report output paths
    # Timestamp (yyyyMMdd-HHmmss) is automatically appended to filenames
    # Example: ghost-meetings-report_20251204-143022.csv
    OutputPath         = './reports/ghost-meetings-report.csv'
    ExcelOutputPath    = './reports/ghost-meetings-report.xlsx'

    # Notification settings
    SendInquiry        = $false
    NotificationFrom   = 'noreply@contoso.com'
    NotificationTemplate = 'Please confirm if this meeting is still required for {0}.'

    # Processing settings (PS7+ only)
    ThrottleLimit      = 4
}
