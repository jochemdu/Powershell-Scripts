# Find-GhostRoomMeetings Configuration
# Copy this file to config.psd1 and update values for your environment.
# NOTE: Do not store credentials or secrets in this file.
#
# USAGE FROM EXCHANGE MANAGEMENT SHELL:
#   .\Find-GhostRoomMeetings.ps1 -LocalSnapin -ConfigPath .\config.psd1 -Verbose
#
# Or without config file:
#   .\Find-GhostRoomMeetings.ps1 -LocalSnapin -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx -OrganizationSmtpSuffix contoso.com -Verbose

@{
    # Connection settings
    Connection         = @{
        # Type: Auto, OnPrem, or EXO
        Type        = 'OnPrem'

        # Exchange PowerShell endpoint (On-Prem only, ignored when using LocalSnapin)
        ExchangeUri = 'http://exchange.contoso.com/PowerShell/'

        # Explicit EWS URL (recommended for LocalSnapin mode, skips Autodiscover)
        # Example: 'https://mail.contoso.com/EWS/Exchange.asmx'
        EwsUrl      = $null

        # Use Autodiscover for EWS endpoint (set to $false when using EwsUrl)
        Autodiscover = $true

        # Skip SSL certificate validation (for self-signed or untrusted certs)
        SkipCertificateCheck = $false

        # Use local Exchange snap-in instead of remote PowerShell
        # Set to $true when running on Exchange server or server with Exchange Management Tools
        LocalSnapin = $false
    }

    # Impersonation settings
    Impersonation      = @{
        # SMTP address for EWS impersonation (must have ApplicationImpersonation rights)
        SmtpAddress = 'service.account@contoso.com'
    }

    # Path to EWS Managed API assembly
    # Common locations:
    #   - C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll
    #   - D:\Scripts\dll\Microsoft.Exchange.WebServices.dll
    EwsAssemblyPath    = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    # Organization email domain suffix (primary internal domain)
    # Used to identify internal vs external organizers
    # Also used to match external addresses to internal users by local part
    # Example: 'user@external.contoso.com' will be matched to 'user@contoso.com'
    OrganizationSmtpSuffix = 'contoso.com'

    # Date range settings
    MonthsAhead        = 12
    MonthsBehind       = 0

    # Report output paths
    # Timestamp (yyyyMMdd-HHmmss) is automatically appended to filenames
    # Example: ghost-meetings-report-20251204-143022.csv
    OutputPath         = './reports/ghost-meetings-report.csv'
    ExcelOutputPath    = './reports/ghost-meetings-report.xlsx'

    # Notification settings (optional)
    SendInquiry        = $false
    NotificationFrom   = 'noreply@contoso.com'
    NotificationTemplate = 'Please confirm if this meeting is still required for {0}.'

    # Processing settings (PS7+ only)
    ThrottleLimit      = 4
}
