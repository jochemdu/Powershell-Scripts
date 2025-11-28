# Find-UnderutilizedRoomBookings Configuration
# Copy this file to config.psd1 and update values for your environment.
# NOTE: Do not store credentials or secrets in this file.

@{
    # Connection settings
    Connection        = @{
        # Type: Auto, OnPrem, or EXO
        Type                 = 'OnPrem'

        # Exchange PowerShell endpoint (On-Prem only)
        ExchangeUri          = 'http://exchange.contoso.com/PowerShell/'

        # Explicit EWS URL (optional, skips Autodiscover if set)
        EwsUrl               = $null

        # Use Autodiscover for EWS endpoint
        Autodiscover         = $true

        # Proxy server URL (optional, e.g., http://proxy.contoso.com:8080)
        ProxyUrl             = $null

        # Authentication method: Kerberos, Negotiate, Basic, Default
        # Use 'Negotiate' if you encounter Kerberos SPN errors
        Authentication       = 'Kerberos'

        # Skip SSL certificate validation (for self-signed or mismatched certs)
        SkipCertificateCheck = $false
    }

    # Impersonation settings
    Impersonation     = @{
        # SMTP address for EWS impersonation
        SmtpAddress = 'service.account@contoso.com'
    }

    # Path to EWS Managed API assembly
    EwsAssemblyPath   = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    # Date range settings
    MonthsAhead       = 1
    MonthsBehind      = 0

    # Underutilization thresholds
    MinimumCapacity   = 6   # Only check rooms with this capacity or higher
    MaxParticipants   = 2   # Flag meetings with this many or fewer participants

    # Report output path
    OutputPath        = './reports/underutilized-room-bookings.csv'
}
