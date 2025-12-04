# Find-UnderutilizedRoomBookings Configuration
# Copy this file to config.psd1 and update values for your environment.
# NOTE: Do not store credentials or secrets in this file.
#
# USAGE FROM EXCHANGE MANAGEMENT SHELL:
#   .\Find-UnderutilizedRoomBookings.ps1 -LocalSnapin -ConfigPath .\config.psd1 -Verbose
#
# Or without config file:
#   .\Find-UnderutilizedRoomBookings.ps1 -LocalSnapin -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx -MinimumCapacity 6 -MaxParticipants 2 -Verbose

@{
    # Connection settings
    Connection        = @{
        # Type: Auto, OnPrem, or EXO
        Type                 = 'OnPrem'

        # Exchange PowerShell endpoint (On-Prem only, ignored when using LocalSnapin)
        ExchangeUri          = 'http://exchange.contoso.com/PowerShell/'

        # Explicit EWS URL (recommended for LocalSnapin mode, skips Autodiscover)
        # Example: 'https://mail.contoso.com/EWS/Exchange.asmx'
        EwsUrl               = $null

        # Use Autodiscover for EWS endpoint (set to $false when using EwsUrl)
        Autodiscover         = $true

        # Proxy server URL (optional, e.g., http://proxy.contoso.com:8080)
        ProxyUrl             = $null

        # Authentication method: Kerberos, Negotiate, Basic, Default
        # Use 'Negotiate' if you encounter Kerberos SPN errors
        Authentication       = 'Kerberos'

        # Skip SSL certificate validation (for self-signed or mismatched certs)
        SkipCertificateCheck = $false

        # Use local Exchange snap-in instead of remote PowerShell
        # Set to $true when running on Exchange server or server with Exchange Management Tools
        LocalSnapin          = $false
    }

    # Impersonation settings
    Impersonation     = @{
        # SMTP address for EWS impersonation (must have ApplicationImpersonation rights)
        SmtpAddress = 'service.account@contoso.com'
    }

    # Path to EWS Managed API assembly
    # Common locations:
    #   - C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll
    #   - D:\Scripts\dll\Microsoft.Exchange.WebServices.dll
    EwsAssemblyPath   = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    # Organization email domain suffix (primary internal domain)
    # Used to identify internal vs external organizers
    # Also used to match external addresses to internal users by local part
    # Example: 'user@external.contoso.com' will be matched to 'user@contoso.com'
    OrganizationSuffix = 'contoso.com'

    # Date range settings
    MonthsAhead       = 1
    MonthsBehind      = 0

    # Underutilization thresholds
    MinimumCapacity   = 6   # Only check rooms with this capacity or higher
    MaxParticipants   = 2   # Flag meetings with this many or fewer participants

    # Report output path
    # Timestamp (yyyyMMdd-HHmmss) is automatically appended to filenames
    # Example: underutilized-room-bookings-20251204-143022.csv
    OutputPath        = './reports/underutilized-room-bookings.csv'
}
