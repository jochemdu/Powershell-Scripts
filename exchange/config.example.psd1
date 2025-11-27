# Example PowerShell Data File (.psd1) for Find-GhostRoomMeetings.ps1
# 
# IMPORTANT: Do NOT store credentials in this file!
# Credentials must be provided via -Credential parameter or Get-Credential prompt.
#
# Usage:
#   $cred = Get-Credential
#   .\Find-GhostRoomMeetings.ps1 -ConfigPath .\config.example.psd1 -Credential $cred

@{
    # Exchange connection settings
    ExchangeUri = 'http://exchange.contoso.com/PowerShell/'
    ConnectionType = 'OnPrem'  # Options: 'Auto', 'OnPrem', 'EXO'
    
    # EWS configuration
    EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    EwsUrl = $null  # Optional: explicit EWS URL for Autodiscover bypass
    
    # Calendar scan window
    MonthsAhead = 12
    MonthsBehind = 0
    
    # Output settings
    OutputPath = 'C:\Reports\ghost-meetings-report.csv'
    ExcelOutputPath = 'C:\Reports\ghost-meetings-report.xlsx'
    
    # Organization settings
    OrganizationSmtpSuffix = 'contoso.com'
    ImpersonationSmtp = 'service.account@contoso.com'
    
    # Notification settings (optional)
    SendInquiry = $false
    NotificationFrom = 'noreply@contoso.com'
    NotificationTemplate = 'Please confirm if this meeting is still required for {0}.'
}

