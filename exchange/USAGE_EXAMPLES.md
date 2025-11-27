# Find-GhostRoomMeetings.ps1 - Usage Examples

## PowerShell 1.0+ Compatible Usage

### Example 1: Basic Usage with Credential Prompt
```powershell
.\Find-GhostRoomMeetings.ps1 `
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' `
    -OrganizationSmtpSuffix 'contoso.com' `
    -ImpersonationSmtp 'service.account@contoso.com'
```
The script will prompt for credentials via `Get-Credential`.

### Example 2: Using Configuration File with Explicit Credential
```powershell
$cred = Get-Credential -Message 'Enter Exchange service account credentials'
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.example.psd1' `
    -Credential $cred
```

### Example 3: Test Mode (No Exchange Connection)
```powershell
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.example.psd1' `
    -TestMode
```
Useful for testing configuration without connecting to Exchange.

### Example 4: With Email Notifications
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.example.psd1' `
    -Credential $cred `
    -SendInquiry `
    -NotificationFrom 'noreply@contoso.com' `
    -NotificationTemplate 'Please confirm if this meeting is still required for {0}.'
```

### Example 5: Custom Date Range
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.example.psd1' `
    -Credential $cred `
    -MonthsAhead 6 `
    -MonthsBehind 3
```
Scans meetings from 3 months ago to 6 months in the future.

### Example 6: Excel Export
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.example.psd1' `
    -Credential $cred `
    -ExcelOutputPath 'C:\Reports\ghost-meetings.xlsx'
```
Requires ImportExcel module: `Install-Module ImportExcel`

## Configuration File Format

Create a `.psd1` file (PowerShell Data File):

```powershell
@{
    ExchangeUri = 'http://exchange.contoso.com/PowerShell/'
    ConnectionType = 'OnPrem'
    EwsAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    MonthsAhead = 12
    MonthsBehind = 0
    OutputPath = 'C:\Reports\ghost-meetings-report.csv'
    OrganizationSmtpSuffix = 'contoso.com'
    ImpersonationSmtp = 'service.account@contoso.com'
}
```

## Important Notes

- **No Secrets in Config**: Credentials are NEVER stored in configuration files
- **Credential Parameter Required**: Always provide `-Credential` explicitly or use `Get-Credential`
- **PowerShell 1.0+ Compatible**: Works with all PowerShell versions
- **EWS Assembly Required**: Ensure EWS Managed API is installed
- **Service Account Needed**: Requires account with FullAccess to room mailboxes

