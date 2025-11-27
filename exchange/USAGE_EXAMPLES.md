# Exchange Scripts - Usage Examples

## Find-GhostRoomMeetings

### Example 1: Basic Usage with Credential Prompt
```powershell
cd Find-GhostRoomMeetings
.\Find-GhostRoomMeetings.ps1 `
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' `
    -OrganizationSmtpSuffix 'contoso.com' `
    -ImpersonationSmtp 'service.account@contoso.com'
```

### Example 2: Using Configuration File
```powershell
$cred = Get-Credential -Message 'Enter Exchange service account credentials'
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.json' `
    -Credential $cred
```

### Example 3: Test Mode (No Exchange Connection)
```powershell
.\Find-GhostRoomMeetings.ps1 -TestMode -Verbose
```

### Example 4: With Email Notifications
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.json' `
    -Credential $cred `
    -SendInquiry `
    -NotificationFrom 'noreply@contoso.com'
```

### Example 5: Custom Date Range
```powershell
.\Find-GhostRoomMeetings.ps1 `
    -Credential (Get-Credential) `
    -MonthsAhead 6 `
    -MonthsBehind 1
```

### Example 6: Excel Export
```powershell
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath '.\config.json' `
    -Credential (Get-Credential) `
    -ExcelOutputPath './reports/ghost-meetings.xlsx'
```

---

## Find-UnderutilizedRoomBookings

### Example 1: Find Large Rooms with Few Attendees
```powershell
cd Find-UnderutilizedRoomBookings
.\Find-UnderutilizedRoomBookings.ps1 `
    -MinimumCapacity 6 `
    -MaxParticipants 2 `
    -Credential (Get-Credential)
```

### Example 2: Using Configuration File
```powershell
.\Find-UnderutilizedRoomBookings.ps1 `
    -ConfigPath '.\config.json' `
    -Credential (Get-Credential)
```

### Example 3: Strict Thresholds (10+ seats, 1 person)
```powershell
.\Find-UnderutilizedRoomBookings.ps1 `
    -MinimumCapacity 10 `
    -MaxParticipants 1 `
    -Credential (Get-Credential)
```

### Example 4: Extended Date Range
```powershell
.\Find-UnderutilizedRoomBookings.ps1 `
    -MonthsAhead 3 `
    -MonthsBehind 1 `
    -Credential (Get-Credential) `
    -OutputPath './reports/underutilized-q1.csv'
```

---

## Configuration File Format

Both scripts support JSON configuration:

```json
{
  "Connection": {
    "Type": "OnPrem",
    "ExchangeUri": "http://exchange.contoso.com/PowerShell/",
    "EwsUrl": null
  },
  "Impersonation": {
    "SmtpAddress": "service.account@contoso.com"
  },
  "OrganizationSmtpSuffix": "contoso.com",
  "MonthsAhead": 12,
  "MonthsBehind": 0
}
```

## Important Notes

- **No Secrets in Config**: Credentials are NEVER stored in configuration files
- **Credential Parameter Required**: Always provide `-Credential` explicitly
- **EWS Assembly Required**: Ensure EWS Managed API is installed
- **Service Account Needed**: Requires account with ApplicationImpersonation rights

