# Exchange Scripts - Quick Reference

## Available Scripts

| Script | Purpose | Location |
|--------|---------|----------|
| Find-GhostRoomMeetings | Detect ghost meetings | `Find-GhostRoomMeetings/` |
| Find-UnderutilizedRoomBookings | Find underutilized rooms | `Find-UnderutilizedRoomBookings/` |

## Common Tasks

### Scan for Ghost Meetings
```powershell
cd Find-GhostRoomMeetings
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.json -Credential (Get-Credential)
```

### Find Underutilized Room Bookings
```powershell
cd Find-UnderutilizedRoomBookings
.\Find-UnderutilizedRoomBookings.ps1 -MinimumCapacity 6 -MaxParticipants 2 -Credential (Get-Credential)
```

### Test Mode (No Connections)
```powershell
.\Find-GhostRoomMeetings.ps1 -TestMode -Verbose
.\Find-UnderutilizedRoomBookings.ps1 -TestMode -Verbose
```

### Export to Excel
```powershell
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.json -Credential $cred -ExcelOutputPath report.xlsx
```

### Send Notifications (Ghost Meetings)
```powershell
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.json -Credential $cred `
    -SendInquiry -NotificationFrom noreply@contoso.com
```

## Configuration Files

### JSON Format (Recommended)
```json
{
  "Connection": {
    "Type": "OnPrem",
    "ExchangeUri": "http://exchange.contoso.com/PowerShell/"
  },
  "Impersonation": {
    "SmtpAddress": "service@contoso.com"
  },
  "OrganizationSmtpSuffix": "contoso.com",
  "MonthsAhead": 12,
  "MonthsBehind": 0
}
```

### PowerShell Data File Format
```powershell
@{
    Connection = @{
        Type = 'OnPrem'
        ExchangeUri = 'http://exchange.contoso.com/PowerShell/'
    }
    Impersonation = @{
        SmtpAddress = 'service@contoso.com'
    }
    OrganizationSmtpSuffix = 'contoso.com'
}
```

## Key Parameters

### Common Parameters

| Parameter | Purpose | Example |
|-----------|---------|---------|
| ConfigPath | Configuration file | `-ConfigPath config.json` |
| Credential | Exchange credentials | `-Credential (Get-Credential)` |
| MonthsAhead | Future months to scan | `-MonthsAhead 12` |
| MonthsBehind | Past months to scan | `-MonthsBehind 1` |
| OutputPath | CSV output file | `-OutputPath report.csv` |
| TestMode | Test without connecting | `-TestMode` |
| Verbose | Detailed output | `-Verbose` |

### Ghost Meetings Specific

| Parameter | Purpose | Example |
|-----------|---------|---------|
| SendInquiry | Send notifications | `-SendInquiry` |
| NotificationFrom | Sender email | `-NotificationFrom noreply@contoso.com` |
| ExcelOutputPath | Excel output file | `-ExcelOutputPath report.xlsx` |

### Underutilized Rooms Specific

| Parameter | Purpose | Example |
|-----------|---------|---------|
| MinimumCapacity | Min room size | `-MinimumCapacity 6` |
| MaxParticipants | Max attendee threshold | `-MaxParticipants 2` |

## Troubleshooting

### Issue: "EWS assembly not found"
- Download EWS Managed API from Microsoft
- Update EwsAssemblyPath in config

### Issue: "Cannot connect to Exchange"
```powershell
# Verify credentials
$cred = Get-Credential
Test-Connection exchange.contoso.com
```

### Issue: "Access denied to calendar"
- Verify ApplicationImpersonation role
- Check ImpersonationSmtp address

### Issue: "No room mailboxes found"
- Verify Exchange connection
- Check room mailbox RecipientTypeDetails

## Requirements

- PowerShell 5.1+ (7.0+ recommended)
- EWS Managed API
- Exchange Server 2013+ or Exchange Online
- ApplicationImpersonation rights

## First Run Checklist

- [ ] Verify PowerShell version: `$PSVersionTable.PSVersion`
- [ ] Install EWS Managed API
- [ ] Copy config.example.json to config.json
- [ ] Update config with your environment settings
- [ ] Get service account credentials
- [ ] Run in test mode: `-TestMode -Verbose`
- [ ] Review output
- [ ] Run production scan

## Documentation

- [Find-GhostRoomMeetings README](Find-GhostRoomMeetings/README.md)
- [Find-UnderutilizedRoomBookings README](Find-UnderutilizedRoomBookings/README.md)
- [USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)
- [AGENTS.md](AGENTS.md) - Coding standards

