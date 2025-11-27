# Quick Reference Guide

## Version Selection

| Need | Version | Command |
|------|---------|---------|
| Legacy PowerShell | v1 | `.\Find-GhostRoomMeetings.ps1` |
| PowerShell 7+ | v7 | `.\Find-GhostRoomMeetings-v7.ps1` |
| Maximum Performance | v7 | `.\Find-GhostRoomMeetings-v7.ps1` |
| Maximum Compatibility | v1 | `.\Find-GhostRoomMeetings.ps1` |

## Basic Usage

### v1 (Universal)
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath config.example.psd1 `
    -Credential $cred
```

### v7 (Modern)
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath config.example.json `
    -Credential $cred
```

## Common Tasks

### Scan Last 3 Months
```powershell
# v1
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.psd1 -Credential $cred -MonthsBehind 3

# v7
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -MonthsBehind 3
```

### Export to Excel
```powershell
# v1
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.psd1 -Credential $cred -ExcelOutputPath report.xlsx

# v7
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ExcelOutputPath report.xlsx
```

### Send Notifications
```powershell
# v1
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.psd1 -Credential $cred `
    -SendInquiry -NotificationFrom noreply@contoso.com

# v7
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred `
    -SendInquiry -NotificationFrom noreply@contoso.com
```

### Test Mode
```powershell
# v1
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.psd1 -TestMode -Verbose

# v7
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -TestMode -Verbose
```

## Performance Tuning (v7 Only)

### Auto-Detect CPU Cores (Default)
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred
```

### Conservative (2 Threads)
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 2
```

### Aggressive (All Cores)
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit ([Environment]::ProcessorCount)
```

### Sequential (Debugging)
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 1
```

## Configuration Files

### v1 Format (.psd1)
```powershell
@{
    ExchangeUri = 'http://exchange.contoso.com/PowerShell/'
    ConnectionType = 'OnPrem'
    MonthsAhead = 12
    OrganizationSmtpSuffix = 'contoso.com'
}
```

### v7 Format (.json)
```json
{
  "ExchangeUri": "http://exchange.contoso.com/PowerShell/",
  "ConnectionType": "OnPrem",
  "MonthsAhead": 12,
  "OrganizationSmtpSuffix": "contoso.com"
}
```

## Troubleshooting

### Issue: "EWS assembly not found"
```powershell
# Update EwsAssemblyPath in config file
# Or install EWS Managed API from Microsoft
```

### Issue: "Cannot connect to Exchange"
```powershell
# Verify credentials
$cred = Get-Credential
# Test connection
Test-Connection exchange.contoso.com
```

### Issue: Parallel processing errors (v7)
```powershell
# Disable parallelization
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 1
```

### Issue: High memory usage (v7)
```powershell
# Reduce parallel threads
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 2
```

## Performance Comparison

| Deployment | v1 | v7 | Speedup |
|------------|----|----|---------|
| 10 rooms | 45s | 15s | 3x |
| 50 rooms | 225s | 35s | 6.4x |
| 100 rooms | 450s | 65s | 6.9x |

## Key Parameters

| Parameter | Purpose | Example |
|-----------|---------|---------|
| ConfigPath | Configuration file | `-ConfigPath config.json` |
| Credential | Exchange credentials | `-Credential $cred` |
| MonthsAhead | Future months to scan | `-MonthsAhead 12` |
| MonthsBehind | Past months to scan | `-MonthsBehind 3` |
| OutputPath | CSV output file | `-OutputPath report.csv` |
| ExcelOutputPath | Excel output file | `-ExcelOutputPath report.xlsx` |
| SendInquiry | Send notifications | `-SendInquiry` |
| NotificationFrom | Sender email | `-NotificationFrom noreply@contoso.com` |
| ThrottleLimit | Parallel threads (v7) | `-ThrottleLimit 8` |
| TestMode | Test without connecting | `-TestMode` |
| Verbose | Detailed output | `-Verbose` |

## Documentation

- **PS7_FEATURES.md** - PS7 features explained
- **VERSION_COMPARISON.md** - v1 vs v7 comparison
- **USAGE_EXAMPLES.md** - v1 examples
- **USAGE_EXAMPLES_V7.md** - v7 examples
- **PS7_MIGRATION_GUIDE.md** - Migration instructions
- **README.md** - Overview and quick start

## Requirements

### v1
- PowerShell 1.0+
- EWS Managed API
- Exchange Server 2013+

### v7
- PowerShell 7.0+
- EWS Managed API
- Exchange Server 2013+

## Installation

```powershell
# Copy scripts
Copy-Item Find-GhostRoomMeetings*.ps1 C:\Scripts\

# Copy config
Copy-Item config.example.* C:\Scripts\

# Install optional modules
Install-Module ImportExcel -Scope CurrentUser
```

## First Run Checklist

- [ ] Verify PowerShell version
- [ ] Install EWS Managed API
- [ ] Create configuration file
- [ ] Get service account credentials
- [ ] Run test mode: `-TestMode`
- [ ] Review output
- [ ] Run production
- [ ] Monitor performance
- [ ] Adjust throttle limit (v7)

## Support

For detailed information, see:
1. README.md - Overview
2. PS7_FEATURES.md - Feature details
3. USAGE_EXAMPLES_V7.md - Practical examples
4. PS7_MIGRATION_GUIDE.md - Migration help

