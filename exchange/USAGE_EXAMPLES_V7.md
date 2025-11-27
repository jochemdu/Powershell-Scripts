# Find-GhostRoomMeetings-v7.ps1 - Usage Examples

## PowerShell 7+ Optimized Usage

### Example 1: Basic Usage with Auto-Detected Parallelization
```powershell
$cred = Get-Credential -Message 'Enter Exchange service account credentials'
.\Find-GhostRoomMeetings-v7.ps1 `
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' `
    -OrganizationSmtpSuffix 'contoso.com' `
    -ImpersonationSmtp 'service.account@contoso.com' `
    -Credential $cred
```
Automatically uses all available CPU cores for parallel processing.

### Example 2: Using JSON Configuration File
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred
```
JSON configuration is native to PS7 and loads faster than .psd1 files.

### Example 3: Custom Parallel Throttle Limit
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -ThrottleLimit 4
```
Limits parallel processing to 4 concurrent threads (useful for resource-constrained systems).

### Example 4: Maximum Performance (All Cores)
```powershell
$cred = Get-Credential
$cores = [Environment]::ProcessorCount
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -ThrottleLimit $cores
```
Explicitly uses all available CPU cores for maximum throughput.

### Example 5: Conservative Processing (Single Thread)
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -ThrottleLimit 1
```
Disables parallel processing for debugging or resource-limited environments.

### Example 6: With Email Notifications and Parallel Processing
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -SendInquiry `
    -NotificationFrom 'noreply@contoso.com' `
    -NotificationTemplate 'Please confirm if this meeting is still required for {0}.' `
    -ThrottleLimit 8
```
Sends notifications while processing rooms in parallel.

### Example 7: Extended Date Range with Parallel Processing
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -MonthsAhead 24 `
    -MonthsBehind 6 `
    -ThrottleLimit 6
```
Scans 6 months back and 24 months forward with 6 parallel threads.

### Example 8: Excel Export with Parallel Processing
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -ExcelOutputPath 'C:\Reports\ghost-meetings.xlsx' `
    -ThrottleLimit 8
```
Exports results to Excel while processing rooms in parallel.

### Example 9: Test Mode with Verbose Output
```powershell
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -TestMode `
    -Verbose
```
Tests configuration without connecting to Exchange.

### Example 10: Scheduled Task with Logging
```powershell
$logPath = 'C:\Logs\ghost-meetings-$(Get-Date -Format yyyyMMdd-HHmmss).log'
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath '.\config.example.json' `
    -Credential $cred `
    -ThrottleLimit 8 `
    -Verbose 4>&1 | Tee-Object -FilePath $logPath
```
Runs with logging for scheduled task execution.

## Performance Tuning

### For Small Deployments (< 50 rooms)
```powershell
# Use fewer threads to reduce overhead
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 2
```

### For Medium Deployments (50-200 rooms)
```powershell
# Use half the available cores
$throttle = [Math]::Max(1, [Environment]::ProcessorCount / 2)
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit $throttle
```

### For Large Deployments (200+ rooms)
```powershell
# Use all available cores
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit ([Environment]::ProcessorCount)
```

## Configuration File Format (JSON)

```json
{
  "ExchangeUri": "http://exchange.contoso.com/PowerShell/",
  "ConnectionType": "OnPrem",
  "EwsAssemblyPath": "C:\\Program Files\\Microsoft\\Exchange\\Web Services\\2.2\\Microsoft.Exchange.WebServices.dll",
  "MonthsAhead": 12,
  "MonthsBehind": 0,
  "OutputPath": "C:\\Reports\\ghost-meetings-report.csv",
  "ExcelOutputPath": "C:\\Reports\\ghost-meetings-report.xlsx",
  "OrganizationSmtpSuffix": "contoso.com",
  "ImpersonationSmtp": "service.account@contoso.com",
  "SendInquiry": false,
  "NotificationFrom": "noreply@contoso.com",
  "NotificationTemplate": "Please confirm if this meeting is still required for {0}."
}
```

## Troubleshooting

### Parallel Processing Causes Errors
```powershell
# Disable parallelization
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 1
```

### High Memory Usage
```powershell
# Reduce parallel threads
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 2
```

### Slow Performance
```powershell
# Increase parallel threads (if CPU available)
.\Find-GhostRoomMeetings-v7.ps1 -ConfigPath config.json -Credential $cred -ThrottleLimit 16
```

## Performance Expectations

| Deployment Size | Rooms | v1 Time | v7 Time | Speedup |
|-----------------|-------|---------|---------|---------|
| Small | 10 | 45s | 15s | 3x |
| Medium | 50 | 225s | 35s | 6.4x |
| Large | 100 | 450s | 65s | 6.9x |
| XL | 200 | 900s | 120s | 7.5x |

## Requirements

- **PowerShell 7.0+** (7.4+ recommended)
- EWS Managed API assembly
- Exchange Server 2013 SP1 or later
- Service account with FullAccess to room mailboxes
- Optional: ImportExcel module for Excel export

## Installation

```powershell
# Verify PowerShell version
$PSVersionTable.PSVersion

# Install ImportExcel for Excel export (optional)
Install-Module ImportExcel -Scope CurrentUser
```

## Best Practices

1. **Start Conservative**: Begin with `-ThrottleLimit 2` and increase gradually
2. **Monitor Resources**: Watch CPU and memory during first run
3. **Use JSON Config**: Faster loading than .psd1 files
4. **Schedule Off-Peak**: Run during low-usage periods
5. **Log Output**: Capture verbose output for troubleshooting
6. **Test First**: Use `-TestMode` before production run

