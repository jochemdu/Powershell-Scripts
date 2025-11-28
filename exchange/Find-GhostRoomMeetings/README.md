# Find-GhostRoomMeetings

Audits room mailbox calendars to identify "ghost" meetings where the organizer is missing, disabled, or has left the organization.

## Features

- Supports both Exchange On-Premises and Exchange Online
- Scans room mailbox calendars via EWS
- Validates organizer status (Active, Disabled, NotFound, External)
- Exports results to CSV and Excel (XLSX)
- Optional email notifications to attendees of ghost meetings
- Configuration via JSON or PowerShell Data File (PSD1)
- Compatible with PowerShell 5.1+ (7.0+ recommended)

## Requirements

### PowerShell

- **Minimum**: PowerShell 5.1
- **Recommended**: PowerShell 7.0+ (for improved performance)

### Modules & Assemblies

| Component | Purpose | Install Command |
|-----------|---------|-----------------|
| EWS Managed API | Calendar access via Exchange Web Services | [Download from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=42951) |
| ImportExcel | Excel export (optional) | `Install-Module ImportExcel` |
| ExchangeOnlineManagement | Exchange Online connections | `Install-Module ExchangeOnlineManagement` |
| ActiveDirectory | On-prem organizer validation | Included with RSAT |

### Required Rights

| Right | Scope | Purpose |
|-------|-------|---------|
| `ApplicationImpersonation` | Room mailboxes | EWS calendar access |
| Exchange PowerShell access | Organization | Room mailbox enumeration |
| AD Read (on-prem) | Domain | Organizer enabled/disabled state |

## Directory Structure

```
Find-GhostRoomMeetings/
├── Find-GhostRoomMeetings.ps1   # Main script
├── config.example.json           # JSON configuration template
├── config.example.psd1           # PowerShell config template
├── README.md                     # This file
└── modules/
    └── ExchangeCore/
        └── ExchangeCore.psm1     # Shared functions module
```

## Configuration

Copy `config.example.json` to `config.json` and update values:

```json
{
  "Connection": {
    "Type": "OnPrem",
    "ExchangeUri": "http://exchange.contoso.com/PowerShell/",
    "EwsUrl": null,
    "Autodiscover": true
  },
  "Impersonation": {
    "SmtpAddress": "service.account@contoso.com"
  },
  "OrganizationSmtpSuffix": "contoso.com",
  "MonthsAhead": 12,
  "MonthsBehind": 0,
  "OutputPath": "./reports/ghost-meetings-report.csv",
  "ExcelOutputPath": "./reports/ghost-meetings-report.xlsx"
}
```

### Configuration Options

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Connection.Type` | string | Auto | `Auto`, `OnPrem`, or `EXO` |
| `Connection.ExchangeUri` | string | - | Exchange PowerShell endpoint |
| `Connection.EwsUrl` | string | null | Explicit EWS URL (skips Autodiscover) |
| `Impersonation.SmtpAddress` | string | - | SMTP for EWS impersonation |
| `OrganizationSmtpSuffix` | string | - | Email domain (e.g., `contoso.com`) |
| `MonthsAhead` | int | 12 | Months ahead to scan |
| `MonthsBehind` | int | 0 | Months behind to scan |
| `OutputPath` | string | - | CSV report path |
| `ExcelOutputPath` | string | - | Excel report path (optional) |
| `SendInquiry` | bool | false | Send attendee notifications |
| `NotificationFrom` | string | - | From address for notifications |
| `NotificationTemplate` | string | - | Email body template (`{0}` = subject) |
| `ThrottleLimit` | int | CPU count | Parallel processing threads (PS7+) |

> **Security Note**: Do not store credentials in config files. Use `-Credential` parameter or secure retrieval methods.

## Usage Examples

### Basic Scan (Interactive Credentials)

```powershell
.\Find-GhostRoomMeetings.ps1 -Credential (Get-Credential) -Verbose
```

### Using Configuration File

```powershell
.\Find-GhostRoomMeetings.ps1 -ConfigPath ./config.json
```

### Exchange Online

```powershell
.\Find-GhostRoomMeetings.ps1 `
    -ConnectionType EXO `
    -ImpersonationSmtp service@contoso.onmicrosoft.com `
    -OrganizationSmtpSuffix contoso.com `
    -MonthsAhead 6
```

### With Email Notifications

```powershell
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath ./config.json `
    -SendInquiry `
    -NotificationFrom admin@contoso.com
```

### Limited Date Range

```powershell
.\Find-GhostRoomMeetings.ps1 `
    -Credential (Get-Credential) `
    -MonthsAhead 3 `
    -MonthsBehind 0 `
    -OutputPath ./reports/q1-ghost-meetings.csv
```

### Test Mode (No Connections)

```powershell
.\Find-GhostRoomMeetings.ps1 -TestMode -Verbose
```

### Local Snap-in Mode (Run on Exchange Server)

When Remote PowerShell is blocked or returns 401 errors, run the script directly on the Exchange server using the local snap-in:

```powershell
# On the Exchange server - uses current Windows identity
.\Find-GhostRoomMeetings.ps1 `
    -LocalSnapin `
    -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx `
    -ImpersonationSmtp service.account@contoso.com `
    -OrganizationSmtpSuffix contoso.com `
    -Verbose

# With explicit credentials for EWS (e.g., service account)
.\Find-GhostRoomMeetings.ps1 `
    -LocalSnapin `
    -Credential (Get-Credential) `
    -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx `
    -OrganizationSmtpSuffix contoso.com `
    -Verbose
```

**Note**: LocalSnapin mode uses your current Windows identity by default. You can optionally provide `-Credential` to use different credentials for EWS.

## Output

### CSV Report Columns

| Column | Description |
|--------|-------------|
| Room | Room mailbox SMTP address |
| Subject | Meeting subject |
| Start | Meeting start time |
| End | Meeting end time |
| Organizer | Organizer SMTP address |
| OrganizerStatus | `Active`, `Disabled`, `NotFound`, or `External` |
| IsRecurring | Whether meeting is recurring |
| Attendees | Semicolon-separated attendee list |
| UniqueId | EWS item unique identifier |

### Organizer Status Values

| Status | Description |
|--------|-------------|
| `Active` | Organizer exists and is enabled |
| `Disabled` | Organizer account is disabled |
| `NotFound` | Organizer not found in directory |
| `External` | Organizer is outside the organization |

## Troubleshooting

### Common Issues

**"EWS assembly not found"**
- Download and install EWS Managed API
- Verify `-EwsAssemblyPath` points to correct location

**"Autodiscover failed"**
- Provide explicit EWS URL via `-EwsUrl` or config
- Verify DNS and network connectivity

**"Access denied to calendar"**
- Verify service account has `ApplicationImpersonation` role
- Check impersonation SMTP address

**"ActiveDirectory module not available"**
- Install RSAT tools for on-prem deployments
- Script will skip enabled/disabled check if unavailable

**"401 Unauthorized" or Remote PowerShell blocked**
- Try `-Authentication Negotiate` instead of Kerberos
- Use `-LocalSnapin` to run directly on the Exchange server
- Verify `RemotePowerShellEnabled` is `$true` for the service account

**"Kerberos SPN error" or certificate errors**
- Use `-Authentication Negotiate` for NTLM fallback
- Use `-SkipCertificateCheck` for self-signed or mismatched certificates
- Use `-LocalSnapin` to bypass Remote PowerShell entirely

### Verbose Logging

Use `-Verbose` for detailed progress:

```powershell
.\Find-GhostRoomMeetings.ps1 -ConfigPath ./config.json -Verbose
```

## Related Scripts

- [Find-UnderutilizedRoomBookings](../Find-UnderutilizedRoomBookings/) - Identify underutilized room bookings

## License

See repository LICENSE file.
