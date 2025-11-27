# Find-UnderutilizedRoomBookings

Detects room mailbox meetings where large rooms are booked for few participants, helping identify inefficient room utilization.

## Features

- Supports both Exchange On-Premises and Exchange Online
- Scans room mailbox calendars via EWS
- Filters rooms by minimum capacity threshold
- Identifies meetings with few participants in large rooms
- Exports results to CSV
- Configuration via JSON or PowerShell Data File (PSD1)
- Compatible with PowerShell 5.1+

## Requirements

### PowerShell

- **Minimum**: PowerShell 5.1
- **Recommended**: PowerShell 7.0+

### Modules & Assemblies

| Component | Purpose | Install Command |
|-----------|---------|-----------------|
| EWS Managed API | Calendar access via Exchange Web Services | [Download from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=42951) |
| ExchangeOnlineManagement | Exchange Online connections | `Install-Module ExchangeOnlineManagement` |
| ExchangeCore module | Shared functions (included) | Built-in |

### Required Rights

| Right | Scope | Purpose |
|-------|-------|---------|
| `ApplicationImpersonation` | Room mailboxes | EWS calendar access |
| Exchange PowerShell access | Organization | Room mailbox enumeration |

## Directory Structure

```
Find-UnderutilizedRoomBookings/
├── Find-UnderutilizedRoomBookings.ps1  # Main script
├── config.example.json                  # JSON configuration template
├── config.example.psd1                  # PowerShell config template
├── README.md                            # This file
└── modules/
    └── ExchangeCore/
        └── ExchangeCore.psm1            # Shared functions module
```

## Configuration

Copy `config.example.json` to `config.json` and update values:

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
  "MinimumCapacity": 6,
  "MaxParticipants": 2,
  "OutputPath": "./reports/underutilized-room-bookings.csv"
}
```

### Configuration Options

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Connection.Type` | string | Auto | `Auto`, `OnPrem`, or `EXO` |
| `Connection.ExchangeUri` | string | - | Exchange PowerShell endpoint |
| `Connection.EwsUrl` | string | null | Explicit EWS URL (skips Autodiscover) |
| `Impersonation.SmtpAddress` | string | - | SMTP for EWS impersonation |
| `MonthsAhead` | int | 1 | Months ahead to scan |
| `MonthsBehind` | int | 0 | Months behind to scan |
| `MinimumCapacity` | int | 6 | Only check rooms with this capacity or higher |
| `MaxParticipants` | int | 2 | Flag meetings with this many or fewer participants |
| `OutputPath` | string | - | CSV report path |

> **Security Note**: Do not store credentials in config files. Use `-Credential` parameter or secure retrieval methods.

## Usage Examples

### Basic Scan (Interactive Credentials)

```powershell
.\Find-UnderutilizedRoomBookings.ps1 -Credential (Get-Credential) -Verbose
```

### Using Configuration File

```powershell
.\Find-UnderutilizedRoomBookings.ps1 -ConfigPath ./config.json
```

### Custom Thresholds

```powershell
# Find 8+ seat rooms with only 1-3 attendees
.\Find-UnderutilizedRoomBookings.ps1 `
    -Credential (Get-Credential) `
    -MinimumCapacity 8 `
    -MaxParticipants 3
```

### Exchange Online

```powershell
.\Find-UnderutilizedRoomBookings.ps1 `
    -ConnectionType EXO `
    -ImpersonationSmtp service@contoso.onmicrosoft.com `
    -MinimumCapacity 6 `
    -MaxParticipants 2
```

### Extended Date Range

```powershell
.\Find-UnderutilizedRoomBookings.ps1 `
    -Credential (Get-Credential) `
    -MonthsAhead 3 `
    -MonthsBehind 1 `
    -OutputPath ./reports/q1-underutilized.csv
```

### Test Mode (No Connections)

```powershell
.\Find-UnderutilizedRoomBookings.ps1 -TestMode -Verbose
```

## Output

### CSV Report Columns

| Column | Description |
|--------|-------------|
| Room | Room mailbox SMTP address |
| DisplayName | Room display name |
| Capacity | Room capacity (seats) |
| Subject | Meeting subject |
| Start | Meeting start time |
| End | Meeting end time |
| Organizer | Meeting organizer SMTP |
| ParticipantCount | Number of distinct participants |
| Participants | Semicolon-separated participant list |
| UniqueId | EWS item unique identifier |

### Example Output

| Room | DisplayName | Capacity | Subject | ParticipantCount |
|------|-------------|----------|---------|------------------|
| conf-large@contoso.com | Large Conference Room | 12 | 1:1 Meeting | 2 |
| boardroom@contoso.com | Executive Boardroom | 20 | Quick Sync | 1 |

## Use Cases

### Identifying Wasted Capacity

```powershell
# Find meetings in 10+ seat rooms with just 1 person
.\Find-UnderutilizedRoomBookings.ps1 `
    -MinimumCapacity 10 `
    -MaxParticipants 1
```

### Monthly Utilization Reports

```powershell
# Generate report for past month
.\Find-UnderutilizedRoomBookings.ps1 `
    -MonthsAhead 0 `
    -MonthsBehind 1 `
    -OutputPath "./reports/$(Get-Date -Format 'yyyy-MM')-underutilized.csv"
```

## Troubleshooting

### Common Issues

**"No room mailboxes found with capacity >= X"**
- Verify rooms have `ResourceCapacity` set in Exchange
- Lower the `-MinimumCapacity` threshold

**"EWS assembly not found"**
- Download and install EWS Managed API
- Verify `-EwsAssemblyPath` points to correct location

**"Access denied to calendar"**
- Verify service account has `ApplicationImpersonation` role
- Check impersonation SMTP address

### Verbose Logging

Use `-Verbose` for detailed progress:

```powershell
.\Find-UnderutilizedRoomBookings.ps1 -ConfigPath ./config.json -Verbose
```

## Related Scripts

- [Find-GhostRoomMeetings](../Find-GhostRoomMeetings/) - Identify meetings with missing/disabled organizers

## License

See repository LICENSE file.
