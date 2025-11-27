# PowerShell Scripts Repository

PowerShell-scripts voor Exchange en andere Microsoft-omgevingen, gestructureerd per domein.

## Directory Structure

```
├── exchange/                    # Exchange Server/Online scripts
│   ├── Find-GhostRoomMeetings/     # Ghost meeting detection
│   └── Find-UnderutilizedRoomBookings/  # Room utilization analysis
├── modules/                     # Shared PowerShell modules
├── tests/                       # Pester tests per domain
├── bash/                        # Shell utility scripts
└── .github/workflows/           # CI/CD pipelines
```

## Available Scripts

| Domain | Script | Purpose | Docs |
|--------|--------|---------|------|
| Exchange | [Find-GhostRoomMeetings](exchange/Find-GhostRoomMeetings/) | Detect meetings with missing/disabled organizers | [README](exchange/Find-GhostRoomMeetings/README.md) |
| Exchange | [Find-UnderutilizedRoomBookings](exchange/Find-UnderutilizedRoomBookings/) | Find large rooms booked for few attendees | [README](exchange/Find-UnderutilizedRoomBookings/README.md) |

## Quick Start

### Find Ghost Meetings

```powershell
cd exchange/Find-GhostRoomMeetings
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath .\config.json `
    -Credential (Get-Credential)
```

### Find Underutilized Room Bookings

```powershell
cd exchange/Find-UnderutilizedRoomBookings
.\Find-UnderutilizedRoomBookings.ps1 `
    -MinimumCapacity 6 `
    -MaxParticipants 2 `
    -Credential (Get-Credential)
```

## Requirements

- **PowerShell**: 5.1+ (7.0+ recommended)
- **EWS Managed API**: For calendar access
- **Exchange**: On-premises or Online with appropriate permissions

See individual script READMEs for specific requirements.

## Tests

Pester tests are in the [`tests/`](tests/) directory:

```powershell
# Run all tests
Invoke-Pester -Path tests

# Run exchange tests only
Invoke-Pester -Path tests/exchange
```

## Documentation

- [`AGENTS.md`](AGENTS.md) - Repository coding standards and structure guidelines
- [`exchange/AGENTS.md`](exchange/AGENTS.md) - Exchange-specific guidelines
- Individual script READMEs for detailed usage

## Contributing

See [`AGENTS.md`](AGENTS.md) for coding standards:
- Use Verb-Noun naming convention
- Include `[CmdletBinding()]` and proper help blocks
- One functional change per commit
- Update relevant documentation
