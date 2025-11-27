# Exchange Room Mailbox Auditing Scripts

Scripts voor Exchange Server en Exchange Online beheer.

## Directory Structure

Elk script heeft een eigen subdirectory conform [AGENTS.md](AGENTS.md):

```
exchange/
├── Find-GhostRoomMeetings/           # Ghost meeting detection
│   ├── Find-GhostRoomMeetings.ps1    # Main script
│   ├── config.example.json           # JSON configuration
│   ├── config.example.psd1           # PowerShell configuration
│   ├── README.md                     # Script documentation
│   └── modules/
│       └── ExchangeCore/             # Shared functions
├── Find-UnderutilizedRoomBookings/   # Underutilized room detection
├── AGENTS.md                         # Coding standards
└── README.md                         # This file
```

## Available Scripts

| Script | Purpose | PowerShell | Documentation |
|--------|---------|------------|---------------|
| [Find-GhostRoomMeetings](Find-GhostRoomMeetings/) | Detect meetings with missing/disabled organizers | 5.1+ | [README](Find-GhostRoomMeetings/README.md) |
| [Find-UnderutilizedRoomBookings](Find-UnderutilizedRoomBookings/) | Find large rooms booked for few attendees | 5.1+ | [README](Find-UnderutilizedRoomBookings/README.md) |

## Quick Start

### Find-GhostRoomMeetings

```powershell
.\Find-GhostRoomMeetings\Find-GhostRoomMeetings.ps1 `
    -ConfigPath .\Find-GhostRoomMeetings\config.json `
    -Credential (Get-Credential)
```

### Find-UnderutilizedRoomBookings

```powershell
.\Find-UnderutilizedRoomBookings\Find-UnderutilizedRoomBookings.ps1 `
    -ConfigPath .\Find-UnderutilizedRoomBookings\config.json `
    -Credential (Get-Credential)
```

## Shared Module

Both scripts use the `ExchangeCore` module located in each script's `modules/` directory:

- `Import-ConfigurationFile` - Load JSON/PSD1 configs
- `Connect-ExchangeSession` / `Disconnect-ExchangeSession` - Exchange connections
- `Connect-EwsService` - EWS service setup
- `Get-RoomCalendarItems` - Retrieve calendar meetings
- `Get-ResolvedConnectionType` - Auto-detect OnPrem/EXO

## Additional Documentation

- **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)** - Common usage examples
- **[QUICK_REFERENCE.md](QUICK_REFERENCE.md)** - Parameter quick reference
