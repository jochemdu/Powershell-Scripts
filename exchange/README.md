# Exchange Room Mailbox Auditing Scripts

Scripts voor Exchange Server en Exchange Online beheer.

## Directory Structure

Elk script heeft een eigen subdirectory conform [AGENTS.md](AGENTS.md):

```
exchange/
├── AfasLeaveData/                    # AFAS leave data sync
│   ├── AfasLeaveData.ps1             # Main script
│   ├── config.example.json           # JSON configuration
│   ├── config.example.psd1           # PowerShell configuration
│   ├── README.md                     # Script documentation
│   ├── legacy/                       # Old scripts for refactoring
│   └── modules/
│       └── AfasCore/                 # AFAS/Service Bus functions
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
| [AfasLeaveData](AfasLeaveData/) | Sync AFAS leave data to Exchange calendars via Service Bus | 5.1+ | [README](AfasLeaveData/README.md) |
| [Find-GhostRoomMeetings](Find-GhostRoomMeetings/) | Detect meetings with missing/disabled organizers | 5.1+ | [README](Find-GhostRoomMeetings/README.md) |
| [Find-UnderutilizedRoomBookings](Find-UnderutilizedRoomBookings/) | Find large rooms booked for few attendees | 5.1+ | [README](Find-UnderutilizedRoomBookings/README.md) |

## Quick Start

### AfasLeaveData

```powershell
.\AfasLeaveData\AfasLeaveData.ps1 `
    -ConfigPath .\AfasLeaveData\config.json `
    -Credential (Get-Credential)

# Test mode (no actual connections)
.\AfasLeaveData\AfasLeaveData.ps1 -TestMode -Verbose
```

### Find-GhostRoomMeetings

```powershell
# With config file and credentials
.\Find-GhostRoomMeetings\Find-GhostRoomMeetings.ps1 `
    -ConfigPath .\Find-GhostRoomMeetings\config.psd1 `
    -Credential (Get-Credential)

# LocalSnapin mode (run on Exchange server or server with Exchange Management Tools)
.\Find-GhostRoomMeetings\Find-GhostRoomMeetings.ps1 `
    -LocalSnapin `
    -ConfigPath .\Find-GhostRoomMeetings\config.psd1 `
    -Verbose
```

### Find-UnderutilizedRoomBookings

```powershell
# With config file and credentials
.\Find-UnderutilizedRoomBookings\Find-UnderutilizedRoomBookings.ps1 `
    -ConfigPath .\Find-UnderutilizedRoomBookings\config.psd1 `
    -Credential (Get-Credential)

# LocalSnapin mode (run on Exchange server or server with Exchange Management Tools)
.\Find-UnderutilizedRoomBookings\Find-UnderutilizedRoomBookings.ps1 `
    -LocalSnapin `
    -ConfigPath .\Find-UnderutilizedRoomBookings\config.psd1 `
    -Verbose
```

## Shared Module

Both scripts use the `ExchangeCore` module located in each script's `modules/` directory:

- `Import-ConfigurationFile` - Load JSON/PSD1 configs
- `Connect-ExchangeSession` / `Disconnect-ExchangeSession` - Exchange connections (supports LocalSnapin)
- `Connect-EwsService` - EWS service setup (supports SkipCertificateCheck)
- `Get-RoomCalendarItems` - Retrieve calendar meetings (auto-chunks large calendars)
- `Get-OrganizerState` - Check organizer status with external user matching
- `Get-ResolvedConnectionType` - Auto-detect OnPrem/EXO

## Running on Exchange Server (LocalSnapin Mode)

When Remote PowerShell is blocked or you get 401 errors, use `-LocalSnapin` to run directly on an Exchange server or a server with Exchange Management Tools:

```powershell
# From Exchange Management Shell or any PowerShell on Exchange server
.\Find-GhostRoomMeetings.ps1 -LocalSnapin -ConfigPath .\config.psd1 -Verbose

# With SSL cert bypass (self-signed certs)
.\Find-GhostRoomMeetings.ps1 -LocalSnapin -ConfigPath .\config.psd1 -SkipCertificateCheck -Verbose
```

LocalSnapin mode:
- Uses current Windows identity by default
- Works with Windows PowerShell 5.1 (snap-ins) and PowerShell 7+ (RemoteExchange.ps1)
- Requires Exchange Management Tools installed

## Additional Documentation

- **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)** - Common usage examples
- **[QUICK_REFERENCE.md](QUICK_REFERENCE.md)** - Parameter quick reference
