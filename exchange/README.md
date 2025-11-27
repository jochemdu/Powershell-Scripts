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
| [Find-GhostRoomMeetings](Find-GhostRoomMeetings/) | Detect meetings with missing/disabled organizers | 5.1+ (7+ recommended) | [README](Find-GhostRoomMeetings/README.md) |
| [Find-UnderutilizedRoomBookings](Find-UnderutilizedRoomBookings.ps1) | Find large rooms booked for few attendees | 5.1+ | See below |

## Quick Start

### Find-GhostRoomMeetings

```powershell
# Using config file
.\Find-GhostRoomMeetings\Find-GhostRoomMeetings.ps1 `
    -ConfigPath .\Find-GhostRoomMeetings\config.json `
    -Credential (Get-Credential)

# Direct parameters
.\Find-GhostRoomMeetings\Find-GhostRoomMeetings.ps1 `
    -ConnectionType OnPrem `
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' `
    -ImpersonationSmtp 'service@contoso.com' `
    -Credential (Get-Credential) `
    -MonthsAhead 6
```

## Additional Documentation

- **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)** - Common usage examples
- **[QUICK_REFERENCE.md](QUICK_REFERENCE.md)** - Parameter quick reference

## Find-UnderutilizedRoomBookings.ps1
Spoort vergaderingen op waar grote vergaderruimtes (bijv. 6+ plaatsen) geboekt zijn voor slechts één of enkele deelnemers.

### Vereisten
- PowerShell 1+.
- On-prem: toegang tot de Exchange Management Shell of een remote PowerShell sessie (`-ExchangeUri`).
- Exchange Online: `ExchangeOnlineManagement`-module en moderne authenticatie via `Connect-ExchangeOnline`.
- EWS Managed API assembly beschikbaar op het opgegeven pad (`-EwsAssemblyPath`).
- Impersonationrechten voor de opgegeven serviceaccount (bijv. `ApplicationImpersonation` in EXO).

### Voorbeeldgebruik
```powershell
pwsh -NoProfile -File ./exchange/Find-UnderutilizedRoomBookings.ps1 \
    -ConnectionType Auto \
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' \
    -ImpersonationSmtp 'service@contoso.com' \
    -MinimumCapacity 6 \
    -MaxParticipants 2 \
    -OutputPath './reports/underutilized.csv'
```

### Parameters
- **MinimumCapacity**: Alleen ruimtes scannen met deze minimumcapaciteit of hoger (standaard 6).
- **MaxParticipants**: Signaleer vergaderingen met maximaal dit aantal deelnemers (standaard 2, telt organisator + aanwezigen).
- **MonthsAhead/MonthsBehind**: Datumvenster voor de kalenderquery.
