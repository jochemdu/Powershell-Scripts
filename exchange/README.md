# Exchange Room Mailbox Auditing Scripts

Scripts voor Exchange Server en Exchange Online beheer.

## Available Scripts

### Find-GhostRoomMeetings.ps1 (v1 - Universal)
**Compatibility**: PowerShell 1.0 - 7.x
**Best For**: Legacy environments, maximum compatibility

### Find-GhostRoomMeetings-v7.ps1 (v7 - Modern)
**Compatibility**: PowerShell 7.0+
**Best For**: Modern environments, large deployments (5-8x faster)

## Quick Start

### v1 (Universal - All PowerShell Versions)
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath config.example.psd1 `
    -Credential $cred
```

### v7 (Modern - PowerShell 7+ Only)
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath config.example.json `
    -Credential $cred
```

## Documentation

- **[PS7_FEATURES.md](PS7_FEATURES.md)** - Detailed PS7 features and optimizations
- **[VERSION_COMPARISON.md](VERSION_COMPARISON.md)** - Comparison between v1 and v7
- **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)** - v1 usage examples
- **[USAGE_EXAMPLES_V7.md](USAGE_EXAMPLES_V7.md)** - v7 usage examples
- **[REFACTORING_SUMMARY.md](../REFACTORING_SUMMARY.md)** - v1 refactoring details

## Find-GhostRoomMeetings.ps1 (v1)
Auditeert vergaderingen in zaalpostvakken om zogeheten "ghost meetings" te detecteren waarbij de organisator ontbreekt of gedeactiveerd is.

### Vereisten
- PowerShell 1.0 of later
- On-prem: Exchange Management Shell of remote PowerShell sessie
- Exchange Online: `ExchangeOnlineManagement`-module
- EWS Managed API assembly
- Serviceaccount met EWS-impersonation rechten
- Optioneel: `ImportExcel`-module voor Excel export

### Voorbeeldgebruik
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings.ps1 `
    -ConfigPath config.example.psd1 `
    -Credential $cred `
    -MonthsAhead 6
```

## Find-GhostRoomMeetings-v7.ps1 (v7)
PowerShell 7+ optimized version met parallel processing (5-8x sneller).

### Vereisten
- PowerShell 7.0 of later
- EWS Managed API assembly
- Serviceaccount met EWS-impersonation rechten
- Optioneel: `ImportExcel`-module voor Excel export

### Voorbeeldgebruik
```powershell
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath config.example.json `
    -Credential $cred `
    -ThrottleLimit 8
```

### Performance
- v1: 100 rooms in ~450 seconds
- v7: 100 rooms in ~65 seconds (6.9x faster)


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
