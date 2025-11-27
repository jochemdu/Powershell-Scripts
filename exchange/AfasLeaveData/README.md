# AfasLeaveData

Synchroniseert verlofdata van AFAS naar Exchange/Outlook kalenders via Integration Bus REST API.

## Features

- Ontvangt AFAS verlofdata via Integration Bus REST API (niet directe AFAS connectie)
- Creëert/update kalenderitems in Exchange (On-Premises)
- Verwijdert geannuleerde verlofitems
- Ondersteunt impersonation voor toegang tot gebruikerskalenders
- Get-Mailbox mapping van ITCode naar email (zoals legacy scripts)
- Password file authenticatie (geen SecretManagement)
- CSV en Excel rapportage van sync resultaten
- Compatibele logging met legacy scripts
- Test mode voor droog draaien

## Requirements

### PowerShell

| Versie | Ondersteuning |
|--------|---------------|
| 5.1    | ✅ Volledig   |
| 7.x    | ✅ Volledig   |

### Modules & Assemblies

| Component | Doel | Installatie |
|-----------|------|-------------|
| EWS Managed API | Exchange kalender toegang | [Download](https://www.microsoft.com/en-us/download/details.aspx?id=42951) |
| Exchange Cmdlets | Get-Mailbox voor user mapping | Via Exchange Management Shell |
| ImportExcel | Excel export (optioneel) | `Install-Module ImportExcel` |

### Vereiste Rechten

| Recht | Scope | Doel |
|-------|-------|------|
| ApplicationImpersonation | Exchange | Toegang tot gebruikerskalenders |
| View-Only Recipients | Exchange | Get-Mailbox toegang |

## Directory Structure

```
AfasLeaveData/
├── AfasLeaveData.ps1           # Hoofdscript
├── config.example.json         # JSON configuratie template
├── config.example.psd1         # PowerShell Data configuratie template
├── README.md                   # Deze documentatie
├── legacy/                     # Originele scripts (referentie)
│   ├── Import-CalendarCSV.ps1
│   ├── Remove-CalendarItemsCSV.ps1
│   └── README.md
└── modules/
    └── AfasCore/
        └── AfasCore.psm1       # Gedeelde functies
```

## Configuration

Kopieer `config.example.json` naar `config.json` en pas de waarden aan:

```json
{
  "Connection": {
    "Type": "OnPrem",
    "ExchangeUri": "http://exchange.contoso.com/PowerShell/"
  },
  "Credential": {
    "Username": "svc_account",
    "PasswordFile": "C:\\ScheduledTasks\\AfasLeaveData\\password.txt"
  },
  "Api": {
    "LeaveDataEndpoint": "https://integrationbus.contoso.com:7843/api/v1/getLeaveData",
    "CanceledLeaveEndpoint": "https://integrationbus.contoso.com:7843/api/v1/getCanceledLeaveData",
    "ProxyUrl": "http://proxy.contoso.com:8080"
  },
  "UserMapping": {
    "Strategy": "Mailbox"
  }
}
```

### Password File Aanmaken

```powershell
# Eenmalig uitvoeren op de server:
Read-Host -Prompt "Enter password" -AsSecureString | 
    ConvertFrom-SecureString | 
    Out-File "C:\ScheduledTasks\AfasLeaveData\password.txt"
```

> ⚠️ De password file is gebonden aan de Windows gebruiker en machine.

### Configuration Options

| Parameter | Type | Default | Beschrijving |
|-----------|------|---------|--------------|
| `Connection.Type` | string | `OnPrem` | Exchange type (alleen OnPrem ondersteund) |
| `Connection.ExchangeUri` | string | - | Exchange PowerShell endpoint |
| `Credential.Username` | string | - | Service account username |
| `Credential.PasswordFile` | string | - | Pad naar password file |
| `Api.LeaveDataEndpoint` | string | - | Integration Bus leave data URL |
| `Api.CanceledLeaveEndpoint` | string | - | Integration Bus canceled leave URL |
| `Api.ProxyUrl` | string | null | Proxy URL indien nodig |
| `UserMapping.Strategy` | string | `Mailbox` | `Mailbox`, `Email`, of `MappingTable` |
| `Paths.ScriptPath` | string | - | Werkdirectory voor CSV bestanden |
| `Paths.ProcessedPath` | string | - | Directory voor verwerkte bestanden |
| `Paths.LogPath` | string | - | Directory voor log bestanden |

## Usage Examples

### Basis gebruik

```powershell
# Met config file (credentials uit password file)
.\AfasLeaveData.ps1 -ConfigPath .\config.json

# Test mode (geen wijzigingen)
.\AfasLeaveData.ps1 -ConfigPath .\config.json -TestMode -Verbose
```

### Met expliciete credentials

```powershell
.\AfasLeaveData.ps1 `
    -ConfigPath .\config.json `
    -Credential (Get-Credential)
```

### Scheduled Task

```powershell
# Typische scheduled task setup
powershell.exe -ExecutionPolicy Bypass -File "C:\ScheduledTasks\AfasLeaveData\AfasLeaveData.ps1" -ConfigPath "C:\ScheduledTasks\AfasLeaveData\config.json"
```

## Output

### CSV Report Columns

| Kolom | Beschrijving |
|-------|--------------|
| Employee | Medewerker ID |
| Email | Exchange email adres |
| LeaveType | Type verlof |
| StartDate | Begindatum |
| EndDate | Einddatum |
| Status | Sync status (Created/Updated/Error) |
| CalendarItemId | Exchange item ID |

### Console Output

```
AfasLeaveData v1.0.0
Retrieved 15 leave entries from Integration Bus
Processing: 15/15 [====================] 100%
Created: 12, Updated: 2, Errors: 1
Report saved to: ./reports/afas-leave-sync-report.csv
```

## Log Format

Compatibel met legacy scripts (tab-separated):

```
Context             	Status    	Message
AfasLeaveData       	[START]   	*** Start Logging: 27-11-2024 10:00 ***
Load Exchange CmdLets	[SUCCESS] 	Exchange CmdLets loaded successfully
Create calendar item	[SUCCESS] 	Created MyPlace Leave Booking 2024-12-01 for user@contoso.com
AfasLeaveData       	[STOP]    	*** End Logging: 27-11-2024 10:05 ***
```

## Troubleshooting

### Common Issues

**API endpoint niet bereikbaar**
- Controleer de endpoint URL
- Verify proxy instellingen
- Check firewall/network rules
- Test met: `Invoke-RestMethod -Uri $url -Credential $cred`

**Password file werkt niet**
- Password file is gebonden aan gebruiker/machine
- Maak opnieuw aan op de juiste server
- Check bestandsrechten

**Get-Mailbox faalt voor ITCode**
- Verify Exchange session is geladen
- Check of ITCode bestaat als mailbox
- Probeer `Get-Mailbox -Identity <ITCode>` handmatig

**EWS authentication fails**
- Verify credentials
- Check ApplicationImpersonation rechten
- Test EWS URL met browser

### Debug Mode

```powershell
.\AfasLeaveData.ps1 -ConfigPath .\config.json -Verbose -Debug
```

## Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌──────────────────┐
│    AFAS     │────►│  Integration     │────►│  AfasLeaveData   │
│  (HR Data)  │     │      Bus         │     │    (Script)      │
└─────────────┘     └──────────────────┘     └────────┬─────────┘
                           REST API                   │
                                                      ▼
                                             ┌──────────────────┐
                                             │    Exchange      │
                                             │   (Calendars)    │
                                             └──────────────────┘
```

## Related Scripts

- [Find-GhostRoomMeetings](../Find-GhostRoomMeetings/) - Detecteert ghost meetings
- [Find-UnderutilizedRoomBookings](../Find-UnderutilizedRoomBookings/) - Vindt ondergebruikte ruimteboekingen

## Version History

| Versie | Datum | Wijzigingen |
|--------|-------|-------------|
| 1.0.0  | 2024-xx-xx | Initiële versie |

## License

Zie repository root voor licentie informatie.
