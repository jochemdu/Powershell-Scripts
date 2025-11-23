# Exchange-scripts

Repository met PowerShell-scripts voor Exchange, gestructureerd per domein.

## Structuur
- `exchange/` — Scripts voor Exchange Server en Exchange Online. Zie [`exchange/README.md`](exchange/README.md) voor vereisten en gebruik.
- `modules/` — Gedeelde PowerShell-modules (nog niet gevuld).
- `tests/` — Pester-tests of linting (nog niet gevuld).

Zie [`AGENTS.md`](AGENTS.md) voor de volledige code- en structuurafspraken.

## Beschikbare scripts
- [`exchange/Find-GhostRoomMeetings.ps1`](exchange/Find-GhostRoomMeetings.ps1): auditeert zaalpostvakken op ghost meetings en maakt rapportages.

### Kernparameters
- **ExchangeUri**: Remote PowerShell endpoint voor Exchange.
- **Credential**: Referenties met rechten op mail- en zaalpostvakken.
- **EwsAssemblyPath**: Pad naar de EWS Managed API-assembly.
- **MonthsAhead / MonthsBehind**: Datumbereik voor de controle.
- **OutputPath**: CSV-rapportpad (standaard `ghost-meetings-report.csv` in de huidige map).
- **ExcelOutputPath**: Optioneel Excel-rapport. Vereist het `ImportExcel`-module.
- **OrganizationSmtpSuffix**: Domeinsuffix om interne organisatoren te herkennen.
- **ImpersonationSmtp**: SMTP-adres voor EWS-impersonation en Autodiscover.
- **SendInquiry / NotificationFrom / NotificationTemplate**: Instellingen om deelnemers per e-mail te benaderen.

### Exporteren naar Excel
Wanneer **ExcelOutputPath** is opgegeven, probeert het script het `ImportExcel`-module te laden en schrijft het een `.xlsx`-bestand naast de CSV. Installeer het module vooraf met `Install-Module ImportExcel` indien nodig.
