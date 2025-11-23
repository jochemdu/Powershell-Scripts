# Exchange

Scripts voor Exchange Server en Exchange Online beheer.

## Find-GhostRoomMeetings.ps1
Auditeert vergaderingen in zaalpostvakken om zogeheten "ghost meetings" te detecteren waarbij de organisator ontbreekt of gedeactiveerd is.

### Vereisten
- PowerShell 5.1 of 7+.
- On-prem: toegang tot de Exchange Management Shell of een remote PowerShell sessie (`-ExchangeUri`), plus AD-module voor uitschakelstatus.
- Exchange Online: `ExchangeOnlineManagement`-module en moderne authenticatie via `Connect-ExchangeOnline`.
  - Delegated/OAuth scopes: gebruikersreferentie met passende rollen (bijv. `Organization Management`) of App-Only met `Exchange.ManageAsApp`.
- Serviceaccount met EWS-impersonation en voldoende rechten op zaalpostvakken (EWS moet moderne authenticatie toestaan in EXO).
- Lokale beschikbaarheid van de EWS Managed API-assembly (`-EwsAssemblyPath`).
- Optioneel: het `ImportExcel`-module voor het genereren van een `.xlsx`-rapport.

### Voorbeeldgebruik
```powershell
pwsh -NoProfile -File ./exchange/Find-GhostRoomMeetings.ps1 \
    -ConnectionType Auto \
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' \
    -Credential (Get-Credential) \
    -ImpersonationSmtp 'service@contoso.com' \
    -MonthsAhead 6 \
    -OutputPath 'ghost-meetings.csv'
```

### Exchange Online voorbeeld
```powershell
Import-Module ExchangeOnlineManagement

pwsh -NoProfile -File ./exchange/Find-GhostRoomMeetings.ps1 \
    -ConnectionType EXO \
    -Credential (Get-Credential -UserName 'service@contoso.com') \
    -ImpersonationSmtp 'service@contoso.com' \
    -MonthsAhead 3 \
    -OutputPath './reports/ghost-meetings.csv' \
    -TestMode:$false
```

### Parameters
- **ConnectionType**: `OnPrem`, `EXO` of `Auto` (detectie op `ExchangeUri`). Stuurt de juiste cmdlets (`Get-Mailbox` vs. `Get-ExoMailbox`/`Get-ExoRecipient`).
- **TestMode**: Zet mockbare testmodus aan; slaat daadwerkelijke connecties over en vult dummy-credentials in.
- Overige kernparameters: zie [root README](../README.md) voor uitleg over EWS, rapportpaden en notificaties.

### Tests en rooktest
- Pester-rooktest beschikbaar onder `tests/exchange/Find-GhostRoomMeetings.Tests.ps1` (laadt het EXO-pad met mocks/testmodus).
- Draai alle tests met `Invoke-Pester -Path tests` vanuit de repo-root.

## Find-UnderutilizedRoomBookings.ps1
Spoort vergaderingen op waar grote vergaderruimtes (bijv. 6+ plaatsen) geboekt zijn voor slechts één of enkele deelnemers.

### Vereisten
- PowerShell 5.1 of 7+.
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
- **TestMode**: Skip daadwerkelijke connecties en gebruik dummy-credentials voor Pester-tests.

### Tests en rooktest
- Pester-test beschikbaar onder `tests/exchange/Find-UnderutilizedRoomBookings.Tests.ps1` (maakt gebruik van mocks/TestDrive-outputs).
- Draai alle tests met `Invoke-Pester -Path tests` vanuit de repo-root.
