# Exchange

Scripts voor Exchange Server en Exchange Online beheer.

## Find-GhostRoomMeetings.ps1
Auditeert vergaderingen in zaalpostvakken om zogeheten "ghost meetings" te detecteren waarbij de organisator ontbreekt of gedeactiveerd is.

### Vereisten
- Toegang tot de Exchange Management Shell of een remote PowerShell sessie naar Exchange (`-ExchangeUri`).
- Serviceaccount met EWS-impersonation en voldoende rechten op zaalpostvakken.
- Lokale beschikbaarheid van de EWS Managed API-assembly (`-EwsAssemblyPath`).
- Optioneel: het `ImportExcel`-module voor het genereren van een `.xlsx`-rapport.

### Voorbeeldgebruik
```powershell
pwsh -NoProfile -File ./exchange/Find-GhostRoomMeetings.ps1 \ 
    -ExchangeUri 'http://exchange.contoso.com/PowerShell/' \ 
    -Credential (Get-Credential) \ 
    -ImpersonationSmtp 'service@contoso.com' \ 
    -MonthsAhead 6 \ 
    -OutputPath 'ghost-meetings.csv'
```
