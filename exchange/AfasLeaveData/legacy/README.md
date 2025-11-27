# Legacy Scripts

Deze map bevat oude/legacy scripts die als referentie dienen voor de nieuwe `AfasLeaveData.ps1`.

## Scripts

| Script | Functie | Status |
|--------|---------|--------|
| `Import-CalendarCSV.ps1` | Importeert verlofdata van Integration Bus API naar kalenders | Te refactoren |
| `Remove-CalendarItemsCSV.ps1` | Verwijdert geannuleerde verlofitems uit kalenders | Te refactoren |

## Gemeenschappelijke Kenmerken

Beide scripts gebruiken:
- **Integration Bus API** - REST endpoints voor leave data (niet directe AFAS connectie)
- **Password file** - `C:\ScheduledTasks\AfasLeaveData\password.txt`
- **EWS Managed API** - Voor kalendertoegang met impersonation
- **Get-Mailbox** - Voor ITCode → email mapping
- **CSV tussenbestanden** - Data wordt eerst als CSV opgeslagen
- **Tab-separated logging** - `Context\tStatus\tMessage` formaat

## Migratie naar AfasLeaveData.ps1

De nieuwe `AfasLeaveData.ps1` combineert beide scripts met:
- ✅ Configureerbare paden via config file
- ✅ Zelfde password file aanpak (geen SecretManagement)
- ✅ Zelfde logging formaat (compatibel)
- ✅ Zelfde Get-Mailbox mapping strategie
- ✅ Support voor zowel import als remove operaties
- ✅ Test mode voor validatie zonder wijzigingen

## Password File Aanmaken

```powershell
# Eenmalig uitvoeren op de server waar het script draait:
Read-Host -Prompt "Enter password" -AsSecureString | 
    ConvertFrom-SecureString | 
    Out-File "C:\ScheduledTasks\AfasLeaveData\password.txt"
```

> ⚠️ De password file is gebonden aan de Windows gebruiker en machine waarop deze is aangemaakt.

---

*Legacy scripts kunnen verwijderd worden na succesvolle migratie en testing.*
