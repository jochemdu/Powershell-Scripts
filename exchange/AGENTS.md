# AGENTS.md

## Scope
Deze richtlijnen gelden voor `exchange/` en alle onderliggende mappen en bestanden.

## PowerShell-richtlijnen
- Minimale vereiste PowerShell-versie: 5.1 of PowerShell 7+.
- Vereiste modules/assemblies: Exchange Management Shell (of Exchange Online PowerShell voor EXO), EWS Managed API, en ImportExcel.
- Elk script start met `[CmdletBinding()]`, `Set-StrictMode -Version Latest` en `$ErrorActionPreference = 'Stop'`.
- Gebruik `[Validate*]`-attributen voor parametercontrole en definieer booleans als `switch`.
- Vermijd het opslaan van gevoelige waarden in code of JSON; lees credentials of app secrets via veilige opslag (bijv. SecretManagement) of interactief.

## Configuratie (`-ConfigPath`)
- Config JSON bevat minimaal tenant- of organisatiegegevens en per mailbox of taak de parameters voor het script.
- Voorbeeld (on-prem/EXO zonder impersonation):
  ```json
  {
    "Connection": {
      "Type": "OnPrem", // of "EXO"
      "EwsUrl": "https://mail.contoso.com/EWS/Exchange.asmx", // optioneel bij autodiscover
      "Autodiscover": true
    },
    "Mailboxes": [
      {
        "SmtpAddress": "user@contoso.com",
        "ReportPathCsv": "reports/user.csv",
        "ReportPathXlsx": "reports/user.xlsx"
      }
    ]
  }
  ```
- Voeg voor impersonation een `Impersonation`-object toe met `SmtpAddress` voor `-ImpersonationSmtp`. Beperk dit tot niet-gevoelige info; wachtwoorden, tokens en certificaatpaden blijven buiten het JSON-bestand.

## Verbindingen en autodiscover
- Ondersteun zowel autodiscover als expliciete `-EwsUrl`; documenteer wanneer handmatige URL vereist is (bijv. on-prem zonder autodiscover).
- Geef aan hoe verbindingen verschillen voor on-prem (gebruik bestaande EMS of Exchange Management Shell) versus Exchange Online (moderne authenticatie, `Connect-ExchangeOnline`).
- Maak duidelijk hoe de EWS Managed API wordt geladen (bijv. vanuit assemblies of vooraf geïnstalleerde module).
- 
## Opslag
- Elk powershellprogramma heeft zijn eigen subdirectory. Een powershell programma kan uit meerdere script bestaand
- Alle benodigde modules sla je ook op in dezelfde subdirectory in directory modules.
- Configureerbare parameters sla je ook op in de subdirectory in een asart bestand.
- Elke powershell programma heeft zijn eigen README
- 
## Logging, rapportage en tests
- Log output en rapporten per mailbox naar opgegeven CSV- en/of Excel-paden; gebruik ImportExcel voor `.xlsx`-exports.
- Rooktest: voer scripts uit met `-WhatIf` (of een expliciete testmodus) met een veilige dummy-configuratie.
- Pester: voorzie een dummy-configbestand voor tests; test expectations moeten mockbare connecties en rapportpaden dekken.

## Impersonation
- Wanneer `-ImpersonationSmtp` wordt gebruikt, beschrijf vereiste rechten (bijv. `ApplicationImpersonation`) en valideer het e-mailadres met `[ValidatePattern]`.

## Documentatie
- Documenteer in `README.md` de vereiste modules, PowerShell-versie, config- en rapportpaden en de benodigde rechten voor impersonation of autodiscover.
- In configuratiebestanden beschrijf je ook waar de parameter voor dient. Bij JSON doe je dit door middel van // commentaar
- Documentatie sla je op per programma.
  - QUICK_REFERENCE.md sla je op per programma
  - USAGE_EXAMPLES.md sla je op per programma

## Meeting-room scripts (zoekbereik)
- Stel het standaardzoekbereik altijd in vanaf de huidige datum (bijv. `MonthsBehind = 0`) zodat rapportages niet onbedoeld ver het verleden ingaan.
- Maak `MonthsBehind`/`MonthsAhead` of vergelijkbare tijdsvensters configureerbaar via parameters en via de configbestanden (JSON/PSD1), met duidelijke standaardwaarden.
- Gebruik `[ValidateRange()]` of `[ValidatePattern()]` voor datum- en periodeparameters en definieer defaults in de param block, zodat scripts weigeren te starten wanneer onrealistische of negatieve waarden worden opgegeven.
