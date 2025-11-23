# AGENTS.md

## Scope
Deze richtlijnen gelden voor de volledige repository. Voeg een extra `AGENTS.md` toe in submappen wanneer daar afwijkende of aanvullende afspraken gelden; die hebben voorrang binnen hun map.

## Structuur
- Groepeer scripts per domein in submappen (bijv. `exchange/`, `azure-ad/`, `m365/`, `onprem/`).
- Deelbare functionaliteit hoort in PowerShell-modules onder `modules/<ModuleNaam>/<ModuleNaam>.psm1` met een expliciete exportlijst.
- Voeg per domeinmap een korte `README.md` toe met doel, vereisten en voorbeeldgebruik. Bewaar uitgebreidere handleidingen in `docs/` en eventuele voorbeelden in `examples/`.
- Bewaar tests of linting in `tests/` (bijv. Pester).

## PowerShell-stijl
- Gebruik Verb-Noun namen (Get-/Set-/Test-) en PascalCase voor parameters.
- Begin scripts met `[CmdletBinding()]`, `Set-StrictMode -Version Latest` en `$ErrorActionPreference = 'Stop'`.
- Valideer invoer met `[Validate*]` attributen; gebruik `switch` voor booleans en schrijf duidelijke help-blokken (`.SYNOPSIS`, `.EXAMPLE`).
- Vermijd hardcoded waarden; lees configuratie uit parameters of (optioneel) `.psd1/.json` bestanden.

## Documentatie
- Werk de relevante `README.md` bij wanneer je een script of module toevoegt of wijzigt.
- Beschrijf minimaal doel, vereisten (modules, scopes, rechten), parameters en voorbeeldcommando’s.

## Tests en kwaliteitschecks
- Voeg waar mogelijk Pester-tests toe en run `Invoke-Pester` voor je commit of PR.
- Voer een rooktest uit met `pwsh -NoProfile -File <script>.ps1 -WhatIf` of een testmodus.

## Commits en PR’s
- Eén functionele wijziging per commit; gebruik beschrijvende commit messages.
- Voeg relevante logs of (samengevatte) output toe in de PR-beschrijving wanneer beschikbaar.
