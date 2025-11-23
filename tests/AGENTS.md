# AGENTS.md

## Scope
Deze richtlijnen gelden voor alles binnen `tests/` en submappen.

## Pester-versie en runner
- Gebruik Pester 5.x; schrijf tests in Describe/Context/It-blokken en run met `Invoke-Pester -Path tests`.
- Houd testbestanden PowerShell 5.1+ en PowerShell 7+ compatibel.

## Bestandsnaam- en mapstructuur
- Plaats tests per domein in `tests/<domein>/` zodat ze aansluiten bij de scripts in de gelijknamige domeinmap.
- Gebruik de naamconventie `<Script>.Tests.ps1` (met dezelfde casing als het script) zodat Pester automatisch matcht.

## Testdata, mocks en outputs
- Test scripts met mock-credentials/configuraties via geisoleerde testdata: dummy `.psd1`/`.json` files en `$PSDefaultParameterValues` of `Mock` voor externe calls. Nooit echte secrets of tenant-IDs inchecken.
- Valideer CSV/Excel-output door tijdelijke bestanden te laten wegschrijven naar `$TestDrive` en de headers/rijen in te lezen met `Import-Csv` of `Import-Excel` (mock eventueel `Export-Excel`) in plaats van echte omgevingen te benaderen.
- Gebruik `Should -Throw`/`-Be` voor foutpaden en parametervalidatie, en `Assert-MockCalled` om side-effects te controleren zonder netwerk/system calls.

## Smoke-tests en PR-rapportage
- Voeg een rooktest toe voor ieder script: `pwsh -NoProfile -File <script>.ps1 -WhatIf` (of een expliciete `-TestMode` parameter). Rapporteer de uitkomst in het PR-template onder een sectie "Rooktest".

## Minimale dekking / checklist
- Minimaal: modules laden (bijv. `Import-Module ./modules/<Name>.psm1 -Force`), parameter-validatie (verplichte/optionele/switch), positieve paden én foutpaden.
- Breid uit met taak-specifieke asserts (bijv. filters/sorteringen) waar mogelijk. Wanneer volledige code coverage niet haalbaar is, documenteer de open lijnen in de testuitvoer of PR-notes.
