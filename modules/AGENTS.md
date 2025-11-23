# AGENTS.md

## Scope
Deze richtlijnen gelden voor alle modules onder `modules/` en alle submappen daarin. Gebruik extra `AGENTS.md` bestanden voor module-specifieke uitzonderingen.

## Mappenstructuur
- Elke module leeft in `modules/<Naam>/` met exact twee hoofdartefacten:
  - `modules/<Naam>/<Naam>.psm1`
  - `modules/<Naam>/<Naam>.psd1`
- De naam van de map, het `.psm1`-bestand en het `.psd1`-manifest moeten gelijk zijn.
- Houd gedeelde of ondersteunende scripts binnen dezelfde modulemap; vermijd verspreide bestanden buiten de modulemap.

## Module-export
- Gebruik in elk `.psm1` een expliciete `Export-ModuleMember -Function <lijst>` voor alle publieke functies.
- Voeg alleen functies toe die publiek horen te zijn; alles wat intern is blijft ongeëxporteerd.

## Versiebeheer
- Beheer moduleversies volgens semver (`MAJOR.MINOR.PATCH`) in het `.psd1` manifest (`ModuleVersion`).
- Verhoog de versie bij elke wijziging die wordt vrijgegeven; documenteer breaking changes met een MAJOR bump.

## Comment-based help
- Elke publieke functie in het `.psm1` krijgt volledige comment-based help met minimaal `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE` en `.OUTPUTS`.
- Functies zonder complete help-blokken worden niet geaccepteerd.

## Tests
- Plaats Pester-tests in `tests/` met een duidelijke naam die het doel van de module weerspiegelt.
- Laad modules in tests via `Import-Module` (niet via relatieve dot-sourcing) zodat exports en manifest-vereisten getest worden.
- Richt je tests op functiegedrag, input-validatie en foutafhandeling.

## Linting
- Gebruik PowerShell ScriptAnalyzer (of vergelijkbare linting) met een duidelijke baseline; corrigeer of motiveer afwijkingen.
- Voer linting uit voordat je commit; deel eventuele uitzonderingen in de PR.

## Dependencies
- Geen hardcoded paden in modules of tests; gebruik parameterisatie of configuratiebestanden.
- Definieer externe module-afhankelijkheden in het `.psd1` manifest (`RequiredModules` of `Prerelease` info indien relevant) in plaats van handmatige checks in code.

## PR/documentatie
- Bij nieuwe of gewijzigde modules hoort een bijgewerkte `README.md` in de relevante domeinmap en/of een toelichting in `docs/`.
- Licht testresultaten en linting toe in de PR-beschrijving.
