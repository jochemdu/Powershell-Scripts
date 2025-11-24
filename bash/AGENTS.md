# AGENTS.md

## Scope
Deze richtlijnen gelden voor alle Bash-scripts en aanverwante bestanden in deze `bash/` map en submappen.

## Bash-stijl
- Gebruik `#!/usr/bin/env bash` als shebang en start scripts met `set -euo pipefail` en `IFS=$'\n\t'`.
- Kies duidelijke bestandsnamen met een korte, beschrijvende naam en de extensie `.sh`.
- Vermijd hardcoded paden en credentials; lees variabelen uit omgevingsvariabelen of configuratiebestanden.
- Schrijf functies in `snake_case` en houd scripts idempotent waar mogelijk.

## Documentatie
- Plaats bovenin elk script een korte comment-blok met doel, vereisten en voorbeeldgebruik.
- Werk bijpassende `README.md` in deze map bij wanneer je nieuwe scripts toevoegt of bestaande wijzigt.

## Kwaliteit
- Draai `shellcheck` op nieuwe of gewijzigde scripts en los waarschuwingen op.
- Voeg, waar zinvol, eenvoudige zelftests of dry-run modi toe (bijv. via `-n`/`--dry-run`).
