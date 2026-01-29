# Changelog

Alle belangrijke wijzigingen in dit project worden in dit bestand bijgehouden.

## [2026-01-29] - Grote Opschoonactie & Herstructurering

### Toegevoegd
- Nieuwe mappenstructuur (`/src`, `/docs`, `/scripts`, `/archief`) volgens professionele standaarden.
- Centrale `README.md` in de root met duidelijke instructies voor bouw en installatie.
- `CHANGELOG.md` en `TODO.md` voor beter projectbeheer.
- Geavanceerde Daglichtberekening (alpha, beta, Ad, NEN 2057:2011) volledig geïntegreerd in de hoofdplugin.

### Gewijzigd
- `01_Documentatie` hernoemd naar `/docs`.
- `02_Bronbestanden` hernoemd naar `/src`.
- `03_Tools_Scripts` hernoemd naar `/scripts` en opgeschoond (alleen essentiële scripts behouden).
- `04_Archief` samengevoegd met de nieuwe `/archief` map.
- Alle `ElementId` compatibiliteitsproblemen voor Revit 2025 opgelost (overstap naar `long`).

### Verwijderd
- Losse `daglicht` bronbestanden (nu geïntegreerd in de main project).
- Verouderde Dynamo scripts en redundante implementaties van de daglichtberekening.
- Diverse redundante installatie- en deploymentscripts verplaatst naar het archief.

### Gearchiveerd
- Standalone `Daglicht` en `Dynamo` bronbestanden naar `/archief/legacy_bron`.
- Verouderde deployment scripts naar `/archief/legacy_scripts`.
