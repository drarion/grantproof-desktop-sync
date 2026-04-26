# GrantProof Desktop Sync — v10.3.1

Compagnon PC local-first pour recevoir les preuves terrain, reconstruire les registres et générer les rapports premium DOCX.

## Phase R3

- Rapports utilisateur uniquement dans `projects/<code>/reports`.
- Suppression des anciens `evidence.docx` / `story.docx` dans les dossiers techniques lors de la régénération.
- Rapport premium DOCX haute fidélité : page 1 dashboard, page 2 annexe visuelle si plusieurs images.
- Carte pays automatique avec point orange uniquement si GPS disponible ou capitale reconnue.
- Détection sectorielle automatique : agriculture, sécurité alimentaire, abris, WASH/EHA, santé, nutrition, éducation, protection, cash, coordination.

## Build Windows

Utiliser GitHub Actions : `.github/workflows/build-windows.yml`.


## v10.2.0 — Premium report polish and layout refinement

- Premium activity reports are now generated as true editable DOCX documents: titles, paragraphs, KPI labels, highlights and captions can be edited directly in Word.
- Page 1 keeps the dashboard-style layout while avoiding full-page image rendering.
- Page 2 is a structured visual annex for additional media.
- Country maps are generated from the project country. An orange point is plotted only when GPS coordinates are available.
- Legacy per-evidence Word files remain cleaned from technical evidence folders during rebuild.


## v10.3.0 — Ultra-premium report refinements
- section title icons switched back to single-color styling with more premium icon choices
- header remains full-bleed while internal text aligns with report body margins
- activity overview no longer truncates the main narrative summary
- logframe section now has its own icon and improved spacing
- sector alignment, location map, and page-1 bottom area refined for a cleaner NGO-grade report layout


## v10.3.1 — Layout overflow safety pass
- increased page-1 top band vertical space to prevent truncation in overview/logframe blocks
- tightened secondary spacing and media block height for safer one-page composition
