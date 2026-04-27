# GrantProof Desktop Sync — v10.6.0

Compagnon PC local-first pour recevoir les preuves terrain, reconstruire les registres et générer les rapports premium DOCX.

## Phase R3

- Rapports utilisateur uniquement dans `projects/<code>/reports`.
- Suppression des anciens `evidence.docx` / `story.docx` dans les dossiers techniques lors de la régénération.
- Rapport premium DOCX haute fidélité : page 1 dashboard, page 2 annexe visuelle si plusieurs images.
- Carte pays automatique avec point orange uniquement si GPS disponible ou capitale reconnue.
- Détection sectorielle automatique : agriculture, sécurité alimentaire, abris, WASH/EHA, santé, nutrition, éducation, protection, cash, coordination.

## Build Windows

Utiliser GitHub Actions : `.github/workflows/build-windows.yml`.


## v10.1.0 — Editable premium reports

- Premium activity reports are now generated as true editable DOCX documents: titles, paragraphs, KPI labels, highlights and captions can be edited directly in Word.
- Page 1 keeps the dashboard-style layout while avoiding full-page image rendering.
- Page 2 is a structured visual annex for additional media.
- Country maps are generated from the project country. An orange point is plotted only when GPS coordinates are available.
- Legacy per-evidence Word files remain cleaned from technical evidence folders during rebuild.


## v10.6.0 — perfect premium model mimicry
- page 1 now uses a high-resolution premium dashboard renderer closely aligned with the provided premium report model.
- matching section-title icon language, thin title rules, rounded hero photo, model-like cards and spacing.
- header/meta/map blocks repositioned to match the model proportions.
- annex pages retained after the dashboard page when additional images are available.


## v10.6.0 — Editable model mimicry restoration
- Restores page 1 as a true editable DOCX layout instead of a full-page screenshot.
- Uses model-derived premium icons for title blocks, KPI cards, overview metadata, and sector alignment.
- Raises paragraph/KPI typography to match the model hierarchy, with body text >= 8 pt.
- Enlarges KPI numbers and uses white/pale icon backgrounds as in the model.


## v10.8.0 — Editable premium model clone
- Returned to a fully editable DOCX layout while matching the premium report model more closely.
- Re-extracted model icons from the supplied reference and normalized them for Word rendering.
- Increased typography hierarchy: larger header title, section headings, body text and KPI numbers.
- KPI cards now follow the reference more closely with large bold values and pale icon backgrounds.
- Row heights are now minimum heights rather than exact heights to avoid clipped text.


## R24 — report polish: sector icons, map and bottom band
- Raised body text minimum rendering to 8 pt.
- Enlarged and unclipped sector icons.
- Aligned Highlights and Location titles.
- Improved location map rendering and bottom band vertical space.


## v10.6.2 — report polish: font, sector icons, map and media sizing
- minimum readable report text raised to 10 pt
- sector icons regenerated from higher quality premium assets and enlarged
- proof image enlarged on page 1
- location map enlarged and visually re-centered
- bottom band and overview metadata spacing adjusted

## R26 — premium icon library + header spacing fix
- Regenerated the editable report icon library with a coherent premium line-icon set for header metadata, sections, KPIs, overview info rows, sectors and humanitarian categories.
- Improved header spacing: larger internal margins, wider right metadata block, dedicated white organization/location/calendar icons, and safer text truncation.
- Added real spacer columns between dashboard cards and bottom bands so the DOCX remains editable while looking less like a compact Word table.
- Rebalanced the hero image, KPI cards and location map so the first page fits cleanly without clipping or moving the map to a second page.
- Kept the report as a true editable DOCX: text, tables, images, icons and maps remain independent Word elements.

## R33 — Excel premium report template
- Premium activity reports are now generated as editable `.xlsx` files from `desktop_sync_app/assets/templates/grantproof_report_template_excel.xlsx`.
- Project report outputs are `Project_Report_FR.xlsx` and `Project_Report_EN.xlsx` in each project `reports` folder.
- The existing project registers remain unchanged as `Project_Register_FR.xlsx` and `Project_Register_EN.xlsx`.
- The Excel generator fills template placeholders, inserts the proof image, map, KPI icons and sector icons, and keeps the template formatting/premium layout.

## R34 — premium PPTX report engine
- Premium activity reports are now generated as editable PowerPoint files: `Project_Report_FR.pptx` and `Project_Report_EN.pptx`.
- The report page is built from native PowerPoint objects: editable text boxes, shapes, independent images, map, icons and KPI blocks.
- The layout follows the ONG/UN-style dashboard model: full-width blue header, KPI cards, hero evidence photo, sector block, highlights and location map.
- Existing Excel registers are preserved for structured data.
