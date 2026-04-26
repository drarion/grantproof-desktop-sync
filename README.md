# GrantProof Desktop Sync — v10.8.0

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
