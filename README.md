# GrantProof Desktop Sync

Compagnon Windows local-first pour GrantProof.

## Ce que fait le compagnon

- reçoit les synchronisations envoyées par GrantProof Mobile sur le réseau local
- range les fichiers dans un dossier local `GrantProof`
- affiche un QR code de jumelage compact et un code manuel
- régénère des rapports premium Word et Excel donor-friendly directement dans le dossier projet
- permet de créer un raccourci bureau Windows avec l’icône GrantProof

## Flux produit retenu

1. configurer le poste et le dossier local
2. démarrer le compagnon
3. scanner le QR depuis GrantProof Mobile
4. synchroniser les preuves vers le dossier GrantProof
5. laisser le compagnon régénérer les rapports premium

## Démarrage local

```bash
pip install -r requirements.txt
python desktop_sync_app/app.py
```

## Build Windows avec GitHub Actions

Le workflow `.github/workflows/build-windows.yml` :
- construit un `.exe` avec `PyInstaller`
- applique l’icône GrantProof à l’exécutable
- embarque les assets de marque nécessaires à l’interface

## GitHub Pages

Le dossier `docs/` contient une page de téléchargement reliée aux GitHub Releases du dépôt.

## Sorties générées

Dans chaque projet synchronisé, le compagnon met à jour :
- `projects/<CODE>/reports/Project_Report.docx`
- `projects/<CODE>/reports/Project_Register.xlsx`
- `projects/<CODE>/reports/Project_Report_FR.docx`
- `projects/<CODE>/reports/Project_Report_EN.docx`
- `projects/<CODE>/reports/Project_Register_FR.xlsx`
- `projects/<CODE>/reports/Project_Register_EN.xlsx`
- `evidence.docx` ou `story.docx` dans chaque dossier d’élément
