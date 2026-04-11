# GrantProof Desktop Sync

Compagnon Windows gratuit pour GrantProof.

## À quoi ça sert

- reçoit les synchronisations envoyées par GrantProof Mobile sur le réseau local
- range les fichiers dans un dossier local `GrantProof`
- affiche un QR code de jumelage pour lier un téléphone à un ordinateur précis
- régénère des rapports premium Word et Excel donor-friendly directement dans le dossier projet

## Modèle V1

- pas de serveur GrantProof
- pas de cloud GrantProof
- sync locale uniquement
- PC allumé + logiciel lancé + téléphone sur le même Wi‑Fi

## Démarrage local

1. Installe Python 3.11+
2. Ouvre un terminal dans ce dossier
3. Lance :

```bash
pip install -r requirements.txt
python desktop_sync_app/app.py
```

## Build Windows gratuit avec GitHub Actions

Le workflow `.github/workflows/build-windows.yml` :
- construit un `.exe` avec `PyInstaller`
- publie l'asset dans GitHub Releases quand tu pousses un tag `v*`

## GitHub Pages

Le dossier `docs/` contient une page de téléchargement sobre qui peut être activée avec GitHub Pages.
Quand elle est publiée depuis le repo desktop, elle relie automatiquement les boutons de téléchargement aux GitHub Releases du dépôt.

## Rapports premium

À chaque synchronisation, le compagnon met à jour automatiquement :
- `projects/<CODE>/reports/Project_Report.docx`
- `projects/<CODE>/reports/Project_Register.xlsx`
- `evidence.docx` ou `story.docx` dans chaque dossier d’élément

Le bouton **Régénérer rapports** dans l’application Windows permet aussi de forcer une reconstruction complète.
