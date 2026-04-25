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

Le dossier `docs/` contient une vraie page publique de téléchargement reliée aux GitHub Releases du dépôt.

### Recommandation

Pour obtenir une URL propre, utilise idéalement le nom de dépôt :

```text
grantproof-desktop-sync
```

Une fois GitHub Pages activé sur la branche `main` et le dossier `docs/`, l’adresse publique aura la forme :

```text
https://<votre-utilisateur>.github.io/grantproof-desktop-sync/
```

Cette adresse est celle qu’il faut afficher dans GrantProof Mobile sur l’écran **Installer le compagnon PC**.

## Sorties générées

Dans chaque projet synchronisé, le compagnon met à jour :
- `projects/<CODE>/reports/Project_Report.docx`
- `projects/<CODE>/reports/Project_Register.xlsx`
- `projects/<CODE>/reports/Project_Report_FR.docx`
- `projects/<CODE>/reports/Project_Report_EN.docx`
- `projects/<CODE>/reports/Project_Register_FR.xlsx`
- `projects/<CODE>/reports/Project_Register_EN.xlsx`
- `evidence.docx` ou `story.docx` dans chaque dossier d’élément

## Mise à jour rapports premium V7

Cette version génère les rapports Word directement côté Desktop Sync, sans dépendance cloud :

- page 1 au format dashboard premium, propriétaire de l’ONG, sans logo GrantProof dans le document ;
- nom de l’ONG récupéré depuis les métadonnées synchronisées quand il est disponible ;
- choix automatique d’une image principale parmi les médias liés ;
- ajout d’une page 2 uniquement lorsqu’il existe des images complémentaires ;
- détection métier du secteur humanitaire selon l’activité et la description (ex. abris, agriculture, sécurité alimentaire, EHA/WASH, santé, nutrition, éducation, protection, cash, coordination) ;
- génération automatique d’une carte du pays à partir d’un fond Natural Earth embarqué, avec point orange si les coordonnées GPS sont disponibles ou si la localisation est reconnue par défaut ;
- absence de champs vides de type N/A dans le rapport principal.

Le livrable reste un fichier `.docx` modifiable dans `projects/<CODE>/reports/`.
