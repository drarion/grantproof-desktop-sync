# GrantProof Desktop Sync — plan V1

Objectif : recevoir les preuves du mobile sur un ordinateur du client, puis les écrire dans un dossier local `GrantProof`.

## V1 cible
- Windows d'abord
- ordinateur allumé + logiciel lancé
- même réseau local que le téléphone
- jumelage par QR code
- dossier local choisi par l'utilisateur
- aucun stockage GrantProof

## Distribution gratuite au départ
- repo GitHub public dédié au desktop sync
- installateur publié via GitHub Releases
- page de téléchargement publiée via GitHub Pages

## Prochaine brique technique
- service HTTP local léger sur le PC
- QR code contenant le nom du poste + l'URL locale + un code de jumelage
- envoi manuel déclenché depuis le mobile
