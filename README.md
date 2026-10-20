# Sync KoboToolbox → Matrice de suivi des activités

Récupère les réponses du formulaire Kobo « Contrôle de chantier – Blocs sanitaires »,
met à jour la matrice (onglet `MATRICE DE SUIVI` + journal `HISTORIQUE`) et envoie chaque lundi
`MATRICE_DE_SUIVI_DES_ACTIVITES_du_JJ-MM-AA.xlsx` par e-mail.

## Fonctionnement
| Déclencheur | Action |
|---|---|
| Toutes les 3 h | Met à jour la matrice **seulement s'il y a de nouvelles réponses** (commit auto, pas d'e-mail) |
| Lundi 07:00 (Abidjan) | Met à jour, enregistre la matrice du jour et **l'envoie par e-mail** |
| Bouton « Run workflow » | Au choix, avec ou sans e-mail |
| Webhook Kobo (optionnel) | Mise à jour immédiate à chaque soumission |

La dernière matrice est toujours dans `matrices/` (les 8 plus récentes sont gardées ; l'historique complet est dans Git).

## Mise en place
1. Créer un dépôt GitHub **privé** et y pousser ce dossier.
2. *Settings → Secrets and variables → Actions → Secrets* : `KOBO_ASSET_UID`, `KOBO_TOKEN`,
   `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `MAIL_FROM`, `MAIL_TO` (adresses séparées par des virgules).
   *Variables* : `KOBO_URL` (ex. `https://kf.kobotoolbox.org` ou `https://eu.kobotoolbox.org`).
3. *Settings → Actions → General → Workflow permissions* : « Read and write permissions ».
4. Déployer `docs/questionnaire_blocs_sanitaires_corrige.xlsx` sur Kobo, récupérer l'UID du formulaire
   (dans l'URL) et le token (*Compte → Sécurité*).
5. Onglet *Actions → Sync Kobo → Run workflow* pour un premier test (cocher l'envoi d'e-mail pour tester le mail).

Gmail : utiliser un « mot de passe d'application » (validation en 2 étapes requise), `SMTP_HOST=smtp.gmail.com`, port 587.

## Mise à jour immédiate à chaque soumission (optionnel)
Kobo → Formulaire → *Paramètres → Services REST → Enregistrer un nouveau service* :
- URL : `https://api.github.com/repos/<OWNER>/<REPO>/dispatches`
- En-têtes personnalisés : `Authorization: Bearer <PAT>` (token GitHub fine-grained, droit *Contents: Read & write*),
  `Accept: application/vnd.github+json`
- Enveloppe personnalisée (si disponible) : `{"event_type":"kobo_submission"}`

Sans cette option, les 3 h de délai du planning s'appliquent.

## Test local
```bash
pip install -r requirements.txt
cp .env.example .env   # à compléter
python kobo_matrice_sync.py --no-mail
python kobo_matrice_sync.py --from-json soumissions.json --no-mail   # hors ligne
```
