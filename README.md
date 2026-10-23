# Workflow de demandes d'attestations administratives — UDL-SBA

Application Streamlit en deux espaces :

1. **Demandeur** — remplit son nom, prénom, date et lieu de naissance, choisit son département et la langue souhaitée, puis soumet la demande.
2. **Chef de département** — ouvre son espace sécurisé, voit uniquement les demandes de son département, complète les champs pédagogiques et administratifs et émet le document Word arabe, français ou bilingue.

Départements : Électrotechnique, Automatique, Électronique, Télécommunications.

Le numéro officiel est attribué atomiquement au moment de la génération finale. La numérotation est annuelle et partagée entre les quatre départements. La date officielle utilise le fuseau Africa/Algiers. Le numéro et la date restent inchangés lorsqu'on retélécharge le même document.

## Fichiers à déposer dans GitHub

- `app.py`
- `attestation_generator.py`
- `storage.py`
- `requirements.txt`
- `supabase_schema.sql`
- `Modele_Source_Donnees_Attestations.xlsx`
- `assets/logo_udl.png`
- `.streamlit/config.toml`
- `.streamlit/secrets.toml.example` (modèle seulement)
- `.gitignore`

Ne déposez pas `app_single_stage_backup.py`, `.streamlit/secrets.toml`, une base SQLite, ni un fichier réel contenant des données personnelles. Le `.gitignore` bloque ces fichiers locaux.

## Lancer localement sous Windows

Dans le dossier où se trouve `app.py` :

```bash
py -m pip install -r requirements.txt
```

Copiez `.streamlit/secrets.toml.example` vers `.streamlit/secrets.toml`, puis remplacez les valeurs d'exemple par les vraies valeurs privées. Pour tester localement, les secrets nécessaires sont les mots de passe des chefs (ou `ADMIN_PASSWORD`). Sans configuration Supabase, l'application utilise une base SQLite locale `attestation_requests.sqlite3`.

Démarrez ensuite :

```bash
py -m streamlit run app.py
```

## Déploiement GitHub + Streamlit Community Cloud

1. Créez un dépôt GitHub **privé** et envoyez les fichiers listés ci-dessus.
2. Dans [Streamlit Community Cloud](https://share.streamlit.io/), sélectionnez le dépôt, la branche `main` et `app.py`.
3. Avant de rendre l'application accessible aux demandeurs, créez un projet Supabase. Dans **SQL Editor**, exécutez tout le fichier `supabase_schema.sql`.
4. Dans l'application Streamlit Cloud, ouvrez **Settings → Secrets** et ajoutez les secrets ci-dessous, remplacés par les vraies valeurs :

```toml
# Facultatif : administrateur global, qui peut traiter les quatre départements.
ADMIN_PASSWORD = "MOT_DE_PASSE_ADMIN_LONG_ET_UNIQUE"

# Mots de passe distincts pour chaque chef de département.
CHEF_PASSWORD_ELECTROTECHNIQUE = "MOT_DE_PASSE_CHEF_ELT_UNIQUE"
CHEF_PASSWORD_AUTOMATIQUE = "MOT_DE_PASSE_CHEF_AUTO_UNIQUE"
CHEF_PASSWORD_ELECTRONIQUE = "MOT_DE_PASSE_CHEF_ELN_UNIQUE"
CHEF_PASSWORD_TELECOMMUNICATIONS = "MOT_DE_PASSE_CHEF_TLC_UNIQUE"

# Supabase : garder la clé service-role uniquement dans les Secrets serveur.
SUPABASE_URL = "https://VOTRE_PROJET.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "VOTRE_CLE_SERVICE_ROLE"

# Facultatif : notification e-mail à la soumission d'une demande.
SMTP_HOST = "smtp.votre-etablissement.dz"
SMTP_PORT = "587"
SMTP_USER = "compte-notification@votre-etablissement.dz"
SMTP_PASSWORD = "MOT_DE_PASSE_SMTP"
SMTP_FROM = "compte-notification@votre-etablissement.dz"
PUBLIC_APP_URL = "https://votre-application.streamlit.app"
CHEF_EMAIL_ELECTROTECHNIQUE = "chef.elt@votre-etablissement.dz"
CHEF_EMAIL_AUTOMATIQUE = "chef.auto@votre-etablissement.dz"
CHEF_EMAIL_ELECTRONIQUE = "chef.eln@votre-etablissement.dz"
CHEF_EMAIL_TELECOMMUNICATIONS = "chef.tlc@votre-etablissement.dz"
```

Le nom du secret est sensible aux majuscules et doit être exact. Supabase est recommandé en ligne : le système d'exploitation d'une application Cloud peut remplacer son disque local, alors que la base Supabase conserve les demandes.

## Ce que fait Supabase

- Stocke les demandes, le statut, les informations saisies par le chef et le numéro/date officiels.
- Attribue de façon transactionnelle le numéro suivant pour l'année, sans collision si deux chefs génèrent presque simultanément.
- Réserve l'accès à la base à la clé **service role**, utilisée uniquement côté serveur. RLS est activé et aucune politique publique n'est créée.

**Sécurité de la clé :** ne mettez jamais `SUPABASE_SERVICE_ROLE_KEY` dans le code, le HTML ou GitHub. Conservez-la dans Streamlit Secrets seulement. Ne partagez pas `.streamlit/secrets.toml`.

## Courriel

Si les paramètres SMTP et `CHEF_EMAIL_*` sont présents, un courriel de notification est envoyé au chef lorsqu'une demande arrive. Le courriel contient le département, l'identifiant de suivi et le lien de l'application, mais pas les données de naissance du demandeur. Sans SMTP configuré, le dossier est tout de même enregistré dans l'espace protégé du chef; aucune notification e-mail n'est envoyée.

## Numéro et date officiels

- Le demandeur ne peut pas attribuer de numéro ou de date officiels.
- Le chef complète les champs puis clique sur **Générer l'attestation officielle**.
- Le numéro annuel partagé est attribué à ce moment-là, puis la date locale algérienne est enregistrée.
- Le numéro/date sont conservés pour un téléchargement ultérieur du même document.
- En cas de correction après émission, traitez cette correction selon la procédure administrative de l'établissement; l'application n'écrase pas en silence un document déjà émis.

## Modèle Word

Le générateur (`attestation_generator.py`) utilise le logo extrait du PDF et les structures françaises et arabes reconstruites à partir du scan. Les matières, dates, sections et années des semestres sont saisies par le chef. Le PDF fourni est une image scannée; toute validation officielle de la formulation et de la référence reste à faire par l'établissement.
