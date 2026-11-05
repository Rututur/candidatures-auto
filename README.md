# Candidatures auto

Deux outils pour une recherche de stage (ou d'alternance, d'emploi…) :

1. **Une application Windows** qui envoie tes candidatures par lot depuis Gmail (CV + lettre PDF), prépare les relances et tient ton tableau Excel à jour.
2. **Une surveillance en ligne** (GitHub Actions, même PC éteint) qui t'envoie un **message privé Discord** quand une entreprise répond, et un rappel quand il faut relancer.

Les deux partagent leur état via des **libellés Gmail** (`Candidatures/…`). La surveillance n'a donc jamais besoin de ton fichier Excel, et le dépôt GitHub ne contient aucune donnée personnelle.

```
app/        application de bureau (Tkinter)
common/     code partagé : Gmail (IMAP), détection des réponses, estimation refus/positif, Discord
watcher/    script lancé par GitHub Actions
tools/      setup_discord.py (à lancer une fois)
tests/      python -m unittest
```

---

## Démarrage rapide : l'ordre à suivre

Le bot ne peut pas fonctionner tant que ces 5 points ne sont pas faits, dans cet ordre :

| # | À faire | Comment vérifier que c'est bon |
|---|---|---|
| 1 | Créer le mot de passe d'application Gmail ([§1](#1-préparer-gmail-5-min)) | Tu as 16 lettres |
| 2 | Lancer `lancer.bat` et remplir les réglages ([§2](#2-installer-lapplication-sur-ton-pc)) | « Tester la connexion Gmail » réussit |
| 3 | Créer le bot et lancer `setup_discord.py` ([§3](#3-créer-le-bot-discord-10-min)) | Tu reçois un message privé du bot |
| 4 | Ajouter les 4 secrets sur GitHub ([§4](#4-mettre-la-surveillance-sur-github-10-min)) | « Run workflow » finit en vert |
| 5 | Envoyer une vraie candidature depuis l'application | Le journal affiche `Tracked recipients: 1` ou plus |

> **Le bot ne se lance pas depuis ton PC.** C'est GitHub qui l'exécute. Ton PC sert seulement à l'étape de préparation et aux tests. Si une commande échoue, va voir le [dépannage](#dépannage).

### Dans quel terminal taper les commandes ?

Dans VS Code : menu **Terminal → Nouveau terminal**. Toutes les commandes se lancent **depuis le dossier du projet** (celui qui contient `lancer.bat`).

Regarde le début de la ligne : `PS C:\...>` veut dire **PowerShell**, et `C:\...>` veut dire **invite de commandes (cmd)**. La syntaxe des variables n'est pas la même : `set NOM=valeur` **ne fonctionne pas dans PowerShell**, et n'affiche même pas d'erreur.

---

## 1. Préparer Gmail (5 min)

1. Active la **validation en 2 étapes** sur ton compte Google : <https://myaccount.google.com/security>.
2. Crée un **mot de passe d'application** : <https://myaccount.google.com/apppasswords> (nom : « Candidatures »). Tu obtiens 16 lettres : garde-les pour l'étape 2 et pour GitHub.
3. Dans Gmail → Paramètres → *Transfert et POP/IMAP*, vérifie que **l'accès IMAP** est activé (si l'option apparaît).

> Ce mot de passe donne accès à ta boîte mail. Ne le colle jamais dans un fichier du projet. Tu peux le révoquer à tout moment sur la même page.

## 2. Installer l'application sur ton PC

Il faut **Python 3.10 ou plus** (<https://www.python.org/downloads/>, coche « Add python.exe to PATH »).

1. Dézippe le dossier où tu veux.
2. Double-clique sur **`lancer.bat`**. Au premier lancement, il installe ce qu'il faut (environ 30 s).
3. La fenêtre **Réglages** s'ouvre : adresse Gmail, mot de passe d'application, nom et téléphone (signature), **CV en PDF**, période du stage. Clique sur **Tester la connexion Gmail**, puis sur Enregistrer.
   - Le mot de passe est rangé dans le *Gestionnaire d'identifiants Windows*, pas dans un fichier.
4. Choisis ton tableau de suivi (`Suivi_candidatures_stage_BTS_SIO.xlsx`).

### Utilisation

- **Mode test (activé par défaut)** : tous les mails partent vers ta propre adresse avec `[TEST]` dans l'objet, et le tableur n'est pas modifié. Fais un essai, regarde le rendu dans ta boîte, puis décoche le mode test (en haut à droite).
- **Page Envoyer** : coche les entreprises (case à gauche). Pour chacune, clique dessus, choisis sa **lettre PDF** avec « Choisir… » et retouche le mail si tu veux. Le bouton du bas envoie tout le lot, avec une pause entre deux mails (réglable). Un récapitulatif te demande confirmation avant le départ.
- **Page À relancer** : les candidatures sans réponse depuis 7 jours (réglable). La relance est pré-remplie, elle part **dans le même fil** que le mail d'origine, et tu la modifies et la valides avant envoi.
- **Menu de gauche** : passe d'une page à l'autre. Le nombre entre parenthèses indique ce qui t'attend. **Mode sombre / Mode clair** change l'apparence, et ton choix est mémorisé.
- **Synchroniser Gmail** (automatique à l'ouverture et après chaque envoi réel) : retrouve les réponses et les range dans l'onglet **À valider**. Cette étape ne change aucun statut du tableur.
- **Page À valider** : pour chaque mail reçu, lis le texte (ou clique sur *Ouvrir dans Gmail*), puis clique sur **Approuver** ou choisis une autre réponse : *Refus*, *Réponse à traiter*, *Entretien prévu*, *Acceptée*, *Mail non distribué* ou *Ignorer*. C'est seulement à ce moment que le tableur est modifié. Une entreprise dont la réponse attend ta validation n'apparaît pas dans « À relancer ».
- **Page Tableur** : affiche ton tableur dès l'ouverture de l'application, avec une couleur par statut. Double-clique sur une case pour la modifier : ton fichier Excel est mis à jour aussitôt (ferme-le d'abord dans Excel). « Date de relance » et « Action » sont calculées et ne se modifient pas. Pour ajouter une ligne, utilise Excel.
- **Rechercher, filtre par statut et Colonnes…** : trouve une entreprise et choisis les colonnes visibles. Ça ne change que l'affichage : ton fichier garde toutes ses colonnes.
- **Fichier → Modifier dans Excel** : ouvre ton tableur dans Excel. Enregistre et ferme Excel, puis choisis **Fichier → Recharger**. Si tu oublies, l'application recharge le fichier et te demande de refaire ta modification, pour éviter d'écrire sur la mauvaise ligne.
- **Fichier → Changer de fichier…** : seulement si tu veux utiliser un autre tableur. L'application rouvre toute seule le dernier utilisé.
- **Fichier → Enregistrer une copie…** : enregistre une copie du tableur où tu veux. Le fichier reste sur ton PC : il n'est jamais envoyé sur GitHub ni ailleurs.

### Ce que l'outil change dans ton tableur

- Il ajoute deux colonnes, **Fichier lettre (PDF)** et **Relance envoyée le**, ainsi qu'un statut **Réponse reçue** (dans la liste déroulante, la couleur et la synthèse).
- **Après un envoi**, il met le statut à `Envoyée` (ou `Relancée`), le mode de contact à `Email` et la date du jour.
- **Quand tu valides une réponse** : le statut devient celui que tu as choisi, la date de réponse est remplie si elle est vide, et une note `[Validé …]` est ajoutée. *Mail non distribué* ajoute seulement une note. *Ignorer* ne change rien au tableur, et la relance reste prévue. Les notes `[Auto …]` viennent d'une ancienne version : si tu ignores une de ces réponses, son statut est annulé.
- **Date de relance** = dernier envoi (candidature ou relance) + 7 jours. Avant, c'était date de contact + 10 jours.
- Une copie `… (sauvegarde auto).xlsx` est faite avant la première modification de chaque session.
- **Ferme le fichier dans Excel** avant d'envoyer, de valider une réponse ou de modifier une case dans l'application, sinon il ne peut pas être enregistré. L'outil te prévient.

Les candidatures déjà envoyées avant l'outil (statut `Envoyée`/`Relancée`) sont retrouvées dans tes mails envoyés à la première synchronisation et suivies comme les autres.

## 3. Créer le bot Discord (10 min)

1. Sur <https://discord.com/developers/applications> → **New Application** (nom : « Candidatures »).
2. Menu **Bot** → **Reset Token** → copie le token (c'est un mot de passe : ne le partage pas).
3. Menu **OAuth2 → URL Generator** : coche le scope **bot** (aucune permission nécessaire), ouvre l'URL générée et ajoute le bot à **un serveur où tu es** (crée-en un vide si besoin). C'est obligatoire pour qu'il puisse t'écrire en privé.
4. Dans Discord → Paramètres → Avancés → active le **Mode développeur**. Puis clic droit sur ton nom → **Copier l'identifiant de l'utilisateur**.
5. Vérifie que tu acceptes les messages privés des membres de ce serveur (clic droit sur le serveur → Paramètres de confidentialité).
6. Lance `lancer.bat` au moins une fois si ce n'est pas fait (il crée le dossier `.venv`). Ensuite, dans le dossier du projet, ouvre un terminal et lance une seule fois (la commande est la même dans PowerShell et cmd) :
   ```
   .venv\Scripts\python tools\setup_discord.py
   ```
   Le script te demande le token, puis ton identifiant. **Rien ne s'affiche quand tu colles le token, c'est normal.** Colle-le une seule fois (clic droit si Ctrl+V ne fait rien), puis appuie sur Entrée.

   Le script connecte le bot (Discord l'exige une première fois) et t'envoie un message de test. Tu dois voir `Message de test envoyé` et recevoir le message sur Discord.

## 4. Mettre la surveillance sur GitHub (10 min)

1. Crée un dépôt **public** sur GitHub (ex. `candidatures-auto`) et envoie-y ce dossier. Le `.gitignore` exclut déjà les `.xlsx`, `.pdf` et `.venv`. Remplace `<ton-pseudo>` par ton pseudo GitHub, sans les chevrons.
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<ton-pseudo>/candidatures-auto.git
   git push -u origin main
   ```
   **Si le dépôt existe déjà**, n'utilise pas ces commandes. Pour envoyer tes modifications, il suffit de faire :
   ```
   git add .
   git commit -m "Description de la modification"
   git push
   ```
2. Dans le dépôt → **Settings → Secrets and variables → Actions** :
   - onglet **Secrets** → *New repository secret*, 4 fois. Les noms doivent être **exactement** ceux-ci :

     | Nom du secret | Valeur |
     |---|---|
     | `GMAIL_ADDRESS` | ton adresse Gmail |
     | `GMAIL_APP_PASSWORD` | les 16 lettres de l'étape 1 (pas ton mot de passe Google) |
     | `DISCORD_BOT_TOKEN` | le token de l'étape 3 |
     | `DISCORD_USER_ID` | ton identifiant Discord, une suite de chiffres (pas ton pseudo) |

     Si tu refais « Reset Token » sur Discord, l'ancien token ne marche plus : mets aussi le secret à jour.
   - onglet **Variables** (facultatif) : *New repository variable*, avec `REMINDER_DAYS` dans **Name** et seulement `7` dans **Value** (mets la même valeur que dans l'application).
3. Onglet **Actions** → « Surveillance des réponses » → **Run workflow** pour un premier essai. Ensuite, ça tourne seul.
4. Rafraîchis la page, clique sur l'exécution, puis sur **watch**, puis sur l'étape **Check Gmail and notify on Discord** pour lire le résultat :

   | Ce que tu lis | Ce que ça veut dire |
   |---|---|
   | Coche verte, `Tracked recipients: 0` | La connexion à Gmail marche et les 4 secrets existent, mais il n'y a **rien à surveiller pour l'instant** : voir ci-dessous. |
   | Coche verte, `Tracked recipients: 3` (par exemple) | La connexion à Gmail marche et 3 adresses sont suivies. |
   | Croix rouge | Le message d'erreur explique pourquoi : voir le [dépannage](#dépannage). |

   Une coche verte ne prouve pas que `DISCORD_BOT_TOKEN` et `DISCORD_USER_ID` sont corrects : le bot ne contacte Discord que lorsqu'il a un message à envoyer. Seul le premier message reçu le confirme (voir [Tester toute la chaîne](#tester-toute-la-chaîne-facultatif)).

### Pourquoi `Tracked recipients: 0` ?

Le bot surveille seulement les mails qui portent le libellé `Candidatures/Envoyees`. Ce libellé est posé :
- quand tu envoies une candidature depuis l'application, **hors mode test** ;
- quand l'application se synchronise, sur les candidatures déjà envoyées à la main si leur statut est `Envoyée` ou `Relancée` dans le tableur.

Les mails du **mode test ne sont pas suivis** : ils partent vers ta propre adresse et ne reçoivent pas le libellé.

### Tester toute la chaîne (facultatif)

1. Dans le tableur, ajoute une ligne de test avec **une autre adresse à toi** (pas celle du compte Gmail utilisé par l'outil).
2. Décoche le mode test et envoie-lui une candidature depuis l'application. **Vérifie bien le récapitulatif : seule cette ligne doit être cochée.**
3. Depuis cette autre adresse, réponds au mail.
4. Sur GitHub, clique sur **Run workflow** : tu reçois le message privé sur Discord.
5. Supprime la ligne de test du tableur et retire le libellé `Candidatures/Envoyees` de ce mail dans Gmail.

Bon à savoir :
- **Les secrets restent secrets**, même dans un dépôt public. Par contre, **les journaux des exécutions sont publics**. Le script n'y écrit que des nombres, jamais d'adresse, de nom ou d'objet.
- **« Toutes les 10 minutes » est un maximum, pas une garantie.** GitHub peut retarder ou abandonner des exécutions planifiées, parfois pendant plusieurs heures. Pour un vrai passage toutes les 10 minutes, suis [la section ci-dessous](#vérifier-vraiment-toutes-les-10-minutes-recommandé). Pour tester, utilise **Run workflow**.
- Au premier passage, les réponses de plus de 72 h sont marquées comme vues **sans** t'envoyer de message, pour ne pas te noyer sous l'historique.
- GitHub met en pause les tâches planifiées après 60 jours d'inactivité d'un dépôt. Le workflow se réactive lui-même à chaque passage pour l'éviter.

### Vérifier vraiment toutes les 10 minutes (recommandé)

L'horaire de GitHub n'est pas fiable. Un service gratuit, cron-job.org, peut demander à GitHub de lancer le bot toutes les 10 minutes, comme si tu cliquais sur *Run workflow*. Il ne lit ni tes mails ni tes secrets. L'horaire de GitHub reste en place comme secours.

**1. Crée un token GitHub limité à ce dépôt** sur <https://github.com/settings/personal-access-tokens/new> :
- **Token name** : `cron-candidatures` ;
- **Expiration** : une date après la fin de ta recherche ;
- **Repository access** : *Only select repositories*, puis ton dépôt `candidatures-auto` ;
- **Permissions** : *Actions* sur **Read and write**, et rien d'autre ;
- clique sur **Generate token** et copie-le. Il ne s'affiche qu'une fois, et ne le mets jamais dans un fichier du projet.

**2. Crée la tâche sur <https://cron-job.org>** (compte gratuit), avec **Create cronjob** :
- **URL** (remplace `<ton-pseudo>`) :
  ```
  https://api.github.com/repos/<ton-pseudo>/candidatures-auto/actions/workflows/watch.yml/dispatches
  ```
- **Execution schedule** : toutes les 10 minutes ;
- dans les réglages avancés, **Request method** : `POST` ;
- **Headers** (remplace `<token>`) :
  ```
  Authorization: Bearer <token>
  Accept: application/vnd.github+json
  X-GitHub-Api-Version: 2022-11-28
  Content-Type: application/json
  ```
- **Request body** :
  ```json
  {"ref":"main"}
  ```

**3. Vérifie** : lance un test depuis cron-job.org, qui doit répondre `204`. Une nouvelle ligne doit apparaître dans l'onglet **Actions** de GitHub. Si elle y apparaît, c'est bon.

| Réponse de cron-job.org | Cause | Solution |
|---|---|---|
| `401` | Token incorrect ou expiré | Recrée le token et remplace-le dans le header |
| `403` | Permission *Actions* absente ou en lecture seule | Mets-la sur *Read and write* |
| `404` | URL incorrecte, ou le token n'a pas accès à ce dépôt | Vérifie l'URL et *Repository access* |
| `422` | Le corps de la requête est incorrect | Vérifie `{"ref":"main"}` et le header `Content-Type` |

Pour tout arrêter, désactive la tâche sur cron-job.org et supprime le token sur GitHub.

## Comment une réponse est reconnue

Un mail reçu est considéré comme une réponse s'il arrive **après** ta candidature et :
- soit qu'il est dans le **même fil Gmail**,
- soit qu'il vient de **la même adresse**,
- soit qu'il vient du **même domaine** (`rh@sodebo.fr` répond pour `contact@sodebo.fr`). Ce dernier cas ne s'applique pas aux adresses gmail.com, orange.fr, etc.

L'estimation refus / positif repose sur des mots-clés (« malheureusement », « pas donner suite », « entretien », « disponibilités »…) cherchés **uniquement dans le nouveau texte**, sans l'historique cité. **C'est seulement une proposition : c'est toi qui décides dans la page À valider.** Discord affiche la même estimation, mais la validation se fait uniquement dans l'application.

## Libellés Gmail créés

| Libellé | Rôle |
|---|---|
| `Candidatures/Envoyees` | posé sur chaque candidature et relance envoyée |
| `Candidatures/Notifiees` | réponse déjà signalée sur Discord |
| `Candidatures/Reponses-auto` | réponse automatique (absence, accusé de réception) |
| `Candidatures/Rappel-fait` | rappel de relance déjà envoyé pour ce mail |
| `Candidatures/Validees` | réponse déjà validée dans l'application : elle n'est plus proposée |

Pour arrêter le suivi d'une entreprise, retire le libellé `Candidatures/Envoyees` de son mail dans Gmail.

## Tests

```
.venv\Scripts\python -m unittest
```

### Tester la surveillance sur ton PC

Ce test lit ta boîte Gmail sans rien envoyer sur Discord et sans modifier Gmail. Il vérifie seulement la partie Gmail. Les variables ne durent que le temps du terminal : ferme-le après le test.

**Dans PowerShell** (`PS C:\...>`, le terminal par défaut de VS Code) :
```powershell
$env:GMAIL_ADDRESS = "ton.adresse@gmail.com"
$env:GMAIL_APP_PASSWORD = Read-Host "Mot de passe d'application"
$env:DRY_RUN = "1"
.venv\Scripts\python -m watcher.watch
```
`Read-Host` évite que le mot de passe soit enregistré dans l'historique de PowerShell, mais il reste visible à l'écran pendant que tu le tapes.

**Dans l'invite de commandes** (`C:\...>`, cmd) :
```bat
set GMAIL_ADDRESS=ton.adresse@gmail.com
set GMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx
set DRY_RUN=1
.venv\Scripts\python -m watcher.watch
```

Résultat attendu : `Tracked recipients: 0` si aucun mail n'est suivi pour l'instant, sinon ce nombre suivi des messages `[dry-run]` et d'une ligne `Replies found: ...`.

## Dépannage

| Message | Cause | Solution |
|---|---|---|
| `Missing environment variable GMAIL_ADDRESS` **sur ton PC** | Tu as utilisé `set` dans PowerShell, ou tu as ouvert un nouveau terminal | Utilise `$env:GMAIL_ADDRESS = "..."` ([voir ci-dessus](#tester-la-surveillance-sur-ton-pc)) |
| `Missing environment variable ...` **sur GitHub** | Le secret n'existe pas ou son nom est mal écrit | Crée-le dans *Settings → Secrets and variables → Actions*, onglet **Secrets** (pas *Variables*) |
| `IMAP login failed: [AUTHENTICATIONFAILED]` | Mauvaise adresse, ou mot de passe Google normal au lieu du mot de passe d'application | Recrée un mot de passe d'application et mets à jour le secret |
| `Special folder \All not found` | Le dossier « Tous les messages » est masqué pour IMAP | Gmail → Paramètres → *Libellés* → coche « Afficher dans IMAP » pour Tous les messages et Messages envoyés |
| `Discord a refusé le token du bot` ou `HTTP 401` | Token incorrect, ou régénéré depuis | Refais *Reset Token*, relance `setup_discord.py` et mets à jour le secret |
| `HTTP 403` avec `Cannot send messages to this user` | Le bot n'est pas sur un serveur où tu es, ou tes messages privés sont fermés | Refais les points 3 et 5 du [§3](#3-créer-le-bot-discord-10-min) |
| `HTTP 400` ou `HTTP 404` sur `/users/@me/channels` | Identifiant Discord incorrect | Refais le point 4 du [§3](#3-créer-le-bot-discord-10-min) : une suite de chiffres, pas ton pseudo |
| `Connexion à Discord impossible` | Coupure réseau, VPN ou pare-feu | Vérifie ta connexion, puis relance le script |
| `.venv\Scripts\python` n'est pas reconnu | `.venv` n'existe pas, ou tu n'es pas dans le bon dossier | Lance `lancer.bat` une fois, et place-toi dans le dossier qui le contient |
| `No module named 'watcher'` ou `'common'` | Tu n'es pas dans le dossier du projet | Place-toi dans le dossier qui contient `lancer.bat` |
| `No module named 'websocket'` | Tu utilises un autre Python que celui du projet | Lance `lancer.bat` une fois, et garde `.venv\Scripts\python` au début de la commande |
| `error: remote origin already exists` | Le dépôt est déjà relié à GitHub | Rien à refaire : utilise seulement `git add`, `git commit` et `git push` |
| Coche verte mais aucun message Discord | Aucune réponse nouvelle à signaler, ou `Tracked recipients: 0`, ou réponse de plus de 72 h (marquée sans message) | [Pourquoi `Tracked recipients: 0` ?](#pourquoi-tracked-recipients-0) |

## Limites

- Gmail limite l'envoi à environ 500 mails par jour depuis un compte personnel. Garde une pause d'au moins 20 s entre deux mails pour ne pas passer pour du spam.
- Les candidatures par formulaire web, téléphone ou courrier ne sont pas suivies : il n'y a pas de mail.
