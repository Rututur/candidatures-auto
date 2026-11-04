# Candidatures auto

Deux outils pour une recherche de stage (ou d'alternance, d'emploi…) :

1. **Une application Windows** qui envoie tes candidatures par lot depuis Gmail (CV + lettre PDF), prépare les relances et tient ton tableau Excel à jour.
2. **Une surveillance en ligne** (GitHub Actions, toutes les 10 min, même PC éteint) qui t'envoie un **message privé Discord** quand une entreprise répond, et un rappel quand il faut relancer.

Les deux partagent leur état via des **libellés Gmail** (`Candidatures/…`). La surveillance n'a donc jamais besoin de ton fichier Excel, et le dépôt GitHub ne contient aucune donnée personnelle.

```
app/        application de bureau (Tkinter)
common/     code partagé : Gmail (IMAP), détection des réponses, estimation refus/positif, Discord
watcher/    script lancé par GitHub Actions
tools/      setup_discord.py (à lancer une fois)
tests/      python -m unittest
```

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
- **Onglet Envoyer** : coche les entreprises (case à gauche). Pour chacune, clique dessus, choisis sa **lettre PDF** avec « Choisir… » et retouche le mail si tu veux. Le bouton du bas envoie tout le lot, avec une pause entre deux mails (réglable). Un récapitulatif te demande confirmation avant le départ.
- **Onglet À relancer** : les candidatures sans réponse depuis 7 jours (réglable). La relance est pré-remplie, elle part **dans le même fil** que le mail d'origine, et tu la modifies et la valides avant envoi.
- **Synchroniser avec Gmail** (automatique à l'ouverture) : retrouve les réponses et met le tableur à jour.

### Ce que l'outil change dans ton tableur

- Il ajoute deux colonnes, **Fichier lettre (PDF)** et **Relance envoyée le**, ainsi qu'un statut **Réponse reçue** (dans la liste déroulante, la couleur et la synthèse).
- **Après un envoi**, il met le statut à `Envoyée` (ou `Relancée`), le mode de contact à `Email` et la date du jour.
- **Quand une réponse arrive** : un refus probable donne le statut `Refusée`. Toute autre réponse donne `Réponse reçue`. Dans les deux cas, il remplit la date de réponse et ajoute une note `[Auto …]` à vérifier. Les réponses automatiques (absence, accusé de réception) sont ignorées. Un mail non distribué ajoute une note.
- **Date de relance** = dernier envoi (candidature ou relance) + 7 jours. Avant, c'était date de contact + 10 jours.
- Une copie `… (sauvegarde auto).xlsx` est faite avant la première modification de chaque session.
- **Ferme le fichier dans Excel** avant d'envoyer, sinon il ne peut pas être enregistré. L'outil te prévient et te propose de réessayer.

Les candidatures déjà envoyées avant l'outil (statut `Envoyée`/`Relancée`) sont retrouvées dans tes mails envoyés à la première synchronisation et suivies comme les autres.

## 3. Créer le bot Discord (10 min)

1. Sur <https://discord.com/developers/applications> → **New Application** (nom : « Candidatures »).
2. Menu **Bot** → **Reset Token** → copie le token (c'est un mot de passe : ne le partage pas).
3. Menu **OAuth2 → URL Generator** : coche le scope **bot** (aucune permission nécessaire), ouvre l'URL générée et ajoute le bot à **un serveur où tu es** (crée-en un vide si besoin). C'est obligatoire pour qu'il puisse t'écrire en privé.
4. Dans Discord → Paramètres → Avancés → active le **Mode développeur**. Puis clic droit sur ton nom → **Copier l'identifiant de l'utilisateur**.
5. Vérifie que tu acceptes les messages privés des membres de ce serveur (clic droit sur le serveur → Paramètres de confidentialité).
6. Dans le dossier du projet, ouvre un terminal et lance une seule fois :
   ```
   .venv\Scripts\python tools\setup_discord.py
   ```
   Le script connecte le bot (Discord l'exige une première fois) et t'envoie un message de test.

## 4. Mettre la surveillance sur GitHub (10 min)

1. Crée un dépôt **public** sur GitHub (ex. `candidatures-auto`) et envoie-y ce dossier. Le `.gitignore` exclut déjà les `.xlsx`, `.pdf` et `.venv`.
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<ton-pseudo>/candidatures-auto.git
   git push -u origin main
   ```
2. Dans le dépôt → **Settings → Secrets and variables → Actions** :
   - onglet **Secrets** → *New repository secret*, 4 fois :
     `GMAIL_ADDRESS`, `GMAIL_APP_PASSWORD`, `DISCORD_BOT_TOKEN`, `DISCORD_USER_ID`
   - onglet **Variables** (facultatif) : `REMINDER_DAYS` = `7` (mets la même valeur que dans l'application).
3. Onglet **Actions** → « Surveillance des réponses » → **Run workflow** pour un premier essai. Ensuite, ça tourne seul toutes les 10 minutes.

Bon à savoir :
- **Les secrets restent secrets**, même dans un dépôt public. Par contre, **les journaux des exécutions sont publics**. Le script n'y écrit que des nombres, jamais d'adresse, de nom ou d'objet.
- GitHub peut décaler une exécution de quelques minutes quand ses serveurs sont chargés.
- Au premier passage, les réponses de plus de 72 h sont marquées comme vues **sans** t'envoyer de message, pour ne pas te noyer sous l'historique.
- GitHub met en pause les tâches planifiées après 60 jours d'inactivité d'un dépôt. Le workflow se réactive lui-même à chaque passage pour l'éviter.

## Comment une réponse est reconnue

Un mail reçu est considéré comme une réponse s'il arrive **après** ta candidature et :
- soit qu'il est dans le **même fil Gmail**,
- soit qu'il vient de **la même adresse**,
- soit qu'il vient du **même domaine** (`rh@sodebo.fr` répond pour `contact@sodebo.fr`). Ce dernier cas ne s'applique pas aux adresses gmail.com, orange.fr, etc.

L'estimation refus / positif repose sur des mots-clés (« malheureusement », « pas donner suite », « entretien », « disponibilités »…) cherchés **uniquement dans le nouveau texte**, sans l'historique cité. **C'est une indication : lis toujours le mail.**

## Libellés Gmail créés

| Libellé | Rôle |
|---|---|
| `Candidatures/Envoyees` | posé sur chaque candidature et relance envoyée |
| `Candidatures/Notifiees` | réponse déjà signalée sur Discord |
| `Candidatures/Reponses-auto` | réponse automatique (absence, accusé de réception) |
| `Candidatures/Rappel-fait` | rappel de relance déjà envoyé pour ce mail |

Pour arrêter le suivi d'une entreprise, retire le libellé `Candidatures/Envoyees` de son mail dans Gmail.

## Tests

```
.venv\Scripts\python -m unittest
```

Pour tester la surveillance sur ton PC sans rien envoyer sur Discord ni modifier Gmail :
```
set GMAIL_ADDRESS=ton.adresse@gmail.com
set GMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx
set DRY_RUN=1
.venv\Scripts\python -m watcher.watch
```

## Limites

- Gmail limite l'envoi à environ 500 mails par jour depuis un compte personnel. Garde une pause d'au moins 20 s entre deux mails pour ne pas passer pour du spam.
- Les candidatures par formulaire web, téléphone ou courrier ne sont pas suivies : il n'y a pas de mail.
