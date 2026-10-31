# Veille jurisprudentielle automatisée en droit social

Chaque lundi matin, ce projet récupère les nouveaux arrêts de la chambre sociale de la Cour de cassation publiés au Bulletin, les fait résumer par un LLM et envoie une note de veille au format Word par e-mail.

Conçu lors d'un stage dans un cabinet d'avocats en droit social (clients : CSE, salariés, syndicats), pour remplacer une veille manuelle chronophage. L'outil est en production au sein du cabinet. Ce dépôt en est une copie publique, publiée avec son accord : le workflow n'y est pas actif et aucun identifiant n'y est configuré.

## Fonctionnement

```
GitHub Actions (lundi 06:00 UTC)
        │
        ▼
API Judilibre (OAuth 2.0, portail PISTE)
  /export : arrêts de la chambre sociale des 7 derniers jours
  filtrage des arrêts publiés au Bulletin
        │
        ▼
API Claude (Haiku)
  synthèse de chaque arrêt : faits, apport, intérêt pour les clients
        │
        ▼
Note Word (.docx) ──► e-mail aux juristes + archive téléchargeable
```

## Choix techniques

- **Prompt contraint** : le modèle résume uniquement à partir du texte de l'arrêt, sans inventer de référence, et signale explicitement les informations manquantes. Chaque synthèse se termine par un rappel de validation par un juriste.
- **Filtrage du Bulletin côté script** : le filtre serveur `publication=b` de l'API renvoyait 0 résultat à tort. Le script récupère donc tous les arrêts de la période puis filtre lui-même.
- **Modèle Haiku** : le moins coûteux de la gamme Claude, suffisant pour de la synthèse de texte fourni.
- **Robustesse** : nouvelles tentatives automatiques sur les erreurs temporaires de l'API Claude, et un arrêt en échec n'interrompt pas la note.
- **Sécurité** : aucun identifiant dans le code. Tout passe par des variables d'environnement en local et par les secrets GitHub en production.

## Stack

Python 3.12 · requests · python-docx · API Judilibre (OAuth 2.0) · API Anthropic · SMTP · GitHub Actions

Code développé avec l'assistance de Claude Code. Analyse du besoin du cabinet, choix fonctionnels, conception du prompt, tests sur arrêts réels et identification des limites des LLM : réalisés par mes soins.

## Installation

### En local (test)

```bash
pip install -r requirements.txt

export JUDILIBRE_ENV="sandbox"        # données fictives, pour tester la chaîne
export JUDILIBRE_CLIENT_ID="..."
export JUDILIBRE_CLIENT_SECRET="..."
export ANTHROPIC_API_KEY="..."

python veille_sociale.py
```

Sous Windows (PowerShell), remplacer `export NOM="..."` par `$env:NOM="..."`.

### En production (GitHub Actions)

Dans **Settings > Secrets and variables > Actions**, créer les secrets suivants :

| Secret | Rôle | Obligatoire |
|---|---|---|
| `JUDILIBRE_CLIENT_ID` | Client ID de l'application PISTE | oui |
| `JUDILIBRE_CLIENT_SECRET` | Client Secret de l'application PISTE | oui |
| `ANTHROPIC_API_KEY` | Clé API Anthropic | oui |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD` | Serveur d'envoi des e-mails | pour l'envoi |
| `MAIL_FROM`, `MAIL_TO` | Expéditeur et destinataires (séparés par des virgules) | pour l'envoi |

Sans les secrets SMTP, la note est uniquement archivée dans l'onglet **Actions** du dépôt. Le workflow peut aussi être lancé à la main depuis cet onglet (**Run workflow**).

Pour obtenir les identifiants Judilibre : créer un compte sur [PISTE](https://piste.gouv.fr), créer une application et y associer l'API Judilibre.

## Limites

- Les synthèses sont générées par un LLM : elles doivent être relues par un juriste avant toute diffusion.
- La sandbox Judilibre ne simule que la recherche par mot-clé, avec des données fictives. Seule la production permet une vraie veille.
- GitHub désactive les workflows planifiés d'un dépôt public après 60 jours sans activité.
