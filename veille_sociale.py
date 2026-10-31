#!/usr/bin/env python3
"""
Veille juridique automatisée — droit social
============================================
1) Récupère les arrêts récents de la chambre sociale de la Cour de cassation,
   publiés au Bulletin ("B"), via l'API Judilibre (portail PISTE).
2) Pour chaque arrêt, génère une synthèse via l'API Claude (Anthropic).
3) Assemble une note de veille au format Word (.docx) et l'envoie par e-mail
   (facultatif).

L'API Judilibre exige une authentification OAuth 2.0 : on demande d'abord un
jeton avec le Client ID + Client Secret (onglet "Identifiants Oauth" sur PISTE),
puis on l'envoie dans l'en-tête "Authorization: Bearer ...".

LES IDENTIFIANTS NE SONT JAMAIS ÉCRITS EN DUR ICI. Le script les lit depuis des
variables d'environnement (en local) ou depuis les secrets GitHub (en CI) :
    JUDILIBRE_CLIENT_ID      -> ton Client ID PISTE
    JUDILIBRE_CLIENT_SECRET  -> ton Client Secret PISTE
    ANTHROPIC_API_KEY        -> ta clé API Anthropic (payante, distincte de l'abonnement chat)

Lancement (exemple, sous macOS/Linux) :
    export JUDILIBRE_ENV="sandbox"          # "sandbox" pour tester, "production" ensuite
    export JUDILIBRE_CLIENT_ID="ton-client-id"
    export JUDILIBRE_CLIENT_SECRET="ton-client-secret"
    export ANTHROPIC_API_KEY="ta-cle-anthropic"
    pip install -r requirements.txt
    python veille_sociale.py
"""

import os
import re
import ssl
import sys
import time
import smtplib
import datetime
from email.message import EmailMessage

import requests

try:
    from docx import Document
except ImportError:
    Document = None

# --- Configuration ---------------------------------------------------------

# Dans GitHub Actions, un secret non défini arrive sous forme de chaîne vide :
# on utilise donc "or" (et non la valeur par défaut de .get) pour les replis.
JUDILIBRE_ENV = os.environ.get("JUDILIBRE_ENV") or "sandbox"
JUDILIBRE_BASE = {
    "sandbox": "https://sandbox-api.piste.gouv.fr/cassation/judilibre/v1.0",
    "production": "https://api.piste.gouv.fr/cassation/judilibre/v1.0",
}[JUDILIBRE_ENV]
JUDILIBRE_OAUTH = {
    "sandbox": "https://sandbox-oauth.piste.gouv.fr/api/oauth/token",
    "production": "https://oauth.piste.gouv.fr/api/oauth/token",
}[JUDILIBRE_ENV]

JUDILIBRE_CLIENT_ID = os.environ.get("JUDILIBRE_CLIENT_ID")
JUDILIBRE_CLIENT_SECRET = os.environ.get("JUDILIBRE_CLIENT_SECRET")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"   # le moins coûteux (1$/5$ par M tokens), suffisant pour résumer

DAYS_BACK = int(os.environ.get("JUDILIBRE_DAYS") or "7")   # fenêtre en jours (modifiable via export)

# Nouvelles tentatives sur les erreurs temporaires de l'API Claude
# (surcharge, limite de débit, panne serveur).
CODES_TEMPORAIRES = {429, 500, 502, 503, 504, 529}
NB_ESSAIS = 3

# Envoi e-mail (facultatif : si ces variables sont absentes, le script se contente
# d'enregistrer le fichier Word sans l'envoyer).
SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT") or "587")
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")
MAIL_FROM = os.environ.get("MAIL_FROM") or SMTP_USER or ""
MAIL_TO = os.environ.get("MAIL_TO")   # un ou plusieurs destinataires séparés par des virgules

VEILLE_INSTRUCTIONS = (
    "Tu es un assistant de veille juridique spécialisé en droit social français, "
    "au service d'un cabinet qui conseille des CSE, des salariés et des syndicats. "
    "On te fournit le texte intégral d'un arrêt. Résume-le STRICTEMENT à partir de "
    "ce texte, sans inventer aucune référence ni aucun article. Format :\n"
    "- Faits (2 phrases max)\n"
    "- Apport / solution (distingue ce qui est confirmé de ce qui est nouveau)\n"
    "- Pourquoi c'est important pour des clients CSE / salariés / syndicats\n"
    "Si une information manque, écris \"non précisé dans le texte fourni\". "
    "Termine par : \"⚠️ À vérifier par un juriste avant diffusion.\""
)


def _check_keys():
    missing = [n for n, v in (("JUDILIBRE_CLIENT_ID", JUDILIBRE_CLIENT_ID),
                              ("JUDILIBRE_CLIENT_SECRET", JUDILIBRE_CLIENT_SECRET),
                              ("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY)) if not v]
    if missing:
        sys.exit("Variable(s) d'environnement manquante(s) : " + ", ".join(missing))


# --- Authentification OAuth 2.0 (PISTE) ------------------------------------

_TOKEN = None


def _auth_headers():
    """Retourne les en-têtes avec le jeton Bearer (obtenu une seule fois)."""
    global _TOKEN
    if _TOKEN is None:
        base = {"grant_type": "client_credentials", "scope": "openid"}
        # Méthode 1 : identifiants dans le corps de la requête (client_secret_post)
        r = requests.post(JUDILIBRE_OAUTH,
                          data={**base,
                                "client_id": JUDILIBRE_CLIENT_ID,
                                "client_secret": JUDILIBRE_CLIENT_SECRET},
                          timeout=30)
        # Méthode 2 (repli) : authentification HTTP Basic (client_secret_basic),
        # requise par beaucoup de clients "confidentiels".
        if r.status_code != 200:
            r = requests.post(JUDILIBRE_OAUTH, data=base,
                              auth=(JUDILIBRE_CLIENT_ID, JUDILIBRE_CLIENT_SECRET),
                              timeout=30)
        if r.status_code != 200:
            print(f"\n[Erreur OAuth PISTE · code {r.status_code}] Impossible d'obtenir un jeton.")
            print("Réponse :", r.text[:1000] or "(réponse vide)")
            sys.exit(1)
        _TOKEN = r.json().get("access_token")
    return {"accept": "application/json", "Authorization": f"Bearer {_TOKEN}"}


# --- Étape 1 : collecte des arrêts via Judilibre ---------------------------

def _requete(path, params, headers):
    r = requests.get(f"{JUDILIBRE_BASE}/{path}", params=params, headers=headers, timeout=30)
    if r.status_code != 200:
        print(f"\n[Erreur Judilibre · code {r.status_code}] L'API a refusé la requête.")
        print("URL appelée :", r.url)
        print("Message de l'API :", r.text[:1500] or "(réponse vide)")
        sys.exit(1)
    return r.json()


def _est_bulletin(arret):
    """Vrai si l'arrêt est publié au Bulletin (champ 'publication' contenant 'b')."""
    pub = arret.get("publication", [])
    if isinstance(pub, str):
        return pub == "b"
    return "b" in (pub or [])


def lister_arrets():
    """Collecte les arrêts.
    - En sandbox : appel /search minimal pour valider toute la chaîne (la sandbox
      ne simule que /search avec un mot-clé ; les données renvoyées sont fictives).
    - En production : /export sur la chambre sociale + dates, PUIS filtrage du
      Bulletin côté script (le filtre serveur 'publication=b' renvoie 0 à tort).
      Le texte intégral est déjà fourni par /export."""
    headers = _auth_headers()

    if JUDILIBRE_ENV == "sandbox":
        print("  (sandbox : test de la chaine via /search, donnees de demonstration)")
        data = _requete("search", {"query": "licenciement", "page_size": 5}, headers)
        return data.get("results", [])

    # Production : vraie veille
    date_fin = datetime.date.today()
    date_debut = date_fin - datetime.timedelta(days=DAYS_BACK)
    arrets, batch = [], 0
    while True:
        params = {
            "jurisdiction": "cc",                 # Cour de cassation
            "chamber": "soc",                     # chambre sociale
            "date_start": date_debut.isoformat(),
            "date_end": date_fin.isoformat(),
            "batch": batch,
            "batch_size": 100,
        }
        data = _requete("export", params, headers)
        arrets.extend(data.get("results", []))
        if not data.get("next_batch"):
            break
        batch += 1

    bulletin = [a for a in arrets if _est_bulletin(a)]
    print(f"  ({len(arrets)} arrêts soc. au total sur la période, dont {len(bulletin)} publiés au Bulletin)")
    return bulletin


def texte_integral(arret):
    """Récupère le texte intégral via /decision (ou le texte déjà présent)."""
    if arret.get("text"):
        return arret["text"]
    headers = _auth_headers()
    r = requests.get(f"{JUDILIBRE_BASE}/decision",
                     params={"id": arret.get("id")}, headers=headers, timeout=30)
    if r.status_code != 200:
        return ""
    return r.json().get("text", "")


# --- Étape 2 : synthèse via l'API Claude -----------------------------------

def resumer(texte):
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 1024,
        "system": VEILLE_INSTRUCTIONS,
        "messages": [{"role": "user", "content": texte}],
    }
    for essai in range(1, NB_ESSAIS + 1):
        try:
            r = requests.post("https://api.anthropic.com/v1/messages",
                              headers=headers, json=body, timeout=120)
        except (requests.ConnectionError, requests.Timeout):
            if essai == NB_ESSAIS:
                raise
        else:
            if r.status_code not in CODES_TEMPORAIRES or essai == NB_ESSAIS:
                break
        attente = 5 * 2 ** (essai - 1)              # 5 s, puis 10 s
        print(f"    API Claude indisponible, nouvel essai dans {attente} s...")
        time.sleep(attente)
    r.raise_for_status()
    data = r.json()
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


# --- Mise en forme Word (.docx) --------------------------------------------

def _ajouter_runs(paragraphe, texte):
    """Ajoute du texte en gérant le gras (**...**) et l'italique (*...*) Markdown."""
    pos = 0
    for m in re.finditer(r"\*\*(.+?)\*\*|\*(.+?)\*", texte):
        if m.start() > pos:
            paragraphe.add_run(texte[pos:m.start()])
        if m.group(1) is not None:
            paragraphe.add_run(m.group(1)).bold = True
        else:
            paragraphe.add_run(m.group(2)).italic = True
        pos = m.end()
    if pos < len(texte):
        paragraphe.add_run(texte[pos:])


def _ajouter_synthese(doc, texte):
    """Convertit une synthèse (Markdown léger) en paragraphes Word formatés."""
    for ligne in texte.splitlines():
        l = ligne.strip()
        if not l:
            continue
        if l.startswith("### "):
            doc.add_heading(l[4:].strip(), level=3)
        elif l.startswith("## "):
            doc.add_heading(l[3:].strip(), level=2)
        elif l.startswith("# "):
            doc.add_heading(l[2:].strip(), level=2)
        elif l[:2] in ("- ", "* "):
            _ajouter_runs(doc.add_paragraph(style="List Bullet"), l[2:].strip())
        else:
            _ajouter_runs(doc.add_paragraph(), l)


# --- Envoi e-mail (facultatif) ---------------------------------------------

def envoyer_email(chemin, nb):
    """Envoie le fichier Word en pièce jointe, si les variables SMTP sont définies."""
    if not (SMTP_HOST and SMTP_USER and SMTP_PASSWORD and MAIL_TO):
        print("  (envoi e-mail désactivé : variables SMTP non définies)")
        return
    today = datetime.date.today().isoformat()
    msg = EmailMessage()
    msg["Subject"] = f"Veille juridique droit social — {today}"
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO
    msg.set_content(
        f"Bonjour,\n\nCi-joint la note de veille du {today} "
        f"({nb} arrêt(s) publié(s) au Bulletin).\n\n"
        "Document généré automatiquement — à valider avant diffusion.\n"
    )
    with open(chemin, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(chemin),
        )
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as serveur:
        serveur.starttls(context=ssl.create_default_context())
        serveur.login(SMTP_USER, SMTP_PASSWORD)
        serveur.send_message(msg)
    print(f"  E-mail envoyé à {MAIL_TO}")


# --- Étape 3 : assemblage de la note ---------------------------------------

def main():
    _check_keys()
    if Document is None:
        sys.exit("La bibliothèque python-docx est requise. Installe-la avec :\n"
                 "    pip install python-docx")
    print(f"[Judilibre · {JUDILIBRE_ENV}] Recherche des arrêts soc. 'B' "
          f"des {DAYS_BACK} derniers jours...")
    arrets = lister_arrets()
    print(f"  -> {len(arrets)} arrêt(s) à traiter.")

    today = datetime.date.today().isoformat()
    doc = Document()
    doc.add_heading("Note de veille juridique — droit social", level=0)
    doc.add_paragraph(f"{today} · {len(arrets)} arrêt(s) publié(s) au Bulletin · "
                      "à valider avant diffusion")

    for i, arret in enumerate(arrets, 1):
        num = arret.get("number", "n° non précisé")
        date = arret.get("decision_date", "date non précisée")
        print(f"  [{i}/{len(arrets)}] Synthèse de l'arrêt {num}...")
        texte = texte_integral(arret)
        if not texte:
            synthese = "(Texte intégral indisponible via l'API — à récupérer manuellement.)"
        else:
            try:
                synthese = resumer(texte)
            except Exception as e:                   # noqa: BLE001
                synthese = f"(Erreur lors de la synthèse : {e})"
        doc.add_heading(f"{i}. Cass. soc. {date} — {num}", level=1)
        _ajouter_synthese(doc, synthese)

    fichier = f"veille_sociale_{today}.docx"
    doc.save(fichier)
    print(f"\nNote Word écrite dans : {fichier}")
    envoyer_email(fichier, len(arrets))
    print("Relis et valide chaque synthèse avant toute diffusion.")


if __name__ == "__main__":
    main()
