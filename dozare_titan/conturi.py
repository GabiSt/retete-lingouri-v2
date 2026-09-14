"""Gestionare conturi de utilizator: creare, autentificare, aprobare si
drepturi — inlocuieste vechea lista fixa UTILIZATORI din config.py.

Structura unui cont (persistat in .dozare_titan/conturi.json):
    {
        "utilizator": "gina",              # numele de login, SCURT
        "alias": "Racasanu Georgeta",       # numele complet, afisat pe
                                             # documentele generate (Fisa
                                             # limita, RetDozare) la rubrica
                                             # "Intocmit" — ramane pe
                                             # documente indiferent cat de
                                             # scurt e numele de login
        "hash": "...",                      # hash-ul parolei (PBKDF2)
        "sare": "...",                      # sarea folosita la hash
        "editor": True/False,               # drept de editare date
        "admin": True/False,                # drept de administrare conturi
                                             # (aprobare, drepturi, stergere)
        "stare": "aprobat" | "in_asteptare" | "respins",
    }

Cheia din dictionarul salvat e utilizatorul (login) cu litere mici, ca sa
nu conteze majusculele la autentificare.

La prima rulare (daca nu exista inca conturi.json), conturile migreaza
automat din vechea lista config.UTILIZATORI, pastrand parolele si
acordand drept de administrare contului care avea deja "editor": True
(ca sa existe un admin de la inceput, fara pasi manuali suplimentari).
"""

import hashlib
import json
import os
import secrets

from .config import DATA_DIR, UTILIZATORI

CONTURI_FILE = os.path.join(DATA_DIR, "conturi.json")

STARE_APROBAT = "aprobat"
STARE_ASTEPTARE = "in_asteptare"
STARE_RESPINS = "respins"

_ITERATII_HASH = 200_000


# ---------------------------------------------------------------------------
# Hash parola
# ---------------------------------------------------------------------------
def hash_parola(parola, sare=None):
    """Calculeaza hash-ul PBKDF2-HMAC-SHA256 al unei parole. Daca 'sare' nu
    e dat, se genereaza una noua (aleatoare). Returneaza (hash_hex, sare_hex)."""
    if sare is None:
        sare = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", parola.encode("utf-8"), bytes.fromhex(sare), _ITERATII_HASH)
    return h.hex(), sare


def verifica_parola(parola, hash_asteptat, sare):
    if not hash_asteptat or not sare:
        return False
    h, _ = hash_parola(parola, sare)
    return secrets.compare_digest(h, hash_asteptat)


# ---------------------------------------------------------------------------
# Incarcare / salvare
# ---------------------------------------------------------------------------
def _migreaza_din_config():
    """Construieste dictionarul de conturi initial din vechea lista
    config.UTILIZATORI (folosita inainte de introducerea conturilor cu
    aprobare). Primul cont cu 'editor': True primeste si drept de admin,
    ca sa existe cineva care poate aproba conturi noi de la inceput."""
    conturi = {}
    admin_acordat = False
    for u in UTILIZATORI:
        h, sare = hash_parola(u["parola"])
        e_editor = bool(u.get("editor"))
        conturi[u["utilizator"].lower()] = {
            "utilizator": u["utilizator"],
            "alias": u.get("nume", u["utilizator"]),
            "hash": h,
            "sare": sare,
            "editor": e_editor,
            "admin": e_editor and not admin_acordat,
            "stare": STARE_APROBAT,
        }
        if e_editor:
            admin_acordat = True
    return conturi


def incarca_conturi():
    if os.path.exists(CONTURI_FILE):
        try:
            with open(CONTURI_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
        return {}
    # Prima rulare: migreaza din config.UTILIZATORI si salveaza imediat,
    # ca sa nu se piarda daca aplicatia se inchide inainte de prima scriere.
    conturi = _migreaza_din_config()
    salveaza_conturi(conturi)
    return conturi


def salveaza_conturi(conturi):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CONTURI_FILE, "w", encoding="utf-8") as f:
        json.dump(conturi, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Operatii
# ---------------------------------------------------------------------------
def gaseste_cont(conturi, utilizator):
    return conturi.get((utilizator or "").strip().lower())


def utilizator_exista(conturi, utilizator):
    return (utilizator or "").strip().lower() in conturi


def creeaza_cererecont(conturi, utilizator, parola, alias):
    """Adauga o cerere de cont noua, in asteptarea aprobarii unui admin.
    Fara drept de editare si fara drept de admin implicit. Presupune ca
    utilizator_exista() a fost deja verificat de apelant."""
    h, sare = hash_parola(parola)
    cheie = utilizator.strip().lower()
    conturi[cheie] = {
        "utilizator": utilizator.strip(),
        "alias": alias.strip(),
        "hash": h,
        "sare": sare,
        "editor": False,
        "admin": False,
        "stare": STARE_ASTEPTARE,
    }
    salveaza_conturi(conturi)
    return conturi[cheie]


def exista_admin(conturi):
    return any(c.get("admin") and c.get("stare") == STARE_APROBAT for c in conturi.values())
