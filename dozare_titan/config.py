"""Constante fixe ale aplicatiei: materiale, tipuri de aliaj, cai de fisiere.

Acesta e primul loc de verificat cand adaugi un tip de aliaj nou sau un
material nou (ex. o sursa de Si/Zr/Mo pentru Ti-VT9).
"""

import os
import sys


def _director_aplicatie():
    """Folderul in care se afla exe-ul (cand ruleaza ca aplicatie compilata
    cu PyInstaller) sau scriptul principal (cand ruleaza cu "python
    main.py"). Datele se salveaza langa aplicatie, nu in profilul
    utilizatorului, ca sa fie portabile (poti muta folderul cu exe-ul in
    alta parte si isi ia datele cu el)."""
    if getattr(sys, "frozen", False):
        # PyInstaller: sys.executable e chiar exe-ul (dozare_titan.exe).
        return os.path.dirname(sys.executable)
    # Rulare normala ca script Python: folderul care contine main.py
    # (radacina proiectului), indiferent de unde e lansata comanda.
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------------------
# Materiale de stoc
# ---------------------------------------------------------------------------
# ATENTIE: aceasta lista e comuna tuturor aliajelor in acest moment. Daca
# Ti-VT9 foloseste materiale diferite (ex. sursa de Si, sursa de Zr, sursa
# de Mo) care nu exista aici, trebuie:
#   1. adaugate ca intrari noi in MATERIALE (cu un "id" nou, unic);
#   2. adaugate campurile de doza corespunzatoare in dialoguri.py
#      (DialogLotNou -> etichete_doza);
#   3. folosite intr-un pas ("id") din specificatia aliajului, in
#      calcule/specificatii.py (vezi SPEC_VT9 acolo pentru un exemplu).
MATERIALE = [
    {"id": "burete", "nume": "Burete de titan"},
    {"id": "aliajAlV", "nume": "Aliaj Al-V"},
    {"id": "alMetal", "nume": "Al metal"},
    {"id": "feMetal", "nume": "Fe metal"},
    {"id": "tio2", "nume": "TiO\u2082"},
    {"id": "aliajAlMo", "nume": "Aliaj AlMo"},
    {"id": "aliajSiTi", "nume": "Prealiaj SiTi"},
    {"id": "zrMetal", "nume": "Zr metal (Zr 702)"},
]
ORDINE_MAT = [m["id"] for m in MATERIALE]

ADMIN_PASS = "titan2026"

# Folder ASCUNS (nume cu punct in fata, ca pe Linux/Mac), asezat langa
# exe-ul aplicatiei (sau langa main.py, daca ruleaza ca script) — NU in
# profilul utilizatorului. Pe Windows, un folder cu punct in fata nu se
# ascunde automat din Explorer (asta e o conventie specifica Linux/Mac),
# dar numele ramane la fel; daca vrei sa fie ascuns si vizual pe Windows,
# seteaza-i atributul "Hidden" din proprietatile folderului dupa prima
# rulare (Explorer -> click dreapta pe folder -> Proprietati -> Ascuns).
DATA_DIR = os.path.join(_director_aplicatie(), ".dozare_titan")
DATA_FILE = os.path.join(DATA_DIR, "date.json")

# Etichetele materialelor asa cum apar pe formularul tiparit "Fisa limita -cda"
# (identice cu formularul pe hartie, inclusiv denumirea repetata "Aluminiu").
FISA_LIMITA_ETICHETE = {
    "burete": "Burete Ti",
    "aliajAlV": "Prealiaj AlV-GfE",
    "alMetal": "Aluminiu",
    "feMetal": "Aluminiu",
    "tio2": "TiO2",
    "aliajAlMo": "Aliaj AlMo",
    "aliajSiTi": "Prealiaj SiTi",
    "zrMetal": "Zr 702",
}

# ---------------------------------------------------------------------------
# Tipuri de aliaj — calculul de dozare/RetDozare e valabil per tip de aliaj.
# Cand apare un tip nou, se adauga o noua intrare aici (cu limitele ei
# chimice) SI o specificatie de dozare corespunzatoare:
#   1. adauga o specificatie noua in calcule/specificatii.py (dupa modelul
#      celor existente: SPEC_TI6AL4V, SPEC_VT9 etc.) — descrie ce materiale
#      intra in reteta si ce element chimic acopera fiecare;
#   2. inregistreaz-o in calcule/__init__.py -> DOZATOARE, sub cheia
#      EXACT IDENTICA cu cea folosita mai jos (ex. "Ti6Al4V-AMS" nu
#      "Ti6Al4V") — combobox-ul din UI populeaza order["tipAliaj"] direct
#      din ALIAJE_DISPONIBILE (cheile de mai jos), deci orice diferenta de
#      scriere intre cele doua locuri face ca acel aliaj sa nu poata fi
#      niciodata calculat.
# ---------------------------------------------------------------------------
ALIAJ_IMPLICIT = "Ti5"
ALIAJE_SPEC = {
    "Ti5": {
        "nume": "Ti5",
        "o_max": 0.002,
        "fe_max": 0.003,
        "n_max": 0.0005,
        "c_max": 0.0008,
        "h_max": 0.000125,
        "al_min": 0.055,
        "al_max": 0.0675,
        "v_min": 0.035,
        "v_max": 0.045,
    },
    "Ti -VT9": {
        "nume": "Ti -VT9",
        "o_max": 0.015,
        "si_min": 0.20,
        "si_max": 0.35,
        "fe_max": 0.10,
        "zr_min": 1.0,
        "zr_max": 2.0,
        "mo_min": 2.8,
        "mo_max": 3.8,
        "al_min": 5.8,
        "al_max": 7.0,
    },
  "Ti6Al4V-AMS": {
        "nume": "Ti6Al4V-AMS",
        "o_max": 0.002,
        "al_min": 0.055,
        "al_max": 0.0675,
        "v_min": 0.035,
        "v_max": 0.045,
    },
}
ALIAJE_DISPONIBILE = list(ALIAJE_SPEC.keys())
