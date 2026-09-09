"""Constante fixe ale aplicatiei: materiale, tipuri de aliaj, cai de fisiere.

Acesta e primul loc de verificat cand adaugi un tip de aliaj nou sau un
material nou (ex. o sursa de Si/Zr/Mo pentru Ti-VT9).
"""

import os

# ---------------------------------------------------------------------------
# Materiale de stoc
# ---------------------------------------------------------------------------
# ATENTIE: aceasta lista e comuna tuturor aliajelor in acest moment. Daca
# Ti-VT9 foloseste materiale diferite (ex. sursa de Si, sursa de Zr, sursa
# de Mo) care nu exista aici, trebuie:
#   1. adaugate ca intrari noi in MATERIALE (cu un "id" nou, unic);
#   2. adaugate campurile de doza corespunzatoare in dialoguri.py
#      (DialogLotNou -> etichete_doza);
#   3. folosite in formula ta din calcule/vt9.py.
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

DATA_DIR = os.path.join(os.path.expanduser("~"), ".dozare_titan")
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
# Comenzile de pana acum au fost toate pentru Ti6Al4V; cand apar alte tipuri,
# se adauga o noua intrare aici (cu limitele ei chimice) si un generator
# de formula dedicat in calcule/ (vezi calcule/vt9.py pentru punctul de
# start al Ti-VT9).
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
