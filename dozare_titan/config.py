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
    {"id": "zy4", "nume": "Zircaloy 4 (Zy4)"},
    {"id": "snMetal", "nume": "Sn metal"},
]
ORDINE_MAT = [m["id"] for m in MATERIALE]


def _director_resurse():
    """Folderul de resurse (assets) al aplicatiei — functioneaza atat rulat
    ca script Python cat si compilat cu PyInstaller (caz in care fisierele
    adaugate cu --add-data "dozare_titan/assets<sep>dozare_titan/assets",
    unde <sep> e ";" pe Windows si ":" pe Linux/Mac, ajung in sys._MEIPASS)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "dozare_titan", "assets")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


# Logo-ul oficial, afisat in antetul documentelor generate (Fisa limita,
# RetDozare — .xlsx si .pdf). PANA PRIMESTI LOGO-UL OFICIAL, la aceasta cale
# se afla un PLACEHOLDER generat automat — inlocuieste DOAR fisierul
# "logo.png" din dozare_titan/assets/ cu imaginea oficiala (acelasi nume);
# codul nu trebuie schimbat. Daca fisierul lipseste, documentele revin
# automat la varianta text "ZIROM TITANIUM".
CALE_LOGO = os.path.join(_director_resurse(), "logo.png")

# ---------------------------------------------------------------------------
# Conturi de utilizator — SURSA VECHE, folosita DOAR o singura data, la
# prima pornire a aplicatiei, ca sa migreze automat in noul sistem de
# conturi (parole hash-uite, cereri de cont cu aprobare de admin, drepturi
# de editare/administrare) — vezi dozare_titan/conturi.py.
#
# Dupa prima pornire, conturile reale se gestioneaza din aplicatie:
#   - un utilizator nou isi cere cont din ecranul de autentificare
#     ("Creeaza cont nou"), introducand un nume de utilizator (scurt,
#     folosit doar la login) si un alias (numele complet, care ramane
#     afisat pe documentele generate — Fisa limita, RetDozare — la rubrica
#     "Intocmit", indiferent cat de scurt e numele de login);
#   - un cont cu drept de administrare aproba sau respinge cererea si
#     acorda drepturi (editare / administrare) din fereastra principala
#     ("Administrare conturi");
#   - datele sunt in .dozare_titan/conturi.json, NU mai in acest fisier.
#
# Aceasta lista ramane doar ca sablon pentru migrarea initiala — daca
# .dozare_titan/conturi.json exista deja, e ignorata complet. Sef Sectie
# Lingouri ramane FIX (SEF_SECTIE_LINGOURI mai jos), indiferent cine e logat.
# ---------------------------------------------------------------------------
UTILIZATORI = [
    {"utilizator": "georgeta", "parola": "titan2026", "nume": "Racasanu Georgeta", "editor": True},
    {"utilizator": "vizitator", "parola": "vizitator", "nume": "Vizitator", "editor": False},
]

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
# (identice cu formularul pe hartie, inclusiv numele furnizorului asociat
# fiecarui material — vezi sablonul "FISA LIMITA cda ...").
FISA_LIMITA_ETICHETE = {
    "burete": "Burete Ti L Yang",
    "aliajAlV": "Prealiaj AlV-GfE",
    "alMetal": "Aluminiu",
    "feMetal": "Fier electrolitic",
    "tio2": "TiO2-Kronos",
    "aliajAlMo": "Aliaj AlMo",
    "aliajSiTi": "Prealiaj SiTi",
    "zrMetal": "Zr 702",
    "zy4": "Zircaloy 4 (Zy4)",
    "snMetal": "Sn metalic",
}

# Ordinea de afisare a materialelor pe formularul "Fisa limita -cda" — usor
# diferita de ORDINE_MAT (aici TiO2 apare inaintea Fierului electrolitic,
# ca in sablonul tiparit).
ORDINE_AFISARE_MATERIALE = [
    "burete", "aliajAlV", "alMetal", "tio2", "feMetal", "aliajAlMo", "aliajSiTi", "zrMetal",
    "zy4", "snMetal",
]
ORDINE_FISA_LIMITA = ORDINE_AFISARE_MATERIALE  # alias istoric

# Prescurtarile folosite in antetul tabelului "Bilant presare [Kg]" din
# RetDozare (coloane inguste, deci nume scurte — burete apare ca "Ti", nu
# ca "Burete de titan").
MATERIAL_ABREVIERE_BILANT = {
    "burete": "Ti",
    "aliajAlV": "AlV",
    "alMetal": "Al",
    "feMetal": "Fe",
    "tio2": "TiO2",
    "aliajAlMo": "AlMo",
    "aliajSiTi": "SiTi",
    "zrMetal": "Zr",
    "zy4": "Zy4",
    "snMetal": "Sn",
}

# Numele materialelor asa cum apar in tabelele "Loturi si compozitie
# initiala" / "Initial" din RetDozare — din nou usor diferite de numele
# din MATERIALE (mai descriptive, ca pe formularul tiparit).
MATERIAL_NUME_RETDOZARE = {
    "burete": "Burete Ti",
    "aliajAlV": "Aliaj Al V",
    "alMetal": "Al metal",
    "feMetal": "Fe metal",
    "tio2": "TiO2",
    "aliajAlMo": "Aliaj Al-Mo",
    "aliajSiTi": "Prealiaj Si-Ti",
    "zrMetal": "Zr metal",
    "zy4": "Zy 4",
    "snMetal": "Sn met.",
}

# Elementele chimice relevante pentru fiecare material, folosite ca sa
# stim ce celule raman GOALE (nu "0.0000%") in tabelul "Initial" — un
# material nu are sens sa arate, de ex., "% V" daca nu contine deloc V.
MATERIAL_ELEMENTE_RELEVANTE = {
    "burete": ["Ti", "O", "Fe"],
    "aliajAlV": ["Al", "V", "O", "Fe"],
    "alMetal": ["Al", "O", "Fe"],
    "feMetal": ["Fe"],
    "tio2": ["Ti", "O"],
    "aliajAlMo": ["Al", "Mo", "O", "Fe"],
    "aliajSiTi": ["Ti", "Si", "O", "Fe"],
    "zrMetal": ["Zr", "O", "Fe"],
    "zy4": ["Zr", "Sn", "O", "Fe"],
    "snMetal": ["Sn", "O"],
}

# Numele afisate implicit pe formulare, acolo unde formularul chiar
# contine niste nume fixe (sef de sectie / intocmit) sau un beneficiar
# implicit — pot fi editate ulterior in cod daca se schimba persoanele.
SEF_SECTIE_LINGOURI = "Chiru Dan"
INTOCMIT_NUME = "Racasanu Georgeta"
BENEFICIAR_IMPLICIT = "Zirom"

# Codurile formularelor tiparite (colt dreapta-jos), identice cu cele de pe
# formularele pe hartie folosite pana acum.
COD_FORMULAR_RETDOZARE = "PGQ 036.F8.00"
COD_FORMULAR_FISA_LIMITA = "PGQ 036.F10.00"

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

# Materialele pe care le foloseste FIECARE tip de aliaj (aceleasi ca in
# calcule/specificatii.py si calcule/ti6242.py). Un material care nu apare
# in lista aliajului nu cere lot si nu se calculeaza, chiar daca tinta
# elementului lui e > 0 (ex. la Ti 6-2-4-2 zirconiul vine din Zy4, deci
# "Zr metal (Zr 702)" nu trebuie cerut). Aliajele care lipsesc de aici
# pastreaza vechea regula (doar dupa tinta elementului).
MATERIALE_PE_ALIAJ = {
    "Ti5": ("burete", "aliajAlV", "alMetal", "tio2", "feMetal"),
    "Ti -VT9": ("burete", "aliajAlMo", "alMetal", "zrMetal", "feMetal",
                "aliajSiTi", "tio2"),
    "Ti6Al4V-AMS": ("burete", "aliajAlV", "alMetal", "tio2"),
    "Ti 6-2-4-2": ("burete", "aliajAlMo", "alMetal", "zrMetal", "zy4",
                   "snMetal", "aliajSiTi", "tio2"),
}

# Grupuri de materiale ALTERNATIVE in cadrul unui aliaj: la Ti 6-2-4-2,
# zirconiul poate veni fie din Zr metal (Zr 702, pur), fie din Zy4
# (Zircaloy 4, care aduce si putin Sn) — depinde ce lot e pe stoc, NU
# trebuie ambele deodata. Fiecare tuplu de mai jos e un grup: e nevoie de
# lot selectat la CEL PUTIN UNUL dintre membrii lui, nu la fiecare.
GRUPE_ALTERNATIVE_PE_ALIAJ = {
    "Ti 6-2-4-2": (("zrMetal", "zy4"),),
}
ALIAJE_SPEC = {
    "Ti5": {
        "nume": "Ti5",
        "titlu_formular": "AMS 4928 X",
        "grad": "Ti gr5",
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
        "titlu_formular": "VT9",
        "grad": "Ti VT9",
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
        "titlu_formular": "AMS 4928 X",
        "grad": "Ti gr5",
        "o_max": 0.002,
        "al_min": 0.055,
        "al_max": 0.0675,
        "v_min": 0.035,
        "v_max": 0.045,
    },
    # Ti 6-2-4-2 (Ti-6Al-2Sn-4Zr-2Mo-0.08Si). Cheia e IDENTICA cu numele din
    # coloana "Aliaj" din standarde.xlsx, deci standardele (AMS 4975 /
    # AMS 4976 ...) se leaga singure. Limitele sunt cele din Excel-ul de
    # productie "Cda 24883-Ti6242" (STANDARD AMS 4976), ca FRACTII.
    "Ti 6-2-4-2": {
        "nume": "Ti 6-2-4-2",
        "titlu_formular": "AMS 4976",
        "grad": "Ti 6242",
        "al_min": 0.055,
        "al_max": 0.065,
        "mo_min": 0.018,
        "mo_max": 0.022,
        "sn_min": 0.018,
        "sn_max": 0.022,
        "zr_min": 0.036,
        "zr_max": 0.044,
        "o_max": 0.0015,
        "si_min": 0.0006,
        "si_max": 0.001,
        "fe_max": 0.001,
    },
}
ALIAJE_DISPONIBILE = list(ALIAJE_SPEC.keys())


# ---------------------------------------------------------------------------
# Legatura cu standarde.xlsx (vezi standarde.py)
# ---------------------------------------------------------------------------
# Aliajele cu reteta implementata (cheile din ALIAJE_SPEC) se leaga de numele
# aliajului din coloana "Aliaj" a fisierului standarde.xlsx, ca sa primeasca
# lista de standarde. Ti5 si Ti6Al4V-AMS sunt ambele titan grad 5.
# Ti -VT9 nu apare in Excel, deci nu are standarde (ramane cu limitele vechi
# din ALIAJE_SPEC).
#
# Aliajele din Excel care NU sunt legate aici apar automat in aplicatie ca
# PLACEHOLDER-e (se pot alege pe comanda, cu standardele si limitele lor,
# dar fara calcul de dozare). Ca sa transformi un placeholder intr-un aliaj
# real: adauga-l in ALIAJE_SPEC si in calcule.DOZATOARE sub cheia EXACT egala
# cu numele din Excel (ex. "Titan grad 2"); legatura de mai jos nu mai e
# necesara pentru el.
ALIAJE_EXCEL = {
    "Ti5": "Titan grad 5 (Ti 6Al 4V)",
    "Ti6Al4V-AMS": "Titan grad 5 (Ti 6Al 4V)",
}
