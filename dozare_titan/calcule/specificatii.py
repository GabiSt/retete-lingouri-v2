"""Specificatii declarative de dozare, per grad, pentru motorul generic
(calcule/generic.py). Fiecare intrare descrie DOAR structura retetei —
ce materiale exista si ce element acopera fiecare, in ce ordine trebuie
calculate. Loturile efective (comps) si tintele (target) vin din reteta
comenzii, ca la orice alt calcul.

Specificatiile SPEC_TI6AL4V, SPEC_VT9 etc. de mai jos sunt cele deja
validate pe date reale (CDA-uri VT9, Ti6242, Ti-834) fata de Excel-urile
existente. SPEC_TI6AL4V_FARA_FE e o varianta specifica acestei aplicatii,
adaugata pentru a pastra EXACT acelasi set de materiale pe care il avea
formula veche din calcule/ti6al4v.py (CDA real de Ti6Al4V-AMS, care nu
foloseste Fe metal ca material dozat).

Cand adaugi un aliaj nou in aplicatie:
  1. Adauga o specificatie noua aici (dupa modelul celor existente).
  2. Inregistreaz-o in calcule/__init__.py -> DOZATOARE, sub cheia EXACTA
     folosita in config.ALIAJE_SPEC.
"""

SPEC_TI6AL4V = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlV", "element": "V"},
        {"id": "alMetal", "element": "Al"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
        {"id": "feMetal", "element": "Fe", "dupa_burete": True},
    ],
}

# Varianta folosita de aplicatie pentru tipul de aliaj "Ti6Al4V" (CDA real
# de Ti6Al4V-AMS): identica cu SPEC_TI6AL4V, dar FARA pasul de Fe metal —
# acel CDA nu are nicio coloana de dozare pentru el.
SPEC_TI6AL4V_FARA_FE = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlV", "element": "V"},
        {"id": "alMetal", "element": "Al"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

SPEC_VT9 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "alMetal", "element": "Al"},
        {"id": "zrMetal", "element": "Zr"},
        {"id": "feMetal", "element": "Fe"},
        {"id": "aliajSiTi", "element": "Si"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

SPEC_TI_7AL_4MO = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "alMetal", "element": "Al"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
        {"id": "feMetal", "element": "Fe", "dupa_burete": True},
    ],
}

SPEC_TI6242 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "alMetal", "element": "Al"},
        {"id": "zrMetal", "element": "Zr"},
        {"id": "snMetal", "element": "Sn"},
        {"id": "aliajSiTi", "element": "Si"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

SPEC_TI6246 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "alMetal", "element": "Al"},
        {"id": "zrMetal", "element": "Zr"},
        {"id": "snMetal", "element": "Sn"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
        {"id": "feMetal", "element": "Fe", "dupa_burete": True},
    ],
}

SPEC_VT20 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "aliajAlV", "element": "V"},
        {"id": "alMetal", "element": "Al"},
        {"id": "zrMetal", "element": "Zr"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

SPEC_TI834 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlMo", "element": "Mo"},
        {"id": "aliajAlNb", "element": "Nb"},
        {"id": "alMetal", "element": "Al"},
        {"id": "zrMetal", "element": "Zr"},
        {"id": "snMetal", "element": "Sn"},
        {"id": "aliajSiTi", "element": "Si"},
        {"id": "carbon", "element": "C"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

# Grade "simple", fara aliaje-mama: doar burete + Fe metal (optional) + TiO2
SPEC_TI_CP_CU_FE = {
    "burete": "burete",
    "pasi": [
        {"id": "tio2", "element": "O", "dupa_burete": True},
        {"id": "feMetal", "element": "Fe", "dupa_burete": True},
    ],
}

SPEC_TI_CP_FARA_FE = {
    "burete": "burete",
    "pasi": [
        {"id": "tio2", "element": "O", "dupa_burete": True},
    ],
}

SPEC_TI43 = {
    "burete": "burete",
    "pasi": [
        {"id": "aliajAlV", "element": "V"},
        {"id": "alMetal", "element": "Al"},
        {"id": "snMetal", "element": "Sn"},
        {"id": "tio2", "element": "O", "dupa_burete": True},
        {"id": "feMetal", "element": "Fe", "dupa_burete": True},
    ],
}

SPEC_TI6AL4V_ELI = SPEC_TI6AL4V  # aceeasi structura (AlV + Al metal + TiO2 [+ Fe optional])
