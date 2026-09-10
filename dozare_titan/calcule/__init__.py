"""Punctul unic de intrare pentru calculul de dozare a unei bare.

calculeaza_bara() alege AUTOMAT formula corecta in functie de tipul de
aliaj al comenzii (tip_aliaj), din registrul DOZATOARE de mai jos.

De la integrarea motorului generic (calcule/generic.py), formulele NU mai
sunt scrise separat pe fisier per aliaj (fostele ti6al4v.py / vt9.py /
ti5.py). Fiecare tip de aliaj e acum doar o SPECIFICATIE declarativa
(calcule/specificatii.py) care descrie ce materiale intra in reteta si ce
element chimic acopera fiecare, in ce ordine. Motorul generic calculeaza
mereu balanta de masa COMPLETA (scade INTAI, imparte DUPA) — fara
scurtaturile de precedenta a operatorilor gasite in Excel-urile vechi
(vezi calcule/generic.py pentru detalii si motivul acestei schimbari).

Comportamentul observabil pentru "Ti5" si "Ti -VT9" e neschimbat in ceea
ce priveste MATERIALELE afisate; singura diferenta e ca acum cateva
materiale (in special TiO2) pot iesi cu cateva grame diferit fata de
formulele vechi, pentru ca balanta e acum completa si corecta (formulele
vechi aveau o particularitate de precedenta care lasa neimpartite niste
scaderi — vezi test_generic.py din pachetul motor_test pentru validare pe
date reale).

Cand adaugi un aliaj nou:
  1. Adauga o specificatie noua in specificatii.py (dupa modelul celor
     existente), cu materialele si elementele corespunzatoare.
  2. Inregistreaz-o aici, in DOZATOARE, sub cheia EXACTA folosita in
     config.ALIAJE_SPEC.
Restul aplicatiei (UI, export RetDozare/Fisa limita) foloseste doar
calculeaza_bara() si nu trebuie modificat.
"""

from .generic import calculeaza_generic
from . import specificatii as spec
from .comun import (  # noqa: F401  (re-exportate pentru comoditate)
    lot_ti,
    lot_rest,
    target_ti,
    material_necesar,
    ELEMENT_PER_MATERIAL,
    _componente_loturi,
    _rezumat_calcul,
)
from ..utils import to_float


def _dozator(specificatie):
    """Construieste o functie calculeaza_dozare_kg(target, comps, p),
    legata de o specificatie declarativa data, gata de inregistrat in
    DOZATOARE."""
    def calculeaza_dozare_kg(target, comps, p):
        return calculeaza_generic(target, comps, p, specificatie)
    return calculeaza_dozare_kg


# Registrul de formule: tip_aliaj (identic cu cheile din config.ALIAJE_SPEC)
# -> functie calculeaza_dozare_kg(target, comps, p) -> (rezultat, eroare)
#
#   "Ti5"          -> SPEC_TI6AL4V:         burete + Aliaj Al-V + Al metal +
#                                            TiO2 + Fe metal (acelasi set de
#                                            materiale ca vechiul calcule/ti5.py)
#   "Ti -VT9"      -> SPEC_VT9:             burete + Aliaj Al-Mo + Al metal +
#                                            Zr metal + Fe metal + Prealiaj
#                                            SiTi + TiO2 (acelasi set ca
#                                            vechiul calcule/vt9.py)
#   "Ti6Al4V-AMS"  -> SPEC_TI6AL4V_FARA_FE: burete + Aliaj Al-V + Al metal +
#                                            TiO2, FARA Fe metal (acelasi
#                                            set ca vechiul calcule/ti6al4v.py
#                                            — CDA-ul real de Ti6Al4V-AMS nu
#                                            foloseste Fe metal)
#
# ATENTIE: cheia trebuie sa fie IDENTICA cu cea din config.ALIAJE_SPEC (nu
# cu numele "prescurtat" al aliajului) — combobox-ul din UI populeaza
# order["tipAliaj"] direct din config.ALIAJE_DISPONIBILE, deci o cheie
# diferita aici (ex. "Ti6Al4V" in loc de "Ti6Al4V-AMS") inseamna ca acel
# tip de aliaj pica mereu pe eroarea "Nu exista o formula de dozare...".
DOZATOARE = {
    "Ti5": _dozator(spec.SPEC_TI6AL4V),
    "Ti -VT9": _dozator(spec.SPEC_VT9),
    "Ti6Al4V-AMS": _dozator(spec.SPEC_TI6AL4V_FARA_FE),
}


def calculeaza_bara(tip_aliaj, target, lot_sel, portie, nr_presari=1):
    """Calculeaza dozarea unei bare, pentru tipul de aliaj dat.

    Parametri:
      tip_aliaj  -- cheia tipului de aliaj (ex. "Ti6Al4V", "Ti -VT9"),
                    de regula order["tipAliaj"]
      target     -- compozitia chimica tinta a retetei
      lot_sel    -- loturile selectate pe material (dict material_id -> lot)
      portie     -- portia, in kg
      nr_presari -- numarul de presari (dozarea per portie se inmulteste
                    cu acesta pentru a obtine consumul total din stoc)

    Returneaza un dict cu rezultatul calculului, sau {"eroare": "..."}
    daca ceva nu a putut fi calculat.
    """
    comps, eroare = _componente_loturi(lot_sel, target)
    if eroare:
        return {"eroare": eroare}

    p = to_float(portie)
    if p <= 0:
        return {"eroare": "Introdu o portie valida (kg)."}

    n = to_float(nr_presari, 1)
    if n <= 0:
        n = 1

    functie_dozare = DOZATOARE.get(tip_aliaj)
    if functie_dozare is None:
        return {"eroare": f"Nu exista o formula de dozare implementata pentru tipul de aliaj \u201e{tip_aliaj}\u201d."}

    rezultat, eroare = functie_dozare(target, comps, p)
    if eroare:
        return {"eroare": eroare}

    return _rezumat_calcul(target, comps, lot_sel, rezultat, p, n)
