"""Punctul unic de intrare pentru calculul de dozare a unei bare.

calculeaza_bara() alege AUTOMAT formula corecta in functie de tipul de
aliaj al comenzii (tip_aliaj), din registrul DOZATOARE de mai jos.

Cand adaugi un aliaj nou:
  1. Scrie formula lui intr-un fisier nou (dupa modelul ti6al4v.py),
     cu semnatura calculeaza_dozare_kg(target, comps, p) -> (rezultat, eroare).
  2. Inregistreaz-o aici, in DOZATOARE, sub cheia EXACTA folosita in
     config.ALIAJE_SPEC (ex. "Ti -VT9").
Restul aplicatiei (UI, export RetDozare/Fisa limita) foloseste doar
calculeaza_bara() si nu trebuie modificat.
"""

from . import ti5
from . import vt9
from . import ti6al4v
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

# Registrul de formule: tip_aliaj (identic cu cheile din config.ALIAJE_SPEC)
# -> functie calculeaza_dozare_kg(target, comps, p) -> (rezultat, eroare)
DOZATOARE = {
    "Ti5": ti5.calculeaza_dozare_kg,
    "Ti -VT9": vt9.calculeaza_dozare_kg,
    "Ti6Al4V": ti6al4v.calculeaza_dozare_kg,
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
