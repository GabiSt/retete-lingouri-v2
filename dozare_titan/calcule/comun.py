"""Functii de calcul folosite indiferent de tipul de aliaj.

Aceste functii lucreaza cu structura actuala de materiale/elemente
(Al, V, O, Fe, Ti). Daca formula unui aliaj nou (ex. Ti-VT9) are nevoie de
elemente in plus (Si, Zr, Mo...), cel mai probabil va trebui sa extinzi si
functiile de aici (in special lot_ti/target_ti, care presupun ca Ti e
"restul" din Al+V+O+Fe+N) sau sa scrii variante proprii in calcule/vt9.py
si sa le folosesti direct de acolo, fara sa modifici acest fisier.
"""

from ..config import ORDINE_MAT
from ..utils import to_float


def lot_ti(lot):
    """Calculeaza procentul de Titan din lot.

    Doar Burete Ti si TiO2 contin Ti.
    """
    material = lot.get("material", "")

    # Burete Ti si Prealiaj SiTi: Ti = 100 - (toate celelalte elemente dozate)
    # (pentru burete, dozaMo/dozaSi/dozaZr sunt normal 0; pentru aliajSiTi,
    # dozaAl/dozaV/dozaMo/dozaZr sunt normal 0 — ramane 100 - O - Fe - Si)
    if material == "burete" or material == "aliajSiTi":
        return 100 - (
            to_float(lot.get("dozaAl")) + to_float(lot.get("dozaV")) +
            to_float(lot.get("dozaO")) + to_float(lot.get("dozaFe")) +
            to_float(lot.get("dozaN")) + to_float(lot.get("dozaMo")) +
            to_float(lot.get("dozaSi")) + to_float(lot.get("dozaZr"))
        )

    # TiO2: are ~60% Ti (40% O)
    elif material == "tio2":
        return 60.0

    # Aliaj AlV / AlMo, Al metal, Fe metal, Zr metal: nu contin Ti
    elif material in ("aliajAlV", "aliajAlMo", "alMetal", "feMetal", "zrMetal"):
        return 0.0

    # Default
    return 0.0


def lot_rest(lot):
    return to_float(lot.get("stocIntrare")) - to_float(lot.get("consum"))


def target_ti(target):
    return 100 - (
        to_float(target.get("al")) + to_float(target.get("v")) +
        to_float(target.get("o")) + to_float(target.get("fe")) +
        to_float(target.get("n", 0)) + to_float(target.get("mo")) +
        to_float(target.get("si")) + to_float(target.get("zr"))
    )


# Materialul din stanga se calculeaza din balanta elementului din dreapta.
# Daca tinta acelui element e 0%, materialul nu mai e necesar in amestec —
# nu trebuie selectat lot pentru el, e pur si simplu ignorat.
ELEMENT_PER_MATERIAL = {
    "aliajAlV": "v",
    "alMetal": "al",
    "tio2": "o",
    "feMetal": "fe",
    "aliajAlMo": "mo",
    "aliajSiTi": "si",
    "zrMetal": "zr",
}


def material_necesar(mat_id, target):
    """True daca materialul chiar trebuie sa aiba un lot selectat.

    Burete de titan e mereu necesar (e incarcatura de baza a retetei);
    celelalte materiale sunt necesare doar daca tinta elementului lor
    asociat e diferita de 0%.
    """
    elem = ELEMENT_PER_MATERIAL.get(mat_id)
    if elem is None:
        return True
    return abs(to_float(target.get(elem))) > 1e-9


def _componente_loturi(lot_sel, target):
    """Extrage compozitia (%) fiecarui lot selectat, indexata dupa material.

    Materialele care nu sunt necesare (tinta elementului asociat = 0%) nu
    au nevoie de lot selectat — se trateaza cu compozitie zero si sunt
    ignorate din calcul.
    """
    comps = {}
    for k in ORDINE_MAT:
        l = lot_sel.get(k)
        if not l:
            if material_necesar(k, target):
                return None, f"Selecteaza un lot pentru {k}."
            comps[k] = {
                "Al": 0.0, "V": 0.0, "O": 0.0, "Fe": 0.0, "Ti": 0.0,
                "Mo": 0.0, "Si": 0.0, "Zr": 0.0,
            }
            continue
        comps[k] = {
            "Al": to_float(l.get("dozaAl")),
            "V": to_float(l.get("dozaV")),
            "O": to_float(l.get("dozaO")),
            "Fe": to_float(l.get("dozaFe")),
            "Ti": lot_ti(l),
            "Mo": to_float(l.get("dozaMo")),
            "Si": to_float(l.get("dozaSi")),
            "Zr": to_float(l.get("dozaZr")),
        }
    return comps, None


def _rezumat_calcul(target, comps, lot_sel, rezultat, p, n=1):
    """Construieste dictionarul de rezultat afisat in UI, plecand de la
    cantitatile pe material.

    "rezultat" e dozarea pentru O SINGURA presare; "rezultatTotal" e dozarea
    inmultita cu numarul de presari (n) — aceasta e cantitatea REALA care se
    scade din stoc pentru bara respectiva.

    Aceasta functie e comuna tuturor aliajelor: orice formula noua (ex.
    calcule/vt9.py) trebuie doar sa returneze un dict {material_id: kg}
    pentru o singura portie — restul (compozitia rezultata, verificarea de
    stoc etc.) se calculeaza aici, la fel pentru toate aliajele.
    """
    rezultat_total = {k: v * n for k, v in rezultat.items()}

    negativ = any(rezultat_total[k] < -0.005 for k in rezultat_total)

    depaseste_stoc = False
    for k in rezultat_total:
        lot = lot_sel.get(k)
        if lot and rezultat_total[k] > lot_rest(lot) + 0.001:
            depaseste_stoc = True

    masa_totala = sum(rezultat.values())

    ti_rezultat = 0
    al_rezultat = 0
    v_rezultat = 0
    o_rezultat = 0
    fe_rezultat = 0
    mo_rezultat = 0
    si_rezultat = 0
    zr_rezultat = 0
    for k in rezultat:
        ti_rezultat += rezultat[k] * comps[k]["Ti"] / 100
        al_rezultat += rezultat[k] * comps[k]["Al"] / 100
        v_rezultat += rezultat[k] * comps[k]["V"] / 100
        o_rezultat += rezultat[k] * comps[k]["O"] / 100
        fe_rezultat += rezultat[k] * comps[k]["Fe"] / 100
        mo_rezultat += rezultat[k] * comps[k]["Mo"] / 100
        si_rezultat += rezultat[k] * comps[k]["Si"] / 100
        zr_rezultat += rezultat[k] * comps[k]["Zr"] / 100

    divizor = masa_totala if masa_totala > 1e-9 else 1
    return {
        "rezultat": rezultat,
        "rezultatTotal": rezultat_total,
        "nrPresari": n,
        "negativ": negativ or depaseste_stoc,
        "tiRezultat": ti_rezultat / divizor * 100,
        "tiTarget": target_ti(target),
        "portie": p,
        "portieEfectiva": masa_totala,
        "portieTotala": masa_totala * n,
        "compozitie_rezultata": {
            "Al": al_rezultat / divizor * 100,
            "V": v_rezultat / divizor * 100,
            "O": o_rezultat / divizor * 100,
            "Fe": fe_rezultat / divizor * 100,
            "Ti": ti_rezultat / divizor * 100,
            "Mo": mo_rezultat / divizor * 100,
            "Si": si_rezultat / divizor * 100,
            "Zr": zr_rezultat / divizor * 100,
        }
    }
