"""Formula de dozare pentru Ti6Al4V.

Reproduce EXACT calculul din Excel (CDA 036, "Ti6Al4V AMS", foaia
RetDozare"), pe un lant de formule secventiale — nu prin rezolvarea
unui siste"m cu solutie unica.

ATENTIE: reteta veche (versiunea anterioara a acestui fisier) corespundea
de fapt aliajului Ti5, nu Ti6Al4V. Aceasta versiune e aliniata pe CDA-ul
real de Ti6Al4V (AMS) si NU mai contine Fe metal ca material dozat.

Ordinea de calcul (identica cu Excel, randul 27 din "RetDozare"):
  1. Aliaj Al-V se calculeaza din balanta de V (singurul material cu V).
  2. Al metal se calculeaza din balanta de Al, dupa ce se scade
     contributia de Al din Aliajul Al-V — PARTICULARITATE: scaderea NU
     se mai imparte la %Al din Al metal (=A27*D23-C27*C17, in Excel).
     Practic se presupune ca Al metalul e ~100% pur; eroarea introdusa
     e neglijabila cat timp lotul chiar e foarte pur (>99.8%).
  3. Burete Ti = portia de baza - Aliaj Al-V - Al metal (balanta de masa
     a incarcaturii principale).
  4. TiO2 se adauga suplimentar pentru a acoperi balanta de O — aici
     formula e "curata", intreaga balanta se imparte la %O din TiO2.

Fe metal NU mai e un pas al acestei formule — CDA-ul de Ti6Al4V (AMS)
nu are nicio coloana de dozare pentru el (spre deosebire de reteta veche,
gresit atribuita acestui aliaj).
"""

from ..utils import to_float


def calculeaza_dozare_kg(target, comps, p):
    """Aplica formulele Excel (CDA 036 / RetDozare) pentru o portie de p kg.

    Parametri:
      target -- compozitia chimica tinta a retetei (dict cu chei
                "al", "v", "o", in procente). Cheia "fe" nu se mai
                foloseste la aceasta reteta.
      comps  -- compozitia (%) fiecarui lot selectat, indexata dupa
                material (vezi calcule.comun._componente_loturi)
      p      -- portia, in kg

    Returneaza (rezultat, eroare):
      rezultat -- dict {material_id: kg} pentru O SINGURA portie, sau None
      eroare   -- mesaj de eroare (string), sau None daca s-a calculat OK
    """
    v_tinta_kg = to_float(target.get("v")) / 100 * p
    al_tinta_kg = to_float(target.get("al")) / 100 * p
    o_tinta_kg = to_float(target.get("o")) / 100 * p

    # 1. Aliaj Al-V din balanta de V (ignorat daca tinta V = 0%)
    if abs(v_tinta_kg) < 1e-9:
        aliaj_kg = 0.0
    else:
        v_alv = comps["aliajAlV"]["V"]
        if abs(v_alv) < 1e-9:
            return None, "Lotul de Aliaj Al-V nu are %V definit — nu se poate calcula."
        aliaj_kg = v_tinta_kg / (v_alv / 100)

    # 2. Al metal din balanta de Al (ignorat daca tinta Al = 0%). Se pastreaza
    #    EXACT forma din Excel: =A27*D23-C27*C17 — scaderea contributiei
    #    Aliajului Al-V NU se mai imparte la %Al din Al metal.
    al_din_alv = aliaj_kg * comps["aliajAlV"]["Al"] / 100
    if abs(al_tinta_kg) < 1e-9:
        al_metal_kg = 0.0
    else:
        al_metal_kg = al_tinta_kg - al_din_alv

    # 3. Burete Ti = restul incarcaturii principale
    burete_kg = p - aliaj_kg - al_metal_kg

    # 4. TiO2 adaugat suplimentar, din balanta de O (ignorat daca tinta O = 0%)
    o_din_burete = burete_kg * comps["burete"]["O"] / 100
    o_din_alv = aliaj_kg * comps["aliajAlV"]["O"] / 100
    o_din_al = al_metal_kg * comps["alMetal"]["O"] / 100
    if abs(o_tinta_kg) < 1e-9:
        tio2_kg = 0.0
    else:
        o_tio2 = comps["tio2"]["O"]
        if abs(o_tio2) < 1e-9:
            return None, "Lotul de TiO2 nu are %O definit — nu se poate calcula."
        tio2_kg = (o_tinta_kg - o_din_burete - o_din_alv - o_din_al) / (o_tio2 / 100)

    rezultat = {
        "burete": round(burete_kg, 3),
        "aliajAlV": round(aliaj_kg, 3),
        "alMetal": round(al_metal_kg, 3),
        "tio2": round(tio2_kg, 3),
    }
    return rezultat, None
