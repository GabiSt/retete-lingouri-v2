"""Formula de dozare pentru Ti6Al4V.

Reproduce EXACT calculul din Excel (foaia "Retete"), pe un lant de formule
secventiale — nu prin rezolvarea unui sistem 3x3 cu solutie unica.

Ordinea de calcul (identica cu Excel):
  1. Aliaj Al-V se calculeaza din balanta de V (singurul material cu V).
  2. Al metal se calculeaza din balanta de Al, dupa ce se scade
     contributia de Al din Aliajul Al-V.
  3. Burete Ti = portia de baza - Aliaj Al-V - Al metal (balanta de masa
     a incarcaturii principale).
  4. TiO2 se adauga suplimentar pentru a acoperi balanta de O.
  5. Fe metal se adauga suplimentar pentru a acoperi balanta de Fe.

TiO2 si Fe metal sunt doze mici, adaugate PESTE incarcatura principala
(ca in Excel), deci masa finala a barei poate fi usor peste "portie".
"""

from ..utils import to_float


def calculeaza_dozare_kg(target, comps, p):
    """Aplica formulele Excel pentru o portie de p kg.

    Parametri:
      target -- compozitia chimica tinta a retetei (dict cu chei
                "al", "v", "o", "fe", in procente)
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
    fe_tinta_kg = to_float(target.get("fe")) / 100 * p

    # 1. Aliaj Al-V din balanta de V (ignorat daca tinta V = 0%)
    if abs(v_tinta_kg) < 1e-9:
        aliaj_kg = 0.0
    else:
        v_alv = comps["aliajAlV"]["V"]
        if abs(v_alv) < 1e-9:
            return None, "Lotul de Aliaj Al-V nu are %V definit — nu se poate calcula."
        aliaj_kg = v_tinta_kg / (v_alv / 100)

    # 2. Al metal din balanta de Al, dupa ce scadem contributia Aliajului
    #    Al-V (ignorat daca tinta Al = 0%)
    al_din_alv = aliaj_kg * comps["aliajAlV"]["Al"] / 100
    if abs(al_tinta_kg) < 1e-9:
        al_metal_kg = 0.0
    else:
        al_am = comps["alMetal"]["Al"]
        if abs(al_am) < 1e-9:
            return None, "Lotul de Al metal nu are %Al definit — nu se poate calcula."
        al_metal_kg = (al_tinta_kg - al_din_alv) / (al_am / 100)

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

    # 5. Fe metal adaugat suplimentar, din balanta de Fe (ignorat daca tinta Fe = 0%)
    fe_din_burete = burete_kg * comps["burete"]["Fe"] / 100
    fe_din_alv = aliaj_kg * comps["aliajAlV"]["Fe"] / 100
    fe_din_al = al_metal_kg * comps["alMetal"]["Fe"] / 100
    if abs(fe_tinta_kg) < 1e-9:
        fe_metal_kg = 0.0
    else:
        fe_fe = comps["feMetal"]["Fe"]
        if abs(fe_fe) < 1e-9:
            return None, "Lotul de Fe metal nu are %Fe definit — nu se poate calcula."
        fe_metal_kg = (fe_tinta_kg - fe_din_burete - fe_din_alv - fe_din_al) / (fe_fe / 100)

    rezultat = {
        "burete": round(burete_kg, 3),
        "aliajAlV": round(aliaj_kg, 3),
        "alMetal": round(al_metal_kg, 3),
        "tio2": round(tio2_kg, 3),
        "feMetal": round(fe_metal_kg, 3),
    }
    return rezultat, None
