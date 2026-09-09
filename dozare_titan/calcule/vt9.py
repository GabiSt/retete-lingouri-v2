"""Formula de dozare pentru Ti-VT9.

Reproduce EXACT calculul din Excel (CDA 26854, foaia "RetDozare"), pe un
lant de formule secventiale — la fel ca la Ti6Al4V (vezi calcule/ti6al4v.py),
NU prin rezolvarea unui sistem cu solutie unica.

Elemente de aliere pentru VT9 (OST 1 90013-81): Al, Mo, Zr, Fe, Si, O.
Nu exista Vanadiu la VT9 — cheia "v" din target nu se foloseste aici.

Materiale implicate (rand 8-15 din "RetDozare"):
  - Ti burete      -> id "burete"
  - Aliaj Al-Mo    -> id "aliajAlMo"   (echivalentul Aliajului Al-V la Ti6Al4V)
  - Al metal       -> id "alMetal"
  - Zr 702         -> id "zrMetal"
  - Fe electrolitic-> id "feMetal"
  - Prealiaj SiTi  -> id "aliajSiTi"
  - TiO2           -> id "tio2"

Ordinea de calcul (identica cu Excel, randul 22 din "RetDozare"):
  1. Aliaj Al-Mo se calculeaza din balanta de Mo (singurul material cu Mo).
  2. Al metal se calculeaza din balanta de Al: (tinta Al) / (%Al Al-metal)
     minus contributia de Al (in kg) adusa de Aliajul Al-Mo — EXACT ca in
     Excel (=B22*E19/E11-D22*E10): scaderea NU se mai imparte la %Al din
     Al metal, e o particularitate a formulei originale si se pastreaza.
  3. Zr metal se calculeaza direct din balanta de Zr (independent).
  4. Fe electrolitic se calculeaza direct din balanta de Fe (independent,
     fara sa se scada Fe-ul adus de celelalte materiale — ca in Excel).
  5. Prealiaj SiTi se calculeaza direct din balanta de Si (independent).
  6. Ti burete = portia de baza - Aliaj Al-Mo - Al metal - Zr metal
     - Fe electrolitic - SiTi (balanta de masa a incarcaturii principale).
  7. TiO2 se adauga suplimentar pentru a acoperi balanta de O: se imparte
     tinta de O la %O din TiO2, apoi se scade O adus (in kg) de TOATE
     celelalte materiale (burete, Al-Mo, Al metal, Zr metal, Fe
     electrolitic, SiTi) — EXACT ca in Excel (=B22*I19/G15-C22*G9-D22*G10
     -E22*G11-F22*G12-G22*G13-H22*G14): doar primul termen se imparte la
     %O din TiO2, scaderile sunt mase brute, nu se mai imparte si termenul
     scazut.

TiO2 este o doza mica, adaugata PESTE incarcatura principala (ca in
Excel), deci masa finala a barei poate fi usor peste "portie".
"""

from ..utils import to_float


def calculeaza_dozare_kg(target, comps, p):
    """Aplica formulele Excel (CDA 26854 / RetDozare) pentru o portie de p kg.

    Parametri:
      target -- compozitia chimica tinta a retetei (dict cu chei "al",
                "mo", "zr", "fe", "si", "o", in procente). Cheia "v" nu
                se foloseste la VT9.
      comps  -- compozitia (%) fiecarui lot selectat, indexata dupa
                material. Chei asteptate:
                  comps["burete"]["Ti"|"O"|"Fe"]
                  comps["aliajAlMo"]["Mo"|"Al"|"O"|"Fe"]
                  comps["alMetal"]["Al"|"O"|"Fe"]
                  comps["zrMetal"]["Zr"|"O"]
                  comps["feMetal"]["Fe"|"O"]
                  comps["aliajSiTi"]["Ti"|"O"|"Fe"|"Si"]
                  comps["tio2"]["O"]
                Elementele care nu apar la un anumit lot (ex. O la Al
                metal / Fe electrolitic) trebuie sa vina din comun.py ca
                0, la fel ca in Excel unde celula respectiva e goala.
      p      -- portia, in kg

    Returneaza (rezultat, eroare):
      rezultat -- dict {material_id: kg} pentru O SINGURA portie, sau None
      eroare   -- mesaj de eroare (string), sau None daca s-a calculat OK
    """
    mo_tinta_kg = to_float(target.get("mo")) / 100 * p
    al_tinta_kg = to_float(target.get("al")) / 100 * p
    zr_tinta_kg = to_float(target.get("zr")) / 100 * p
    fe_tinta_kg = to_float(target.get("fe")) / 100 * p
    si_tinta_kg = to_float(target.get("si")) / 100 * p
    o_tinta_kg = to_float(target.get("o")) / 100 * p

    # 1. Aliaj Al-Mo din balanta de Mo (ignorat daca tinta Mo = 0%)
    if abs(mo_tinta_kg) < 1e-9:
        aliaj_kg = 0.0
    else:
        mo_almo = comps["aliajAlMo"]["Mo"]
        if abs(mo_almo) < 1e-9:
            return None, "Lotul de Aliaj Al-Mo nu are %Mo definit — nu se poate calcula."
        aliaj_kg = mo_tinta_kg / (mo_almo / 100)

    # 2. Al metal din balanta de Al (ignorat daca tinta Al = 0%). Se pastreaza
    #    EXACT forma din Excel: =B22*E19/E11-D22*E10 — scaderea contributiei
    #    Aliajului Al-Mo NU se mai imparte la %Al din Al metal.
    if abs(al_tinta_kg) < 1e-9:
        al_metal_kg = 0.0
    else:
        al_am = comps["alMetal"]["Al"]
        if abs(al_am) < 1e-9:
            return None, "Lotul de Al metal nu are %Al definit — nu se poate calcula."
        al_din_alv = aliaj_kg * comps["aliajAlMo"]["Al"] / 100
        al_metal_kg = al_tinta_kg / (al_am / 100) - al_din_alv

    # 3. Zr metal din balanta de Zr (independent, ignorat daca tinta Zr = 0%)
    if abs(zr_tinta_kg) < 1e-9:
        zr_metal_kg = 0.0
    else:
        zr_zrmet = comps["zrMetal"]["Zr"]
        if abs(zr_zrmet) < 1e-9:
            return None, "Lotul de Zr metal nu are %Zr definit — nu se poate calcula."
        zr_metal_kg = zr_tinta_kg / (zr_zrmet / 100)

    # 4. Fe electrolitic din balanta de Fe (independent, ignorat daca tinta Fe = 0%)
    if abs(fe_tinta_kg) < 1e-9:
        fe_metal_kg = 0.0
    else:
        fe_fe = comps["feMetal"]["Fe"]
        if abs(fe_fe) < 1e-9:
            return None, "Lotul de Fe electrolitic nu are %Fe definit — nu se poate calcula."
        fe_metal_kg = fe_tinta_kg / (fe_fe / 100)

    # 5. Prealiaj SiTi din balanta de Si (independent, ignorat daca tinta Si = 0%)
    if abs(si_tinta_kg) < 1e-9:
        si_ti_kg = 0.0
    else:
        si_siti = comps["aliajSiTi"]["Si"]
        if abs(si_siti) < 1e-9:
            return None, "Lotul de Prealiaj SiTi nu are %Si definit — nu se poate calcula."
        si_ti_kg = si_tinta_kg / (si_siti / 100)

    # 6. Burete Ti = restul incarcaturii principale
    burete_kg = p - aliaj_kg - al_metal_kg - zr_metal_kg - fe_metal_kg - si_ti_kg

    # 7. TiO2 adaugat suplimentar, din balanta de O (ignorat daca tinta O = 0%)
    o_din_burete = burete_kg * comps["burete"]["O"] / 100
    o_din_almo = aliaj_kg * comps["aliajAlMo"]["O"] / 100
    o_din_al = al_metal_kg * comps["alMetal"].get("O", 0) / 100
    o_din_zr = zr_metal_kg * comps["zrMetal"]["O"] / 100
    o_din_fe = fe_metal_kg * comps["feMetal"].get("O", 0) / 100
    o_din_siti = si_ti_kg * comps["aliajSiTi"]["O"] / 100
    if abs(o_tinta_kg) < 1e-9:
        tio2_kg = 0.0
    else:
        o_tio2 = comps["tio2"]["O"]
        if abs(o_tio2) < 1e-9:
            return None, "Lotul de TiO2 nu are %O definit — nu se poate calcula."
        # Ca la Excel: doar tinta se imparte la %O din TiO2, deducerile sunt
        # mase brute (nu se mai imparte inca o data termenul scazut).
        tio2_kg = (
            o_tinta_kg / (o_tio2 / 100)
            - o_din_burete
            - o_din_almo
            - o_din_al
            - o_din_zr
            - o_din_fe
            - o_din_siti
        )

    rezultat = {
        "burete": burete_kg,
        "aliajAlMo": aliaj_kg,
        "alMetal": al_metal_kg,
        "zrMetal": zr_metal_kg,
        "feMetal": fe_metal_kg,
        "aliajSiTi": si_ti_kg,
        "tio2": tio2_kg,
    }
    return rezultat, None
