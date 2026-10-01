"""Reteta de dozare Ti 6-2-4-2 (Ti-6Al-2Sn-4Zr-2Mo-0.08Si).

Sursa: Excel-ul de productie "Cda 24883-Ti6242-1VAR440-LI1-LI2", foaia
"RetDozare", blocul "Dozare in bara [Kg]" (randurile 20-25 pentru Reteta 1,
57-62 pentru Reteta 2).

Materiale (ids din config.MATERIALE) si ce element aduce fiecare:

    burete       Ti de baza (restul); O si Fe din buletin
    aliajAlMo    Mo (elementul principal) + Al
    alMetal      Al
    zrMetal      Zr (Zr 702, pur — fara Sn)          } alternative: se
    zy4          Zr (elementul principal) + Sn       } foloseste UN SINGUR
                                                       } lot, oricare e pe
                                                       } stoc (vezi mai jos)
    snMetal      Sn
    aliajSiTi    Si
    tio2         O (ultimul: completeaza oxigenul pana la tinta)

Zirconiul poate veni fie din Zr metal (Zr 702, zirconiu comercial-pur),
fie din Zy4 (Zircaloy 4, un aliaj Zr-Sn) — depinde ce lot are uzina pe
stoc la un moment dat (vezi si "Cda 26830-Ti6242", care foloseste Zr 702
in loc de Zy4 fata de "Cda 24883"). Functia alege AUTOMAT sursa: cea
dintre ele care are compozitie de Zr introdusa (lot selectat) in
``comps``. Doar Zy4 aduce si Sn; Zr 702 nu.

Ordinea calculului (aceeasi ca in Excel): Mo -> Al -> Zr -> Sn -> Si -> O -> Ti.
Fiecare material acopera elementul lui MINUS ce au adus deja materialele
calculate inainte de el (Al din AlMo, Sn din Zy4, O din toate).

Doua moduri de calcul (parametrul ``mod``):

  "complet" (implicit, ca restul aplicatiei) — balanta de masa completa:
      kg_Al = (Al_tinta*B - Al_adus) / puritate_Al
      adica se scade INTAI ce e deja adus si se imparte DUPA la puritatea
      materialului, ca sa iasa exact tinta;

  "excel" — reproduce EXACT formulele din foaia RetDozare, inclusiv
      particularitatea de precedenta a operatorilor:
      kg_Al = Al_tinta*B / puritate_Al - Al_adus
      (scaderea ramane neimpartita la puritate). Pentru Al, Sn si mai ales
      TiO2 (puritate 40%) rezultatul iese usor diferit; in plus, Excel-ul nu
      scade TiO2 din cei B kg (Ti = B - Mo - Al - Zr - Sn - Si).

Modulul nu depinde de restul aplicatiei (doar de utils.to_float, cand exista).
"""

# Materialele acestei retete, in ordinea afisarii.
MATERIALE_6242 = ("burete", "aliajAlMo", "alMetal", "zrMetal", "zy4",
                  "snMetal", "aliajSiTi", "tio2")

# Cele doua surse ALTERNATIVE de zirconiu — se foloseste UNA SINGURA,
# oricare are lot (compozitie de Zr) selectat; zy4 e preferata daca,
# dintr-un motiv sau altul, ambele ar avea compozitie (nu ar trebui sa se
# intample in practica).
MATERIALE_ZR_ALTERNATIVE = ("zy4", "zrMetal")

# Materialele "secundare" (nu burete/TiO2), folosite la insumarea a ce s-a
# adus deja dintr-un element inainte de a doza urmatorul. Contine AMBELE
# surse de Zr — cea nefolosita ramane mereu cu kg 0, deci nu strica suma.
MATERIALE_SECUNDARE = ("aliajAlMo", "alMetal", "zy4", "zrMetal", "snMetal", "aliajSiTi")

# Elementul principal al fiecarui material (cel dupa care se doza).
ELEMENT_PRINCIPAL = {
    "aliajAlMo": "Mo", "alMetal": "Al", "zy4": "Zr", "zrMetal": "Zr",
    "snMetal": "Sn", "aliajSiTi": "Si", "tio2": "O",
}

# Elementele dozate, in ordinea calculului.
ELEMENTE_DOZATE = ("mo", "al", "zr", "sn", "si", "o")

# Modul folosit de aplicatie. "complet" = balanta de masa corecta; "excel" =
# reproduce exact foaia de calcul (cu particularitatea ei de precedenta).
MOD_IMPLICIT = "complet"

BAZA_KG = 10.0          # Excel-ul calculeaza la 10 kg si apoi scaleaza


def _to_float(x, implicit=0.0):
    try:
        if x is None or x == "":
            return implicit
        return float(str(x).replace(",", "."))
    except (TypeError, ValueError):
        return implicit


def _fractii(comps):
    """comps -> {material: {Element: fractie 0..1}}.

    Loturile pastreaza compozitia in PROCENTE (dozaAl = 33.8 inseamna 33.8%).
    Ca sa nu depinda de scara in care vine ``comps``, se verifica elementele
    principale ale materialelor (Mo in AlMo, Al in Al metal, Zr in Zy4 ...:
    intotdeauna peste 30%): daca oricare are o valoare peste 1.5, totul e in
    procente; daca sunt toate sub 1.5, totul e in fractii. Aceeasi scara se
    aplica si buretelui (care nu are element principal).
    """
    norm = {mat: {str(k).removeprefix("doza").capitalize(): _to_float(v)
                  for k, v in (elemente or {}).items()}
            for mat, elemente in (comps or {}).items()}
    principale = [norm[m].get(ELEMENT_PRINCIPAL[m], 0.0)
                  for m in norm if m in ELEMENT_PRINCIPAL]
    principale = [v for v in principale if v > 0]
    procente = (max(principale) > 1.5) if principale else True
    scara = 100.0 if procente else 1.0
    return {mat: {k: v / scara for k, v in el.items()} for mat, el in norm.items()}


def _tinta_fractii(target):
    """target (in procente, cheile al/mo/sn/zr/o/si) -> fractii."""
    return {e: _to_float((target or {}).get(e)) / 100.0 for e in ELEMENTE_DOZATE}


def doza_la_10kg(tinta, f, mod="complet"):
    """Dozarea pentru BAZA_KG (10 kg), in kg pe material.

    tinta -- {"mo","al","zr","sn","si","o": fractie}   (0.068 = 6.8%)
    f     -- {material: {Element: fractie}}
    Returneaza (kg, eroare) cu kg = {material_id: kg}.
    """
    def c(mat, el):
        return f.get(mat, {}).get(el, 0.0)

    B = BAZA_KG
    for mat in ("aliajAlMo", "alMetal", "snMetal", "aliajSiTi", "tio2"):
        principal = ELEMENT_PRINCIPAL[mat]
        cerut = tinta["o" if principal == "O" else principal.lower()]
        if cerut > 0 and c(mat, principal) <= 0:
            return None, (f"Lipseste compozitia lotului la materialul \u201e{mat}\u201d "
                          f"({principal} = 0) — nu pot doza {principal}.")

    # Sursa de Zr efectiv disponibila (are lot cu compozitie de Zr).
    zr_mat = next((m for m in MATERIALE_ZR_ALTERNATIVE if c(m, "Zr") > 0), None)
    if tinta["zr"] > 0 and zr_mat is None:
        return None, ("Lipseste compozitia de Zr a lotului selectat (zrMetal "
                      "sau zy4) — nu pot doza Zr.")
    if zr_mat is None:
        zr_mat = "zy4"          # tinta Zr = 0: nu se foloseste, ramane la 0 kg

    complet = (mod != "excel")
    kg = {"burete": 0.0, "zy4": 0.0, "zrMetal": 0.0}

    # Mo -> AlMo
    kg["aliajAlMo"] = B * tinta["mo"] / c("aliajAlMo", "Mo") if tinta["mo"] else 0.0
    # Al -> Al metal (minus Al adus de AlMo)
    al_adus = kg["aliajAlMo"] * c("aliajAlMo", "Al")
    if tinta["al"]:
        if complet:
            kg["alMetal"] = (B * tinta["al"] - al_adus) / c("alMetal", "Al")
        else:
            kg["alMetal"] = B * tinta["al"] / c("alMetal", "Al") - al_adus
    else:
        kg["alMetal"] = 0.0
    # Zr -> sursa aleasa (Zr metal sau Zy4)
    kg[zr_mat] = B * tinta["zr"] / c(zr_mat, "Zr") if tinta["zr"] else 0.0
    # Sn -> Sn metal (minus Sn adus de Zy4 — Zr metal nu aduce Sn)
    sn_adus = kg[zr_mat] * c(zr_mat, "Sn")
    if tinta["sn"]:
        if complet:
            kg["snMetal"] = (B * tinta["sn"] - sn_adus) / c("snMetal", "Sn")
        else:
            kg["snMetal"] = B * tinta["sn"] / c("snMetal", "Sn") - sn_adus
    else:
        kg["snMetal"] = 0.0
    # Si -> prealiaj SiTi
    kg["aliajSiTi"] = B * tinta["si"] / c("aliajSiTi", "Si") if tinta["si"] else 0.0

    # Titan = restul (in Excel TiO2 NU se scade din B; in modul complet da)
    altele = sum(kg[m] for m in MATERIALE_SECUNDARE)

    # O -> TiO2, dupa ce s-a adunat oxigenul adus de toate celelalte
    if complet:
        # Ti si TiO2 depind unul de altul (O adus de Ti): se rezolva exact.
        #   Ti = B - altele - TiO2
        #   O_tinta*B = O_alte + Ti*O_ti + TiO2*O_tio2
        o_alte = sum(kg[m] * c(m, "O") for m in MATERIALE_SECUNDARE)
        numitor = c("tio2", "O") - c("burete", "O")
        if numitor <= 0:
            return None, "Oxigenul din TiO2 trebuie sa fie mai mare decat cel din burete."
        kg["tio2"] = (B * tinta["o"] - o_alte - (B - altele) * c("burete", "O")) / numitor
        kg["burete"] = B - altele - kg["tio2"]
    else:
        kg["burete"] = B - altele
        o_alte = sum(kg[m] * c(m, "O") for m in
                     ("burete",) + MATERIALE_SECUNDARE)
        kg["tio2"] = B * tinta["o"] / c("tio2", "O") - o_alte if tinta["o"] else 0.0

    # Daca materiile prime aduc DEJA mai mult oxigen decat tinta, TiO2 ar iesi
    # negativ: nu se adauga TiO2 (0 kg) si oxigenul din bara iese putin peste
    # tinta — se vede in compozitia rezultata si e verificat fata de standard.
    if kg["tio2"] < 0:
        kg["tio2"] = 0.0
        if complet:
            kg["burete"] = B - altele
    return kg, None


def compozitie_din_kg(kg, f):
    """Compozitia (in %) a incarcaturii, pornind de la kg pe material."""
    total = sum(kg.values())
    if total <= 0:
        return {}
    rezultat = {}
    for el in ("Al", "Mo", "Sn", "Zr", "Si", "O", "Fe"):
        rezultat[el] = 100.0 * sum(kg.get(m, 0.0) * f.get(m, {}).get(el, 0.0)
                                   for m in kg) / total
    return rezultat


def calculeaza_dozare_kg(target, comps, p, mod=None):
    """Semnatura comuna tuturor formulelor din calcule.DOZATOARE:
    (target, comps, p) -> (rezultat, eroare).

    target -- compozitia tinta in bara, in % (cheile al, mo, sn, zr, o, si)
    comps  -- compozitia (%) a lotului ales, pe material
    p      -- portia, in kg
    rezultat -- {material_id: kg} pentru O SINGURA portie
    """
    try:
        p = float(p)
        f = _fractii(comps)
        kg10, eroare = doza_la_10kg(_tinta_fractii(target), f, mod or MOD_IMPLICIT)
        if eroare:
            return None, eroare
        return {m: v * p / BAZA_KG for m, v in kg10.items()}, None
    except Exception as e:      # nu lasam o formula sa doboare interfata
        return None, f"Eroare la calculul Ti 6-2-4-2: {type(e).__name__}: {e}"
