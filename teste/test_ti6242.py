"""Reteta Ti 6-2-4-2: comparatie cu Excel-ul de productie
"Cda 24883-Ti6242-1VAR440-LI1-LI2" (foaia RetDozare) si legatura cu standardele.

    python3 teste/test_ti6242.py

Valorile "asteptate" sunt copiate din celulele Excel-ului (referinta in
comentarii); Excel-ul e sursa de adevar pentru modul "excel".
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dozare_titan.calcule import ti6242, DOZATOARE                    # noqa: E402
from dozare_titan.config import ALIAJE_SPEC, MATERIALE                # noqa: E402
from dozare_titan import standarde as st                              # noqa: E402

erori = []
nr = [0]


def verifica(eticheta, obtinut, asteptat):
    nr[0] += 1
    if obtinut != asteptat:
        erori.append(f"{eticheta}: obtinut {obtinut!r}, asteptat {asteptat!r}")


def aprox(eticheta, obtinut, asteptat, tol=1e-9):
    nr[0] += 1
    if abs(obtinut - asteptat) > tol:
        erori.append(f"{eticheta}: obtinut {obtinut!r}, asteptat {asteptat!r}")


ORDINE = ("burete", "aliajAlMo", "alMetal", "zy4", "snMetal", "aliajSiTi", "tio2")

# Loturile din foaia RetDozare (compozitii in %), aceleasi la ambele retete.
COMPS = {
    "burete":    {"O": 0.03, "Fe": 0.005},                          # H10, I10
    "aliajAlMo": {"Mo": 65.1, "Al": 34.7, "O": 0.09, "Fe": 0.07},   # D11:I11
    "alMetal":   {"Al": 99.87},                                     # E12
    "zy4":       {"Zr": 97.8, "Sn": 1.557, "O": 0.22},              # F13:H13
    "snMetal":   {"Sn": 99.97, "O": 0.05},                          # G14, H14
    "aliajSiTi": {"Si": 61.5, "O": 0.02, "Fe": 0.04},               # J15, H15, I15
    "tio2":      {"O": 40.0},                                       # H16
}
STOC = dict(zip(ORDINE, (3250, 1065, 1072, 389, 60.7, 41.5, 1000)))   # K10:K16

# Tinta = "Dozare in bara" (D6:N6 / D43:N43), in %.
TINTA_R1 = {"al": 6.8, "mo": 1.96, "sn": 2.1, "zr": 3.9, "o": 0.047, "si": 0.08}
TINTA_R2 = {"al": 7.0, "mo": 1.93, "sn": 2.226, "zr": 3.9, "o": 0.033, "si": 0.08}

# --- valorile din Excel --------------------------------------------------
R1_10KG = (8.50687743935004, 0.301075268817204, 0.576412032416335, 0.398773006134969,
           0.20385412320015, 0.0130081300813008, 0.00794513972514624)        # C23:I23
R1_PORTIE = (34.0275097574002, 1.20430107526882, 2.30564812966534, 1.59509202453988,
             0.815416492800601, 0.0520325203252032, 0.031780558900585)       # C24:I24
R1_BARA = (782.632724420204, 27.6989247311828, 53.0299069823028, 36.6871165644172,
           18.7545793344138, 1.19674796747967, 0.730952854713454)            # C25:I25
R1_REST = (2467.3672755798, 1037.30107526882, 1018.9700930177, 352.312883435583,
           41.9454206655862, 40.3032520325203, 999.269047145287)             # C31:I31
R2_10KG = (8.47725684096156, 0.296466973886329, 0.598037144601346, 0.398773006134969,
           0.216457904334491, 0.0130081300813008, 0.0044518714795334)        # C60:I60
R2_BARA = (779.907629368464, 27.2749615975422, 55.0194173033238, 36.6871165644172,
           19.9141271987731, 1.19674796747967, 0.409572176117072)            # C62:I62
R2_REST = (1687.45964621133, 1010.02611367128, 963.950675714373, 315.625766871166,
           22.0312934668131, 39.1065040650407, 998.85947496917)              # C68:I68


def calc(tinta, p, mod, comps=COMPS):
    rez, err = ti6242.calculeaza_dozare_kg(tinta, comps, p, mod)
    verifica(f"fara eroare ({mod}, p={p})", err, None)
    return rez


# ---------------------------------------------------------------------------
print("1. Modul 'excel' = celulele din Excel (Reteta 1)")
r = calc(TINTA_R1, 10, "excel")
for m, ex in zip(ORDINE, R1_10KG):
    aprox(f"R1 la 10 kg / {m}  (C23:I23)", r[m], ex)
r = calc(TINTA_R1, 40, "excel")
for m, ex in zip(ORDINE, R1_PORTIE):
    aprox(f"R1 portie 40 kg / {m}  (C24:I24)", r[m], ex)
bara = {m: r[m] * 23 for m in ORDINE}
for m, ex in zip(ORDINE, R1_BARA):
    aprox(f"R1 bara (23 presari) / {m}  (C25:I25)", bara[m], ex, 1e-8)
for m, ex in zip(ORDINE, R1_REST):
    aprox(f"R1 stoc ramas dupa bara / {m}  (C31:I31)", STOC[m] - bara[m], ex, 1e-8)

print("2. Modul 'excel' = celulele din Excel (Reteta 2, pe stocul ramas de Reteta 1)")
r = calc(TINTA_R2, 10, "excel")
for m, ex in zip(ORDINE, R2_10KG):
    aprox(f"R2 la 10 kg / {m}  (C60:I60)", r[m], ex)
bara2 = {m: calc(TINTA_R2, 40, "excel")[m] * 23 for m in ORDINE}
for m, ex in zip(ORDINE, R2_BARA):
    aprox(f"R2 bara / {m}  (C62:I62)", bara2[m], ex, 1e-8)
for m, ex, s1 in zip(ORDINE, R2_REST, R1_REST):
    aprox(f"R2 stoc ramas / {m}  (C68:I68)", s1 - bara2[m], ex, 1e-8)

# ---------------------------------------------------------------------------
print("3. Modul 'complet' (implicit): compozitia iese exact pe tinta")
verifica("modul implicit al aplicatiei", ti6242.MOD_IMPLICIT, "complet")
f = ti6242._fractii(COMPS)
r = calc(TINTA_R1, 10, "complet")
comp = ti6242.compozitie_din_kg(r, f)
for el, cheie in (("Al", "al"), ("Mo", "mo"), ("Sn", "sn"), ("Zr", "zr"), ("Si", "si"), ("O", "o")):
    aprox(f"R1 complet: {el} rezultat = tinta", comp[el], TINTA_R1[cheie], 1e-9)
aprox("R1 complet: masa totala = 10 kg", sum(r.values()), 10.0, 1e-12)

rex = calc(TINTA_R1, 10, "excel")
compex = ti6242.compozitie_din_kg(rex, f)
verifica("R1 excel: O iese PESTE tinta (particularitatea de precedenta)", compex["O"] > TINTA_R1["o"] * 1.4, True)
verifica("R1: TiO2 complet < TiO2 din Excel", r["tio2"] < rex["tio2"], True)
verifica("R1: Al metal complet < Al metal din Excel", r["alMetal"] < rex["alMetal"], True)
for m in ("aliajAlMo", "zy4", "aliajSiTi"):
    aprox(f"R1: {m} identic in ambele moduri", r[m], rex[m], 1e-12)

print("4. Reteta 2: O din materiile prime depaseste deja tinta -> TiO2 = 0")
r2 = calc(TINTA_R2, 10, "complet")
aprox("R2 complet: TiO2 = 0", r2["tio2"], 0.0, 1e-15)
comp2 = ti6242.compozitie_din_kg(r2, f)
verifica("R2 complet: O peste tinta 0.033 dar sub maximul AMS (0.15)", 0.033 < comp2["O"] < 0.15, True)
aprox("R2 complet: masa totala = 10 kg", sum(r2.values()), 10.0, 1e-12)
for el, cheie in (("Al", "al"), ("Mo", "mo"), ("Sn", "sn"), ("Zr", "zr"), ("Si", "si")):
    aprox(f"R2 complet: {el} rezultat = tinta", comp2[el], TINTA_R2[cheie], 1e-9)

# ---------------------------------------------------------------------------
print("5. Formatul lui comps: procente, fractii sau chei 'dozaX'")
ref = calc(TINTA_R1, 40, "complet")
in_fractii = {m: {e: v / 100 for e, v in el.items()} for m, el in COMPS.items()}
in_doza = {m: {"doza" + e: v for e, v in el.items()} for m, el in COMPS.items()}
for nume, comps in (("fractii", in_fractii), ("chei dozaX", in_doza)):
    alt = calc(TINTA_R1, 40, "complet", comps)
    for m in ORDINE:
        aprox(f"{nume} / {m}", alt[m], ref[m], 1e-9)

print("6. Erori clare")
rez, err = ti6242.calculeaza_dozare_kg(TINTA_R1, {**COMPS, "zy4": {}}, 40)
verifica("lot Zy4 fara Zr -> mesaj", (rez, "zy4" in (err or "")), (None, True))
rez, err = ti6242.calculeaza_dozare_kg(TINTA_R1, {}, 40)
verifica("fara loturi -> eroare, nu exceptie", rez is None and bool(err), True)

# ---------------------------------------------------------------------------
print("7. Inregistrare in aplicatie si legatura cu standardele")
verifica("aliaj real (in ALIAJE_SPEC)", "Ti 6-2-4-2" in ALIAJE_SPEC, True)
verifica("formula in DOZATOARE", DOZATOARE.get("Ti 6-2-4-2") is ti6242.calculeaza_dozare_kg, True)
verifica("nu mai e placeholder", st.este_placeholder("Ti 6-2-4-2"), False)
ids = {m["id"] for m in MATERIALE}
verifica("materialele retetei exista in config.MATERIALE", set(ORDINE) <= ids, True)

cheie_excel = "Ti 6-2-4-2"
nume_std = [s["nume"] for s in st.standarde_pentru(cheie_excel)]
verifica("standardele 6242 din Excel (AMS 4976L, AMS 4975P)", nume_std, ["AMS 4976L", "AMS 4975P"])
verifica("aliajul apare o singura data in lista (nu si ca placeholder)",
         [i["cheie"] for i in st.aliaje_disponibile()].count(cheie_excel), 1)

# Limitele din config (foaia RetDozare, randul "STANDARD AMS 4976") = cele din Excel-ul de standarde
ams4976 = st.standarde_pentru(cheie_excel)[0]
spec_comanda = st.spec_efectiv({"tipAliaj": cheie_excel, "standard": "AMS 4976L"})
for el, sfx in (("Al", "min"), ("Al", "max"), ("Mo", "min"), ("Mo", "max"), ("Sn", "min"), ("Sn", "max"),
                ("Zr", "min"), ("Zr", "max"), ("Si", "min"), ("Si", "max"), ("O", "max"), ("Fe", "max")):
    k = f"{el.lower()}_{sfx}"
    aprox(f"limita {k} din config = standard AMS 4976L", ALIAJE_SPEC[cheie_excel][k], spec_comanda[k], 1e-12)

print("\n" + "=" * 70)
if erori:
    print(f"ESUAT — {len(erori)} din {nr[0]} verificari:\n")
    for e in erori:
        print("  \u2717 " + e)
    sys.exit(1)
print(f"OK — toate cele {nr[0]} verificari trec.")
