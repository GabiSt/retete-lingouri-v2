"""Verificare automata a logicii de distributie a barelor, pe un caz REAL
de productie: Cda 26858-Ti5-2VAR-d600-L1-L4 (12 bare de 900 kg).

Ruleaza fara interfata grafica:

    python3 teste/test_cda_26858.py

Datele de intrare (loturi, compozitii, cantitati) si valorile asteptate
sunt luate direct din Excel-ul de productie
"Cda 26858-Ti5-2VAR-d600-L1-L4-draft.xls", tab-urile "DateIntrare" si
"Retete" (tabelele "Bilant presare [Kg]").

Ce verifica:
  1. dozarea pe portie de 40 kg, pentru fiecare dintre cele 3 retete;
  2. distributia celor 12 bare pe retete:
        Reteta 1: barele 1-4      (se termina prealiajul Al-V la bara 5)
        Reteta 2: bara 5 (report) + barele 6-10  (se termina buretele)
        Reteta 3: bara 11 (report) + bara 12
  3. consumul si restul fiecarui lot dupa fiecare reteta;
  4. retragerea consumului (trebuie sa readuca stocul exact de unde a
     plecat, pentru fiecare reteta).
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dozare_titan.calcule import calculeaza_bara, lot_rest          # noqa: E402
from dozare_titan.calcule import planificare                        # noqa: E402
from dozare_titan.config import ORDINE_MAT                          # noqa: E402

TIP_ALIAJ = "Ti5"
TOL = 0.12          # kg — toleranta fata de Excel (vezi nota de la final)
CORECTIE = []       # se completeaza pe parcurs (vezi verifica_rest)
TOL_DOZA = 0.005    # kg pe portia de 40 kg

erori = []
verificari = [0]


def verifica(eticheta, obtinut, asteptat, tol=TOL):
    verificari[0] += 1
    if abs(obtinut - asteptat) > tol:
        erori.append(f"{eticheta}: obtinut {obtinut:.3f}, asteptat {asteptat:.3f} "
                     f"(diferenta {obtinut - asteptat:+.3f})")
        return False
    return True


def verifica_rest(eticheta, obtinut, excel, corectii=()):
    """Compara un rest de lot cu Excel-ul, corectat cu diferenta cunoscuta
    de dozare a barelor de report.

    In Excel, bara mutata pe reteta urmatoare e scazuta acolo cu dozarea
    RETETEI NOI (asa e construit sablonul: coloana scade mereu dozarea
    foii curente). Aplicatia o scade cu DOZAREA VECHE, cea cu care s-a
    presat efectiv bara — regula reala din sectie. Diferenta, pe material,
    e exact (doza noua - doza veche) pentru fiecare bara de report, si se
    aduna de la reteta la reteta cat timp lotul ramane acelasi.

    corectii = lista de (calc_reteta_noua, calc_reteta_veche, material).
    """
    corectie = sum(cn["rezultatTotal"].get(m, 0.0) - cv["rezultatTotal"].get(m, 0.0)
                   for cn, cv, m in corectii)
    return verifica(eticheta, obtinut, excel + corectie)


def verifica_egal(eticheta, obtinut, asteptat):
    verificari[0] += 1
    if obtinut != asteptat:
        erori.append(f"{eticheta}: obtinut {obtinut!r}, asteptat {asteptat!r}")
        return False
    return True


# ---------------------------------------------------------------------------
# 1. Loturile, exact ca in tab-ul "DateIntrare"
# ---------------------------------------------------------------------------
def lot(id_, material, nume, stoc, **doze):
    l = {"id": id_, "material": material, "lot": nume, "furnizor": "",
         "stocIntrare": stoc, "consum": 0.0}
    l.update({k: v for k, v in doze.items()})
    return l


LOTS = [
    # Burete Ti 260215-439: O 0.027%, Fe 0.008%
    lot("burete1", "burete", "260215-439", 8478, dozaO=0.027, dozaFe=0.008),
    # Burete Ti 260409-945: O 0.023%, Fe 0.003%
    lot("burete2", "burete", "260409-945", 8250, dozaO=0.023, dozaFe=0.003),
    # Aliaj AlV 360744: V 65.3%, Al 34.3%, O 0.05%, Fe 0.2%
    lot("alv1", "aliajAlV", "360744", 273, dozaV=65.3, dozaAl=34.3, dozaO=0.05, dozaFe=0.2),
    # Aliaj AlV 361748: V 65.7%, Al 33.8%, O 0.1%, Fe 0.2%
    lot("alv2", "aliajAlV", "361748", 1119, dozaV=65.7, dozaAl=33.8, dozaO=0.1, dozaFe=0.2),
    # Al metal W660078186: Al 99.87%, Fe 0.07%
    lot("al1", "alMetal", "W660078186", 874, dozaAl=99.87, dozaFe=0.07),
    # Fe metal: Fe 99%
    lot("fe1", "feMetal", "lot nou", 923, dozaFe=99.0),
    # TiO2 KRONOS: O 40%
    lot("tio1", "tio2", "LOT 0087718", 607, dozaO=40.0),
]
LOT_DUPA_ID = {l["id"]: l for l in LOTS}

TARGET = {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.17}

ORDER = {
    "id": "cda26858",
    "nume": "Cda 26858-Ti5-2VAR-d600-L1-L4",
    "tipAliaj": TIP_ALIAJ,
    "nrBare": "12",
    "retete": [{
        "id": "r1", "nume": "Reteta 1", "target": dict(TARGET),
        "portie": "40", "numarPresari": "22.5", "bare": [],
        "lotSel": {"burete": "burete1", "aliajAlV": "alv1", "alMetal": "al1",
                   "feMetal": "fe1", "tio2": "tio1"},
    }],
}


def plan():
    return planificare.planifica(ORDER, LOTS, calculeaza_bara, TIP_ALIAJ)


def rest(id_):
    return lot_rest(LOT_DUPA_ID[id_])


def stare_stoc():
    return {l["id"]: l["consum"] for l in LOTS}


def print_plan(p):
    for intrare in p:
        r = intrare["reteta"]
        bucati = []
        if intrare["report"]:
            bucati.append(f"1 report (dozare de pe {intrare['report']['retetaSursa']})")
        if intrare["bareNoi"]:
            bucati.append(f"{intrare['bareNoi']} bare dozate aici")
        if intrare["bareAsteptare"]:
            bucati.append(f"{intrare['bareAsteptare']} in asteptare")
        if intrare["aplicate"]:
            bucati.append(f"{intrare['aplicate']} cu consum aplicat")
        lim = ""
        if intrare["desf"] and intrare["desf"]["baraReport"]:
            lim = ("  -> cedeaza o bara retetei urmatoare, se termina: " +
                   ", ".join(intrare["desf"]["materialeLimitante"]))
        print(f"   {r['nume']}: {', '.join(bucati) or 'nicio bara'}{lim}")


print(__doc__.splitlines()[1])
print("=" * 78)

# ---------------------------------------------------------------------------
# RETETA 1
# ---------------------------------------------------------------------------
print("\nRETETA 1 — loturi: burete 260215-439, AlV 360744, Al W660078186, Fe, TiO2")

calc1 = calculeaza_bara(TIP_ALIAJ, TARGET, planificare.loturi_selectate(LOTS, ORDER["retete"][0]),
                        "40", "22.5")
assert not calc1.get("eroare"), calc1.get("eroare")

# Excel "Retete", randul PORTIE 40 al retetei 1 (B27:F27)
verifica("R1 dozare portie burete", calc1["rezultat"]["burete"], 35.8479039596805, TOL_DOZA)
verifica("R1 dozare portie aliajAlV", calc1["rezultat"]["aliajAlV"], 2.48085758039816, TOL_DOZA)
verifica("R1 dozare portie alMetal", calc1["rezultat"]["alMetal"], 1.67123845992133, TOL_DOZA)
verifica("R1 dozare portie tio2", calc1["rezultat"]["tio2"], 0.147701592851718, TOL_DOZA)
verifica("R1 dozare portie feMetal", calc1["rezultat"]["feMetal"], 0.0595965511116003, TOL_DOZA)
print(f"   dozare/bara (900 kg): burete {calc1['rezultatTotal']['burete']:.3f}  "
      f"AlV {calc1['rezultatTotal']['aliajAlV']:.3f}  Al {calc1['rezultatTotal']['alMetal']:.3f}  "
      f"TiO2 {calc1['rezultatTotal']['tio2']:.3f}  Fe {calc1['rezultatTotal']['feMetal']:.3f}")

p = plan()
planificare.aplica_plan(ORDER, p)
print_plan(p)

verifica_egal("R1 bare dozate aici", p[0]["bareNoi"], 4)
verifica_egal("R1 material care se termina", p[0]["desf"]["materialeLimitante"], ["aliajAlV"])
verifica_egal("R1 bara cedata mai departe", p[0]["desf"]["baraReport"]["index"], 5)

stoc_inainte_r1 = stare_stoc()
ok, mesaj = planificare.aplica_consum_reteta(ORDER, ORDER["retete"][0], LOTS,
                                             calculeaza_bara, TIP_ALIAJ)
verifica_egal("R1 aplicare consum", ok, True)

# Excel "Retete" R2 (D39:D43) = restul loturilor la intrarea in reteta 2
print(f"   dupa R1 -> rest burete {rest('burete1'):.3f} | AlV {rest('alv1'):.3f} | "
      f"Al {rest('al1'):.3f} | TiO2 {rest('tio1'):.3f} | Fe {rest('fe1'):.3f}")
verifica("R1 rest burete 260215-439", rest("burete1"), 5251.68864362875)
verifica("R1 rest AlV 360744 (golit de bara 5)", rest("alv1"), 0.0, 0.02)
verifica("R1 rest Al metal", rest("al1"), 723.58853860708)
verifica("R1 rest TiO2", rest("tio1"), 593.706856643345)
verifica("R1 rest Fe metal", rest("fe1"), 917.636310399956)

# ---------------------------------------------------------------------------
# RETETA 2 — se introduce lotul nou de AlV
# ---------------------------------------------------------------------------
print("\nRETETA 2 — lot nou de AlV (361748); bara 5 vine de report, cu dozarea retetei 1")
r2 = ORDER["retete"][1]
r2["lotSel"] = {"burete": "burete1", "aliajAlV": "alv2", "alMetal": "al1",
                "feMetal": "fe1", "tio2": "tio1"}

calc2 = calculeaza_bara(TIP_ALIAJ, TARGET, planificare.loturi_selectate(LOTS, r2), "40", "22.5")
# Excel "Retete", randul PORTIE 40 al retetei 2 (B57:F57)
verifica("R2 dozare portie burete", calc2["rezultat"]["burete"], 35.8454758309089, TOL_DOZA)
verifica("R2 dozare portie aliajAlV", calc2["rezultat"]["aliajAlV"], 2.46575342465753, TOL_DOZA)
verifica("R2 dozare portie alMetal", calc2["rezultat"]["alMetal"], 1.68877074443352, TOL_DOZA)
verifica("R2 dozare portie tio2", calc2["rezultat"]["tio2"], 0.144639920252493, TOL_DOZA)
verifica("R2 dozare portie feMetal", calc2["rezultat"]["feMetal"], 0.0596148642051603, TOL_DOZA)

p = plan()
planificare.aplica_plan(ORDER, p)
print_plan(p)

verifica_egal("R2 are bara de report", bool(p[1]["report"]), True)
verifica_egal("R2 bare dozate aici", p[1]["bareNoi"], 5)
verifica_egal("R2 material care se termina", p[1]["desf"]["materialeLimitante"], ["burete"])
verifica_egal("R2 numar total de bare", len(r2["bare"]), 6)

# AlV: bara de report mai ia din lotul NOU doar diferenta (Excel: 1168.72 - 55.48)
report_alv = p[1]["report"]["ramas"].get("aliajAlV", 0.0)
verifica("R2 bara de report — AlV luat din lotul nou", report_alv,
         55.8192955589587 - 49.7228177641654, 0.35)
# ...dar buretele si restul materialelor se consuma INTEGRAL de aici
verifica("R2 bara de report — burete luat de aici", p[1]["report"]["ramas"].get("burete", 0.0),
         806.577839092812)
verifica("R2 bara de report — Al metal luat de aici", p[1]["report"]["ramas"].get("alMetal", 0.0),
         37.6028653482299)

ok, mesaj = planificare.aplica_consum_reteta(ORDER, r2, LOTS, calculeaza_bara, TIP_ALIAJ)
verifica_egal("R2 aplicare consum", ok, True)

# Excel "Retete" R3 (D70:D74) = restul loturilor la intrarea in reteta 3
print(f"   dupa R2 -> rest burete {rest('burete1'):.3f} | AlV {rest('alv2'):.3f} | "
      f"Al {rest('al1'):.3f} | TiO2 {rest('tio1'):.3f} | Fe {rest('fe1'):.3f}")
# Buretele vechi e golit complet: restul lui (412.55 kg, Excel I47) pleaca
# in bara 11, care il duce pe reteta 3 — de-aia restul lotului e 0 aici.
verifica("R2 rest burete 260215-439 (golit de bara 11)", rest("burete1"), 0.0, 0.02)
verifica("R2 bara 11 ia din lotul vechi de burete",
         p[1]["reportOut"]["dinLotAnterior"].get("burete", 0.0), 412.549406456046)
verifica_rest("R2 rest AlV 361748", rest("alv2"), 835.846105435398,
              [(calc2, calc1, "aliajAlV")])
verifica_rest("R2 rest Al metal", rest("al1"), 495.604488108556,
              [(calc2, calc1, "alMetal")])
verifica_rest("R2 rest TiO2", rest("tio1"), 574.180467409259,
              [(calc2, calc1, "tio2")])
verifica_rest("R2 rest Fe metal", rest("fe1"), 909.588303732259,
              [(calc2, calc1, "feMetal")])

# ---------------------------------------------------------------------------
# RETETA 3 — se introduce lotul nou de burete
# ---------------------------------------------------------------------------
print("\nRETETA 3 — lot nou de burete (260409-945); bara 11 vine de report, cu dozarea retetei 2")
r3 = ORDER["retete"][2]
r3["lotSel"] = {"burete": "burete2", "aliajAlV": "alv2", "alMetal": "al1",
                "feMetal": "fe1", "tio2": "tio1"}

calc3 = calculeaza_bara(TIP_ALIAJ, TARGET, planificare.loturi_selectate(LOTS, r3), "40", "22.5")
# Excel "Retete", randul PORTIE 40 al retetei 3 (B88:F88)
verifica("R3 dozare portie burete", calc3["rezultat"]["burete"], 35.8453067170751, TOL_DOZA)
verifica("R3 dozare portie aliajAlV", calc3["rezultat"]["aliajAlV"], 2.46575342465753, TOL_DOZA)
verifica("R3 dozare portie alMetal", calc3["rezultat"]["alMetal"], 1.68893985826733, TOL_DOZA)

p = plan()
planificare.aplica_plan(ORDER, p)
print_plan(p)

verifica_egal("R3 are bara de report", bool(p[2]["report"]), True)
verifica_egal("R3 bare dozate aici", p[2]["bareNoi"], 1)
verifica_egal("R3 numar total de bare", len(r3["bare"]), 2)
verifica_egal("R3 nu mai cedeaza nicio bara", p[2]["desf"]["baraReport"], None)
verifica_egal("comanda: nicio reteta in plus", len(ORDER["retete"]), 3)
verifica_egal("comanda: total bare plasate",
              sum(len(r["bare"]) for r in ORDER["retete"]), 12)

# Bara 11 ia din lotul nou de burete diferenta pana la doza veche
verifica("R3 bara de report — burete luat din lotul nou",
         p[2]["report"]["ramas"].get("burete", 0.0),
         806.523206195451 - 412.549406456046)

ok, mesaj = planificare.aplica_consum_reteta(ORDER, r3, LOTS, calculeaza_bara, TIP_ALIAJ)
verifica_egal("R3 aplicare consum", ok, True)

print(f"   dupa R3 -> rest burete2 {rest('burete2'):.3f} | AlV {rest('alv2'):.3f} | "
      f"Al {rest('al1'):.3f} | TiO2 {rest('tio1'):.3f} | Fe {rest('fe1'):.3f}")
# Excel "Retete", randul de sub B12 al retetei 3 (I78:M78)
verifica("R3 rest burete 260409-945", rest("burete2"),
         7856.03000532186 - 806.519401134191)
verifica_rest("R3 rest AlV 361748", rest("alv2"), 780.366653380604 - 55.4794520547945,
              [(calc2, calc1, "aliajAlV"), (calc3, calc2, "aliajAlV")])
verifica_rest("R3 rest Al metal", rest("al1"), 457.603341297541 - 38.0011468110149,
              [(calc2, calc1, "alMetal"), (calc3, calc2, "alMetal")])

# Lotul vechi de burete trebuie sa fi fost consumat INTEGRAL (Excel
# DateIntrare I13: consum 8478 = tot stocul de intrare)
verifica("lot burete 260215-439 consumat integral", LOT_DUPA_ID["burete1"]["consum"], 8478.0, 0.02)
verifica("lot AlV 360744 consumat integral", LOT_DUPA_ID["alv1"]["consum"], 273.0, 0.02)

# ---------------------------------------------------------------------------
# FISA LIMITA — cantitatile de pe document trebuie sa fie EXACT cele
# scazute din loturi (inclusiv restul luat de barele de report din
# loturile golite in reteta anterioara).
# ---------------------------------------------------------------------------
print("\nFISA LIMITA — consumul de pe document fata de consumul real din loturi")
from dozare_titan.export.fisa_limita import calculeaza_consum_comanda   # noqa: E402

consum_doc = calculeaza_consum_comanda(ORDER, LOTS)
for mat in ORDINE_MAT:
    real = sum(l["consum"] for l in LOTS if l["material"] == mat)
    if real <= 1e-9 and consum_doc.get(mat, {}).get("total", 0.0) <= 1e-9:
        continue
    print(f"   {mat:10s} document {consum_doc[mat]['total']:10.2f} kg   "
          f"loturi {real:10.2f} kg")
    verifica(f"fisa limita {mat}", consum_doc[mat]["total"], real, 0.02)

# ---------------------------------------------------------------------------
# RETRAGEREA CONSUMULUI — in ordine inversa, reteta cu reteta
# ---------------------------------------------------------------------------
print("\nRETRAGEREA CONSUMULUI (in ordine inversa: R3, R2, R1)")

stare = planificare.stare_retragere(ORDER, ORDER["retete"][0])
verifica_egal("R1 nu se poate retrage inaintea R2 (bara cedata e deja consumata)",
              stare["blocat"], True)

for r, eticheta in ((r3, "R3"), (r2, "R2"), (ORDER["retete"][0], "R1")):
    ok, mesaj = planificare.retrage_consum_reteta(ORDER, r, LOTS)
    verifica_egal(f"{eticheta} retragere consum", ok, True)
    print(f"   {eticheta}: {mesaj}")

print(f"   dupa retragere -> consum burete1 {LOT_DUPA_ID['burete1']['consum']:.3f} | "
      f"burete2 {LOT_DUPA_ID['burete2']['consum']:.3f} | "
      f"alv1 {LOT_DUPA_ID['alv1']['consum']:.3f} | alv2 {LOT_DUPA_ID['alv2']['consum']:.3f} | "
      f"al1 {LOT_DUPA_ID['al1']['consum']:.3f}")
for l in LOTS:
    verifica(f"stoc repus la zero dupa retragere ({l['lot']})", l["consum"], 0.0, 0.001)

# ---------------------------------------------------------------------------
# DUPA RETRAGERE — replanificarea trebuie sa dea exact aceeasi distributie
# ---------------------------------------------------------------------------
print("\nREPLANIFICARE dupa retragerea totala (trebuie sa iasa identic: 4 / 1+5 / 1+1)")
p = plan()
planificare.aplica_plan(ORDER, p)
print_plan(p)
verifica_egal("replanificare R1", (p[0]["bareNoi"], bool(p[0]["report"])), (4, False))
verifica_egal("replanificare R2", (p[1]["bareNoi"], bool(p[1]["report"])), (5, True))
verifica_egal("replanificare R3", (p[2]["bareNoi"], bool(p[2]["report"])), (1, True))
verifica_egal("replanificare: tot 12 bare",
              sum(len(r["bare"]) for r in ORDER["retete"]), 12)
verifica_egal("replanificare: tot 3 retete", len(ORDER["retete"]), 3)

# ---------------------------------------------------------------------------
print("\n" + "=" * 78)
if erori:
    print(f"ESUAT — {len(erori)} din {verificari[0]} verificari nu au trecut:\n")
    for e in erori:
        print("  \u2717 " + e)
    sys.exit(1)
print(f"OK — toate cele {verificari[0]} verificari trec (toleranta {TOL} kg fata de Excel).")
print("""
Nota despre toleranta: diferentele fata de Excel sunt de ordinul zecilor de
grame si vin dintr-un singur loc — bara de report. In Excel, bara mutata pe
reteta urmatoare e scazuta acolo cu dozarea RETETEI NOI (asa e construit
sablonul, coloana scade mereu dozarea foii curente), pe cand aplicatia o
scade cu DOZAREA VECHE, cea cu care s-a presat efectiv bara, asa cum se
lucreaza in sectie. Ex. bara 5: burete 806.578 (dozarea retetei 1) in loc de
806.523 (dozarea retetei 2) — 55 de grame. Distributia barelor pe retete si
epuizarea loturilor sunt identice cu Excel-ul.
""")
