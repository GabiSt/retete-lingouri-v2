"""Verificarea GENERARII AUTOMATE a retetelor, pe acelasi caz real:
Cda 26858-Ti5-2VAR-d600-L1-L4 (12 bare de 900 kg).

    python3 teste/test_generare_automata.py

Aici utilizatorul NU mai creeaza retete si nu mai alege loturi pe fiecare
reteta: pune pe COMANDA lista ordonata de loturi (pe material) si atat.
Programul trebuie sa scoata singur exact aceleasi 3 retete ca Excel-ul:

    Reteta 1: barele 1-4                     (se termina AlV 360744)
    Reteta 2: bara 5 (report) + barele 6-10  (se termina buretele 260215-439)
    Reteta 3: bara 11 (report) + bara 12
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dozare_titan.calcule import calculeaza_bara, lot_rest        # noqa: E402
from dozare_titan.calcule import planificare                      # noqa: E402

TIP_ALIAJ = "Ti5"
erori = []
nr = [0]


def verifica(eticheta, obtinut, asteptat, tol=None):
    nr[0] += 1
    ok = (abs(obtinut - asteptat) <= tol) if tol is not None else (obtinut == asteptat)
    if not ok:
        erori.append(f"{eticheta}: obtinut {obtinut!r}, asteptat {asteptat!r}")


def lot(id_, material, nume, stoc, **doze):
    l = {"id": id_, "material": material, "lot": nume, "nrLot": nume,
         "stocIntrare": stoc, "consum": 0.0}
    l.update(doze)
    return l


LOTS = [
    lot("burete1", "burete", "260215-439", 8478, dozaO=0.027, dozaFe=0.008),
    lot("burete2", "burete", "260409-945", 8250, dozaO=0.023, dozaFe=0.003),
    lot("alv1", "aliajAlV", "360744", 273, dozaV=65.3, dozaAl=34.3, dozaO=0.05, dozaFe=0.2),
    lot("alv2", "aliajAlV", "361748", 1119, dozaV=65.7, dozaAl=33.8, dozaO=0.1, dozaFe=0.2),
    lot("al1", "alMetal", "W660078186", 874, dozaAl=99.87, dozaFe=0.07),
    lot("fe1", "feMetal", "lot nou", 923, dozaFe=99.0),
    lot("tio1", "tio2", "LOT 0087718", 607, dozaO=40.0),
]
LOT_DUPA_ID = {l["id"]: l for l in LOTS}

# Tot ce introduce utilizatorul: datele comenzii + listele de loturi.
ORDER = {
    "id": "cda26858", "nume": "Cda 26858-Ti5", "tipAliaj": TIP_ALIAJ,
    "nrBare": "12", "portie": "40", "numarPresari": "22.5",
    "target": {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.17},
    "loturiAlocate": {
        "burete": ["burete1", "burete2"],
        "aliajAlV": ["alv1", "alv2"],
        "alMetal": ["al1"],
        "feMetal": ["fe1"],
        "tio2": ["tio1"],
    },
    "retete": [],
}

print(__doc__.splitlines()[1])
print("=" * 78)

retete, raport = planificare.genereaza_retete(ORDER, LOTS, calculeaza_bara, TIP_ALIAJ)
ORDER["retete"] = retete

print(f"\nGenerate automat: {raport['generate']} retete, "
      f"{raport['barePlasate']} bare plasate, {raport['bareRamase']} ramase")
for r in retete:
    tipuri = [b.get("tipDozare") for b in r["bare"]]
    loturi = " / ".join(
        LOT_DUPA_ID[r["lotSel"][k]]["lot"] for k in ("burete", "aliajAlV", "alMetal")
        if r["lotSel"].get(k)
    )
    print(f"   {r['nume']}: {len(r['bare'])} bare "
          f"({tipuri.count('report')} report + {tipuri.count('normal')} noi)   loturi: {loturi}")

verifica("numar de retete generate", len(retete), 3)
verifica("bare plasate", raport["barePlasate"], 12)
verifica("bare ramase", raport["bareRamase"], 0)
verifica("avertismente", raport["avertismente"], [])

verifica("R1 numar bare", len(retete[0]["bare"]), 4)
verifica("R1 fara bara de report", retete[0]["bare"][0].get("tipDozare"), "normal")
verifica("R1 lot burete", retete[0]["lotSel"]["burete"], "burete1")
verifica("R1 lot AlV", retete[0]["lotSel"]["aliajAlV"], "alv1")

verifica("R2 numar bare", len(retete[1]["bare"]), 6)
verifica("R2 prima bara e de report", retete[1]["bare"][0].get("tipDozare"), "report")
verifica("R2 report vine de pe reteta 1",
         retete[1]["bare"][0].get("retetaSursaDozare"), "Reteta 1")
verifica("R2 a intrat automat in lotul urmator de AlV",
         retete[1]["lotSel"]["aliajAlV"], "alv2")
verifica("R2 pastreaza acelasi burete", retete[1]["lotSel"]["burete"], "burete1")

verifica("R3 numar bare", len(retete[2]["bare"]), 2)
verifica("R3 prima bara e de report", retete[2]["bare"][0].get("tipDozare"), "report")
verifica("R3 a intrat automat in lotul urmator de burete",
         retete[2]["lotSel"]["burete"], "burete2")
verifica("R3 pastreaza acelasi AlV", retete[2]["lotSel"]["aliajAlV"], "alv2")

# Dozarea barei de report e cea VECHE, nu cea a retetei in care ajunge
doza_r1 = calculeaza_bara(TIP_ALIAJ, ORDER["target"],
                          planificare.loturi_selectate(LOTS, retete[0]), "40", "22.5")
verifica("bara 5 pastreaza dozarea retetei 1",
         retete[1]["bare"][0]["calcSnapshot"]["rezultatTotal"]["burete"],
         doza_r1["rezultatTotal"]["burete"], 0.001)

# ---------------------------------------------------------------------------
# Aplicarea consumului, reteta cu reteta — trebuie sa dea aceleasi cifre ca
# in Excel (tab "Retete", tabelele "Bilant presare")
# ---------------------------------------------------------------------------
print("\nAplicarea consumului pe cele 3 retete generate:")
for r in retete:
    ok, mesaj = planificare.aplica_consum_reteta(ORDER, r, LOTS, calculeaza_bara, TIP_ALIAJ)
    verifica(f"{r['nume']} aplicare consum", ok, True)
    print(f"   {r['nume']}: {mesaj}")

print(f"   rest burete 260215-439 {lot_rest(LOT_DUPA_ID['burete1']):.2f} | "
      f"260409-945 {lot_rest(LOT_DUPA_ID['burete2']):.2f} | "
      f"AlV 360744 {lot_rest(LOT_DUPA_ID['alv1']):.2f} | "
      f"361748 {lot_rest(LOT_DUPA_ID['alv2']):.2f} | "
      f"Al {lot_rest(LOT_DUPA_ID['al1']):.2f}")

verifica("lot burete 260215-439 golit complet", lot_rest(LOT_DUPA_ID["burete1"]), 0.0, 0.02)
verifica("lot AlV 360744 golit complet", lot_rest(LOT_DUPA_ID["alv1"]), 0.0, 0.02)
# Excel "Retete" I78 / J78 / K78 (dupa bara 12), corectat cu diferenta
# cunoscuta de dozare a barelor de report (vezi test_cda_26858.py)
verifica("rest burete 260409-945 dupa bara 12", lot_rest(LOT_DUPA_ID["burete2"]),
         7856.03000532186 - 806.519401134191, 0.12)
verifica("rest Al metal dupa bara 12", lot_rest(LOT_DUPA_ID["al1"]),
         457.603341297541 - 38.0011468110149, 0.5)

# ---------------------------------------------------------------------------
# Regenerarea dupa ce primele retete au consumul aplicat: cele consumate
# raman neatinse.
# ---------------------------------------------------------------------------
print("\nRegenerare dupa aplicarea consumului (retetele consumate raman neatinse):")
retete2, raport2 = planificare.genereaza_retete(ORDER, LOTS, calculeaza_bara, TIP_ALIAJ)
print(f"   pastrate {raport2['pastrate']}, generate {raport2['generate']}, "
      f"bare plasate {raport2['barePlasate']}")
verifica("regenerare: toate retetele sunt pastrate", raport2["pastrate"], 3)
verifica("regenerare: nu se mai genereaza nimic", raport2["generate"], 0)
verifica("regenerare: tot 12 bare", raport2["barePlasate"], 12)
verifica("regenerare: aceleasi obiecte de reteta",
         [r["id"] for r in retete2], [r["id"] for r in retete])

# ---------------------------------------------------------------------------
# Cazuri limita: loturi insuficiente si loturi goale in lista
# ---------------------------------------------------------------------------
print("\nCAZ LIMITA — mai multe bare decat ajung loturile (nu trebuie sa se blocheze):")
LOTS2 = [
    lot("b1", "burete", "burete mic", 3000, dozaO=0.027, dozaFe=0.008),
    lot("b0", "burete", "burete gol", 0, dozaO=0.027, dozaFe=0.008),
    lot("b2", "burete", "burete 2", 2000, dozaO=0.023, dozaFe=0.003),
    lot("a1", "aliajAlV", "alv", 900, dozaV=65.3, dozaAl=34.3, dozaO=0.05, dozaFe=0.2),
    lot("m1", "alMetal", "al", 800, dozaAl=99.87, dozaFe=0.07),
    lot("f1", "feMetal", "fe", 900, dozaFe=99.0),
    lot("t1", "tio2", "tio2", 600, dozaO=40.0),
]
ORDER2 = {
    "id": "x", "nume": "Comanda test", "tipAliaj": TIP_ALIAJ, "nrBare": "20",
    "portie": "40", "numarPresari": "22.5",
    "target": {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.17},
    "loturiAlocate": {"burete": ["b1", "b0", "b2"], "aliajAlV": ["a1"],
                      "alMetal": ["m1"], "feMetal": ["f1"], "tio2": ["t1"]},
    "retete": [],
}
retete3, raport3 = planificare.genereaza_retete(ORDER2, LOTS2, calculeaza_bara, TIP_ALIAJ)
print(f"   {raport3['generate']} retete, {raport3['barePlasate']} bare plasate, "
      f"{raport3['bareRamase']} ramase")
for a in raport3["avertismente"]:
    print("   \u26a0 " + a)
nr[0] += 1
if not raport3["avertismente"]:
    erori.append("caz limita: trebuia un avertisment ca nu ajung loturile")
verifica("caz limita: nu se plaseaza mai mult decat se poate",
         raport3["barePlasate"] + raport3["bareRamase"], 20)
verifica("caz limita: lotul gol din lista e sarit",
         all(r["lotSel"].get("burete") != "b0" for r in retete3), True)

print("\n" + "=" * 78)
if erori:
    print(f"ESUAT — {len(erori)} din {nr[0]} verificari:\n")
    for e in erori:
        print("  \u2717 " + e)
    sys.exit(1)
print(f"OK — toate cele {nr[0]} verificari trec.")
