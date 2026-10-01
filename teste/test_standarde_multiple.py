"""Selectia MULTIPLA de standarde pe o comanda (ex. AMS 4975 + AMS 4976).

    python3 teste/test_standarde_multiple.py

Foloseste fisierul livrat cu aplicatia (dozare_titan/assets/standarde.xlsx).
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dozare_titan import standarde as st                              # noqa: E402
from dozare_titan.config import ALIAJE_SPEC                           # noqa: E402

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


FISIER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      "dozare_titan", "assets", "standarde.xlsx")
st.cai_posibile = lambda: [FISIER]
st.reincarca()

ASTM = "ASTM  B 348/ B348M \u2013 21"      # asa apare in Excel (cu spatii duble)

# ---------------------------------------------------------------------------
print("1. Citirea listei de standarde de pe comanda")
verifica("fara nimic ales", st.nume_standarde_alese({"tipAliaj": "Ti5"}), [])
verifica("format vechi (text) ramane valabil",
         st.nume_standarde_alese({"tipAliaj": "Ti5", "standard": "AMS 4928W"}), ["AMS 4928W"])
verifica("lista are prioritate fata de campul vechi",
         st.nume_standarde_alese({"standard": "AMS 4928W", "standarde": ["AMS 6931D"]}), ["AMS 6931D"])
verifica("lista goala explicita = nimic ales",
         st.nume_standarde_alese({"standard": "AMS 4928W", "standarde": []}), [])
verifica("duplicatele si spatiile se curata",
         st.nume_standarde_alese({"standarde": ["AMS 4928W", " ams  4928w ", "AMS 6931D"]}),
         ["AMS 4928W", "AMS 6931D"])
verifica("standard inexistent la aliaj se ignora",
         [s["nume"] for s in st.standarde_alese({"tipAliaj": "Ti5", "standarde": ["AMS 4928W", "ISO 5832-2:2025"]})],
         ["AMS 4928W"])
verifica("fara standarde alese -> None", st.standard_ales({"tipAliaj": "Ti5", "standarde": []}), None)

# ---------------------------------------------------------------------------
print("2. Setarea standardelor (seteaza_standarde)")
o = {"tipAliaj": "Ti5"}
verifica("prima setare = schimbare", st.seteaza_standarde(o, ["AMS 4928W", "AMS 6931D"]), True)
verifica("lista si campul vechi sincronizate", (o["standarde"], o["standard"]), (["AMS 4928W", "AMS 6931D"], "AMS 4928W"))
verifica("aceeasi lista = nicio schimbare", st.seteaza_standarde(o, ["AMS 4928W", "AMS 6931D"]), False)
verifica("golire", (st.seteaza_standarde(o, []), o["standarde"], o["standard"]), (True, [], ""))
st.seteaza_standarde(o, ["AMS 4928W", "nu exista", "AMS 4928W"])
verifica("nume inexistente si duplicate sunt scoase", o["standarde"], ["AMS 4928W"])
# comanda veche: doar "standard"; prima setare identica cu ea nu e o schimbare "reala" pentru date, dar
# trebuie sa completeze lista.
vechi = {"tipAliaj": "Ti5", "standard": "AMS 4928W"}
st.seteaza_standarde(vechi, ["AMS 4928W"])
verifica("comanda veche primeste lista", vechi["standarde"], ["AMS 4928W"])

# schimbarea aliajului pastreaza doar standardele comune
o = {"tipAliaj": "Ti5"}
st.seteaza_standarde(o, ["AMS 4928W", "ASTM  B 381 \u2013 21"])
verifica("Ti5: doua standarde alese", len(o["standarde"]), 2)
o["tipAliaj"] = "Titan grad 2"
verifica("schimbare aliaj = s-a schimbat ceva", st.pastreaza_standarde_valide(o), True)
verifica("ramane doar standardul comun cu grad 2", o["standarde"], ["ASTM  B 381 \u2013 21".replace("  ", " ")])
o["tipAliaj"] = "Ti 6-2-4-2"
st.pastreaza_standarde_valide(o)
verifica("aliaj fara standarde comune -> gol", (o["standarde"], o["standard"]), ([], ""))

# ---------------------------------------------------------------------------
print("3. Standarde combinate: cea mai stricta limita")
o = {"tipAliaj": "Ti5", "standarde": ["AMS 4928W", ASTM]}
std = st.standard_ales(o)
verifica("un standard = chiar acel standard (fara 'componente')", "componente" not in st.standard_ales(
    {"tipAliaj": "Ti5", "standarde": ["AMS 4928W"]}), True)
verifica("nume combinat", std["nume"], "AMS 4928W; ASTM B 348/ B348M \u2013 21")
verifica("doua componente", [c["nume"] for c in std["componente"]], ["AMS 4928W", "ASTM B 348/ B348M \u2013 21"])
aprox("Fe: max 0.30 (AMS) e mai strict decat 0.40 (ASTM)", std["limite"]["Fe"]["max"], 0.30)
aprox("H: 0.0125 (AMS) mai strict decat 0.015 (ASTM)", std["limite"]["H"]["max"], 0.0125)
verifica("Y exista doar la AMS -> ramane", std["limite"]["Y"], {"max": 0.005})
verifica("Al identic", std["limite"]["Al"], {"min": 5.5, "max": 6.75})
verifica("fara conflicte", std["conflicte"], [])
verifica("nu e N/A si are limite", (std["neaplicabil"], std["fara_limite"]), (False, False))

# N/A nu restrange nimic
std_na = st.standard_ales({"tipAliaj": "Ti5", "standarde": ["AMS 4928W", "AMS 2380H"]})
ams = st.standard_ales({"tipAliaj": "Ti5", "standarde": ["AMS 4928W"]})
verifica("AMS 2380H (N/A) nu schimba limitele", std_na["limite"], ams["limite"])
verifica("combinatia cu un N/A nu e N/A", std_na["neaplicabil"], False)
verifica("doar N/A -> neaplicabil",
         st.standard_ales({"tipAliaj": "Ti5", "standarde": ["AMS 2380H"]})["neaplicabil"], True)

# conflict intre standarde (sintetic)
a = {"nume": "A", "limite": {"Fe": {"min": 0.3, "max": 0.5}}, "fara_limite": False, "neaplicabil": False}
b = {"nume": "B", "limite": {"Fe": {"max": 0.2}}, "fara_limite": False, "neaplicabil": False}
c = st.combina_standarde([a, b])
verifica("conflict minim > maxim detectat", c["conflicte"], ["Fe"])
verifica("tooltip semnaleaza conflictul", "se contrazic" in st.text_standard(c), True)

# ---------------------------------------------------------------------------
print("4. Avertismente pe standarde alese impreuna")
target = {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.35}
v = st.verifica_limite(std, target)
verifica("Fe 0.35: depaseste doar AMS (ASTM permite 0.40)",
         [(i["element"], "AMS 4928W" in i["text"]) for i in v], [("Fe", True)])
v = st.verifica_limite(std, {**target, "fe": 0.45})
verifica("Fe 0.45: depaseste AMAMBELE, cu cate un mesaj pe standard",
         sorted("AMS" if "AMS 4928W" in i["text"] else "ASTM" for i in v), ["AMS", "ASTM"])
verifica("in limite la toate", st.verifica_limite(std, {"al": 6.3, "v": 4.0, "o": 0.17, "fe": 0.17}), [])
verifica("o valoare care iese la doua standarde nu se dubleaza per standard",
         len(st.verifica_limite(std, {"al": 9.0})), 2)

# ---------------------------------------------------------------------------
print("5. Formularele tiparite (spec_efectiv)")
spec = st.spec_efectiv(o)
verifica("titlu = standardele alese", spec["titlu_formular"], "AMS 4928W; ASTM B 348/ B348M \u2013 21")
aprox("Fe max tiparit = cea mai stricta limita (0.30% = 0.003)", spec["fe_max"], 0.003)
aprox("H max tiparit = 0.0125% = 0.000125", spec["h_max"], 0.000125)
spec1 = st.spec_efectiv({"tipAliaj": "Ti5", "standard": "AMS 4928W"})
verifica("un standard: comportament neschimbat", spec1["titlu_formular"], "AMS 4928W")
verifica("fara standard: spec-ul vechi neschimbat", st.spec_efectiv({"tipAliaj": "Ti5"}), ALIAJE_SPEC["Ti5"])

# ---------------------------------------------------------------------------
print("6. Comanda 24883 (Ti 6-2-4-2): AMS 4975 + AMS 4976")
o6 = {"tipAliaj": "Ti 6-2-4-2", "standarde": ["AMS 4976L", "AMS 4975P"]}
s6 = st.standard_ales(o6)
verifica("doua componente", [c["nume"] for c in s6["componente"]], ["AMS 4976L", "AMS 4975P"])
verifica("limite identice cu fiecare (AMS 4975P = AMS 4976L in Excel)",
         s6["limite"], st.standarde_pentru("Ti 6-2-4-2")[0]["limite"])
tinta_conc = {"al": 6.0, "mo": 2.0, "sn": 2.0, "zr": 4.0, "o": 0.1, "si": 0.08}    # "Conc.TINTA" din Excel
verifica("concentratia tinta 6/2/2/4 se incadreaza", st.verifica_limite(s6, tinta_conc), [])
v = st.verifica_limite(s6, {"al": 6.8})
verifica("Al 6.8% (doza) iese din standard, cate un mesaj per standard", len(v), 2)
verifica("Sn 2.1 in limite (1.8-2.2)", st.verifica_limite(s6, {"sn": 2.1, "mo": 1.96, "zr": 3.9, "si": 0.08}), [])

print("\n" + "=" * 70)
if erori:
    print(f"ESUAT — {len(erori)} din {nr[0]} verificari:\n")
    for e in erori:
        print("  \u2717 " + e)
    sys.exit(1)
print(f"OK — toate cele {nr[0]} verificari trec.")
