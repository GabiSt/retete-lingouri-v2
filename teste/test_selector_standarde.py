"""Verificare fara ecran a controlului de selectie multipla.

    QT_QPA_PLATFORM=offscreen python3 teste/test_selector_standarde.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication                            # noqa: E402
from dozare_titan.selector_standarde import SelectorStandarde         # noqa: E402

app = QApplication.instance() or QApplication([])
erori, nr = [], [0]


def verifica(eticheta, obtinut, asteptat):
    nr[0] += 1
    if obtinut != asteptat:
        erori.append(f"{eticheta}: obtinut {obtinut!r}, asteptat {asteptat!r}")


NUME = ["AMS 4976L", "AMS 4975P", "MS080"]
s = SelectorStandarde(NUME)
primite = []
s.schimbat.connect(primite.append)

verifica("gol: text implicit", s.text().startswith("\u2014 alege standardele"), True)
verifica("gol: nimic ales", s.alese(), [])

# Deschiderea meniului (simulata direct, fara popup real in mediul offscreen)
s._la_deschidere()

s._casute["AMS 4975P"].setChecked(True)
verifica("un standard bifat: textul e numele", s.text().startswith("AMS 4975P"), True)
s._casute["AMS 4976L"].setChecked(True)
verifica("ordinea = ordinea listei, nu a bifarii", s.alese(), ["AMS 4976L", "AMS 4975P"])
verifica("doua: text cu +1", "AMS 4976L (+1)" in s.text(), True)
verifica("NU s-a emis inca niciun semnal in timp ce meniul e deschis (bifezi mai multe)",
         primite, [])

# Inchiderea meniului: ABIA acum pleaca semnalul, o singura data, cu selectia finala
s._la_inchidere()
verifica("la inchidere: semnalul a plecat o singura data, cu ambele alese",
         primite, [["AMS 4976L", "AMS 4975P"]])

# O noua deschidere/inchidere FARA nicio bifare noua -> nu se mai emite nimic
s._la_deschidere()
s._la_inchidere()
verifica("deschis/inchis fara schimbari: niciun semnal nou", len(primite), 1)

# Debifare intre o deschidere si o inchidere -> un singur semnal nou, la inchidere
s._la_deschidere()
s._casute["AMS 4975P"].setChecked(False)
verifica("debifare: inca niciun semnal cat timp meniul e deschis", len(primite), 1)
s._la_inchidere()
verifica("debifare: semnalul pleaca la inchidere, cu selectia ramasa",
         primite[-1], ["AMS 4976L"])

# seteaza_alese nu emite semnal si actualizeaza textul
n = len(primite)
s.seteaza_alese(["MS080", "AMS 4975P"])
verifica("seteaza_alese: fara semnal", len(primite), n)
verifica("seteaza_alese: bifele", s.alese(), ["AMS 4975P", "MS080"])

# meniul ramane deschis la bifare: casutele sunt widget-uri in meniu, nu actiuni care il inchid
verifica("casutele sunt in meniu", len(s.menu().actions()), 3)

# preinitializare cu standarde alese
s2 = SelectorStandarde(NUME, alese=["AMS 4976L", "MS080"])
verifica("preinitializare", s2.alese(), ["AMS 4976L", "MS080"])
s3 = SelectorStandarde([], text_gol="\u2014 fara standarde in Excel \u2014")
verifica("fara standarde in Excel", (s3.alese(), "fara standarde" in s3.text()), ([], True))

print("=" * 70)
if erori:
    print(f"ESUAT — {len(erori)} din {nr[0]}:")
    for e in erori:
        print("  \u2717 " + e)
    sys.exit(1)
print(f"OK — toate cele {nr[0]} verificari trec.")
