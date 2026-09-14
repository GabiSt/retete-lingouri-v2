#!/usr/bin/env python3
"""Dozare Lingouri Titan — aplicatie desktop (PySide6).

Gestionare stoc de loturi pe materiale si calcul de retete de sarja
(compozitie chimica tinta -> cantitati per material).

Ruleaza cu:
    python3 main.py

La pornire se cere autentificare (utilizator + parola, vezi
dozare_titan/config.py -> UTILIZATORI); privilegiile de editare si numele
afisat la rubrica "Intocmit" pe documentele generate depind de contul cu
care te loghezi.

Structura completa a codului e explicata in README.md si in
dozare_titan/__init__.py.
"""

import sys

from PySide6.QtWidgets import QApplication, QDialog

from dozare_titan.dialoguri import DialogLogin
from dozare_titan.fereastra_principala import FereastraDozareTitan


def main():
    app = QApplication(sys.argv)

    while True:
        login = DialogLogin()
        if login.exec() != QDialog.Accepted or not login.rezultat:
            sys.exit(0)

        fereastra = FereastraDozareTitan(login.rezultat)
        fereastra.show()
        app.exec()

        if not getattr(fereastra, "se_delogheaza", False):
            break

    sys.exit(0)


if __name__ == "__main__":
    main()
