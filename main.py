#!/usr/bin/env python3
"""Dozare Lingouri Titan — aplicatie desktop (PySide6).

Gestionare stoc de loturi pe materiale si calcul de retete de sarja
(compozitie chimica tinta -> cantitati per material).

Ruleaza cu:
    python3 main.py

Structura completa a codului e explicata in README.md si in
dozare_titan/__init__.py.
"""

import sys

from PySide6.QtWidgets import QApplication

from dozare_titan.fereastra_principala import FereastraDozareTitan


def main():
    app = QApplication(sys.argv)
    fereastra = FereastraDozareTitan()
    fereastra.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
