"""Ferestre de dialog: autentificare admin, adaugare/editare lot, istoric lot.

Daca un aliaj nou are elemente in plus in compozitia unui lot (ex. Si, Zr,
Mo pentru Ti-VT9), locul de adaugat campurile noi e in DialogLotNou,
variabila `etichete_doza` din __init__.
"""

from datetime import date

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QDialog, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
    QMessageBox, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from .config import ADMIN_PASS, MATERIALE
from .stiluri import STIL_BUTON_PRINCIPAL, STIL_BUTON_SECUNDAR, STIL_CAMP, CULOARE_EROARE, CULOARE_GRI_TEXT, CULOARE_BORDURA, CULOARE_FUNDAL_SECTIUNE
from .utils import fmt, to_float
from .calcule import lot_ti, lot_rest


class DialogLoginAdmin(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Autentificare admin")
        self.setFixedWidth(340)
        layout = QVBoxLayout(self)

        info = QLabel("Doar administratorul poate adauga loturi noi si corecta stocurile.")
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {CULOARE_GRI_TEXT};")
        layout.addWidget(info)

        layout.addWidget(QLabel("Parola"))
        self.camp_parola = QLineEdit()
        self.camp_parola.setEchoMode(QLineEdit.Password)
        self.camp_parola.setStyleSheet(STIL_CAMP)
        self.camp_parola.returnPressed.connect(self._verifica)
        layout.addWidget(self.camp_parola)

        self.eticheta_eroare = QLabel("")
        self.eticheta_eroare.setStyleSheet(f"color: {CULOARE_EROARE}; font-size: 12px;")
        self.eticheta_eroare.setWordWrap(True)
        layout.addWidget(self.eticheta_eroare)

        rand = QHBoxLayout()
        btn_anuleaza = QPushButton("Anuleaza")
        btn_anuleaza.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_anuleaza.clicked.connect(self.reject)
        btn_intra = QPushButton("Intra")
        btn_intra.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_intra.clicked.connect(self._verifica)
        rand.addWidget(btn_anuleaza)
        rand.addWidget(btn_intra)
        layout.addLayout(rand)

        self.camp_parola.setFocus()

    def _verifica(self):
        if self.camp_parola.text() == ADMIN_PASS:
            self.accept()
        else:
            self.eticheta_eroare.setText("Parola incorecta.")
            self.camp_parola.clear()
            self.camp_parola.setFocus()


class DialogLotNou(QDialog):
    # Campurile de doza (compozitie) ale unui lot. Daca un aliaj nou are
    # nevoie de elemente noi (ex. Si, Zr, Mo pentru Ti-VT9), adauga-le aici
    # ca perechi (cheie_in_lot, eticheta_afisata) — cheia trebuie sa
    # inceapa cu "doza" pentru consistenta cu restul codului (dozaO,
    # dozaFe... -> dozaSi, dozaZr, dozaMo).
    ETICHETE_DOZA = [
        ("dozaO", "Doza O (%)"), ("dozaFe", "Doza Fe (%)"),
        ("dozaN", "Doza N (%)"), ("dozaV", "Doza V (%)"),
        ("dozaAl", "Doza Al (%)"), ("dozaSi", "Doza Si (%)"),
        ("dozaMo", "Doza Mo (%)"), ("dozaZr", "Doza Zr (%)"),
    ]

    def __init__(self, material_nume, parent=None, lot=None, orders=None):
        super().__init__(parent)
        self.editare = lot is not None
        self.lot_original = lot
        self.setWindowTitle(f"Editeaza lot \u2014 {material_nume}" if self.editare else f"Lot nou \u2014 {material_nume}")
        self.setFixedWidth(420)
        self.rezultat = None

        layout = QVBoxLayout(self)
        info = QLabel(
            "Modifica compozitia, stocul de intrare si/sau consumul acestui lot."
            if self.editare else
            "Introdu datele de doza si stocul initial pentru acest lot."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {CULOARE_GRI_TEXT};")
        layout.addWidget(info)

        grid = QGridLayout()
        grid.setSpacing(8)

        self.camp_nr_lot = QLineEdit(lot.get("nrLot", "") if self.editare else "")
        self.camp_nr_lot.setPlaceholderText("ex. TL-2026-014")
        self.camp_stoc = QLineEdit(fmt(lot.get("stocIntrare"), 3) if self.editare else "")
        self.camp_stoc.setPlaceholderText("0")

        grid.addWidget(QLabel("Nr. LOT"), 0, 0)
        grid.addWidget(self.camp_nr_lot, 1, 0)
        grid.addWidget(QLabel("Stoc intrare (kg)"), 0, 1)
        grid.addWidget(self.camp_stoc, 1, 1)

        self.camp_consum = None
        if self.editare:
            self.camp_consum = QLineEdit(fmt(lot.get("consum"), 3))
            grid.addWidget(QLabel("Consum (kg)"), 0, 2)
            grid.addWidget(self.camp_consum, 1, 2)

        self.campuri_doza = {}
        for i, (cheie, text) in enumerate(self.ETICHETE_DOZA):
            valoare_initiala = fmt(lot.get(cheie), 3) if self.editare else "0"
            camp = QLineEdit(valoare_initiala)
            camp.textChanged.connect(self._actualizeaza_preview_ti)
            self.campuri_doza[cheie] = camp
            rand = 2 + i // 3
            col = i % 3
            grid.addWidget(QLabel(text), rand * 2, col)
            grid.addWidget(camp, rand * 2 + 1, col)

        layout.addLayout(grid)

        campuri_stil = [self.camp_nr_lot, self.camp_stoc] + list(self.campuri_doza.values())
        if self.camp_consum is not None:
            campuri_stil.append(self.camp_consum)
        for camp in campuri_stil:
            camp.setStyleSheet(STIL_CAMP)

        self.eticheta_ti = QLabel("Ti calculat: 100.000%")
        self.eticheta_ti.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
        layout.addWidget(self.eticheta_ti)
        self._actualizeaza_preview_ti()

        eticheta_rezervare = QLabel("Rezervat exclusiv pentru comanda (optional)")
        eticheta_rezervare.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10.5px;")
        layout.addWidget(eticheta_rezervare)
        self.combo_rezervare = QComboBox()
        self.combo_rezervare.setStyleSheet(STIL_CAMP)
        self.combo_rezervare.setToolTip(
            "Daca alegi o comanda, acest lot nu va putea fi folosit in retetele altor comenzi "
            "(va aparea un avertisment la selectare)."
        )
        self.combo_rezervare.addItem("\u2014 fara rezervare \u2014", None)
        rezervat_id = lot.get("rezervatComandaId") if self.editare else None
        index_sel = 0
        for i, o in enumerate(orders or [], start=1):
            self.combo_rezervare.addItem(o.get("nume", ""), o["id"])
            if o["id"] == rezervat_id:
                index_sel = i
        self.combo_rezervare.setCurrentIndex(index_sel)
        layout.addWidget(self.combo_rezervare)

        self.eticheta_eroare = QLabel("")
        self.eticheta_eroare.setStyleSheet(f"color: {CULOARE_EROARE}; font-size: 12px;")
        self.eticheta_eroare.setWordWrap(True)
        layout.addWidget(self.eticheta_eroare)

        rand_btn = QHBoxLayout()
        btn_anuleaza = QPushButton("Anuleaza")
        btn_anuleaza.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_anuleaza.clicked.connect(self.reject)
        btn_salveaza = QPushButton("Salveaza modificarile" if self.editare else "Salveaza lot")
        btn_salveaza.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_salveaza.clicked.connect(self._salveaza)
        rand_btn.addWidget(btn_anuleaza)
        rand_btn.addWidget(btn_salveaza)
        layout.addLayout(rand_btn)

    def _actualizeaza_preview_ti(self):
        total = sum(to_float(c.text()) for c in self.campuri_doza.values())
        titlu = self.windowTitle()
        if "TiO2" in titlu:
            self.eticheta_ti.setText("Ti: 60% (fix)")
        elif "Burete" in titlu or "SiTi" in titlu:
            self.eticheta_ti.setText(f"Ti calculat: {100 - total:.3f}%")
        else:
            self.eticheta_ti.setText("Ti: 0% (material fara Ti)")

    def _salveaza(self):
        nr_lot = self.camp_nr_lot.text().strip()
        stoc = to_float(self.camp_stoc.text(), -1)
        if not nr_lot:
            self.eticheta_eroare.setText("Introdu numarul lotului.")
            return
        if stoc <= 0:
            self.eticheta_eroare.setText("Introdu un stoc de intrare valid.")
            return
        consum = to_float(self.camp_consum.text()) if self.camp_consum is not None else 0
        if consum < 0:
            self.eticheta_eroare.setText("Consumul nu poate fi negativ.")
            return
        rezervare_id = self.combo_rezervare.currentData()
        self.rezultat = {
            "nrLot": nr_lot, "stocIntrare": stoc, "consum": consum,
            "data": self.lot_original.get("data") if self.editare else date.today().strftime("%d.%m.%Y"),
            "rezervatComandaId": rezervare_id,
            "rezervatComandaNume": self.combo_rezervare.currentText() if rezervare_id else None,
        }
        for cheie, camp in self.campuri_doza.items():
            self.rezultat[cheie] = to_float(camp.text())
        self.accept()


class DialogIstoricLot(QDialog):
    """Afiseaza toate comenzile/retetele/barele in care un lot a fost
    consumat REAL (consum aplicat pe stoc), pe baza consumSnapshot-ului
    inregistrat la aplicare — aceeasi sursa de date ca la Fisa limita."""

    def __init__(self, lot, mat_nume, istoric, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Istoric lot \u2014 {lot.get('nrLot', '')}")
        self.resize(600, 440)
        layout = QVBoxLayout(self)

        info = QLabel(
            f"<b>{mat_nume}</b> \u2014 Lot {lot.get('nrLot', '')}<br>"
            f"Stoc intrare: {fmt(lot.get('stocIntrare'))} kg &nbsp;\u00b7&nbsp; "
            f"Consum total: {fmt(lot.get('consum'))} kg &nbsp;\u00b7&nbsp; "
            f"Rest: {fmt(lot_rest(lot))} kg"
        )
        info.setTextFormat(Qt.RichText)
        info.setWordWrap(True)
        layout.addWidget(info)

        if not istoric:
            gol = QLabel(
                "Acest lot nu a fost consumat inca in nicio comanda "
                "(niciun consum aplicat pe stoc)."
            )
            gol.setWordWrap(True)
            gol.setAlignment(Qt.AlignCenter)
            gol.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; padding: 30px;")
            layout.addWidget(gol)
        else:
            coloane = ["Comanda", "Data", "Reteta", "Bara", "Material", "Consum din acest lot"]
            tabel = QTableWidget(len(istoric), len(coloane))
            tabel.setHorizontalHeaderLabels(coloane)
            tabel.verticalHeader().setVisible(False)
            tabel.setEditTriggers(QTableWidget.NoEditTriggers)
            tabel.setSelectionMode(QTableWidget.NoSelection)
            tabel.setStyleSheet(
                f"QTableWidget {{ border: 1px solid {CULOARE_BORDURA}; gridline-color: {CULOARE_BORDURA}; }}"
                f"QHeaderView::section {{ background-color: {CULOARE_FUNDAL_SECTIUNE}; color: {CULOARE_GRI_TEXT}; "
                f"border: none; border-bottom: 1px solid {CULOARE_BORDURA}; padding: 6px; font-size: 10.5px; }}"
            )
            tabel.horizontalHeader().setStretchLastSection(True)

            total = 0.0
            for r, intrare in enumerate(istoric):
                nume_mat = next((m["nume"] for m in MATERIALE if m["id"] == intrare["material"]), intrare["material"])
                valori = [
                    intrare["comanda"], intrare["data"], intrare["reteta"],
                    f"B{intrare['bara']}", nume_mat, f"{fmt(intrare['kg'])} kg",
                ]
                for c, val in enumerate(valori):
                    item = QTableWidgetItem(str(val))
                    if c >= 3:
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    tabel.setItem(r, c, item)
                total += intrare["kg"]

            tabel.resizeRowsToContents()
            layout.addWidget(tabel)

            eticheta_total = QLabel(f"Total consumat din acest lot (in comenzile de mai sus): {fmt(total)} kg")
            eticheta_total.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
            layout.addWidget(eticheta_total)

        btn_inchide = QPushButton("Inchide")
        btn_inchide.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_inchide.clicked.connect(self.accept)
        layout.addWidget(btn_inchide)
