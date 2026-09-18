"""Fereastra principala a aplicatiei: tab-ul de stoc pe loturi si tab-ul
de comenzi/retete, cu bilantul de dozare per bara si per reteta."""

import os
from datetime import date

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFrame, QGridLayout, QHBoxLayout,
    QInputDialog, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea,
    QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
    QFileDialog,
)

from .config import (
    MATERIALE, ORDINE_MAT, ALIAJ_IMPLICIT, ALIAJE_SPEC, ALIAJE_DISPONIBILE,
    BENEFICIAR_IMPLICIT,
)
from .stiluri import (
    CULOARE_BLEUMARIN, CULOARE_EROARE, CULOARE_SUCCES, CULOARE_AVERTISMENT,
    CULOARE_GRI_TEXT, CULOARE_FUNDAL_SECTIUNE, CULOARE_BORDURA,
    STIL_BUTON_PRINCIPAL, STIL_BUTON_SECUNDAR, STIL_BUTON_PERICOL, STIL_CAMP,
    TAG_STYLES, _clear_layout,
)
from .utils import uid, fmt, to_float, r2
from .calcule import (
    calculeaza_bara, lot_ti, lot_rest, target_ti, material_necesar,
    desfasoara_reteta, nume_material, TOLERANTA_KG,
)
from .persistenta import incarca_date, salveaza_date, _numar_bare_anterioare
from .dialoguri import (
    DialogLotNou, DialogIstoricLot, DialogAdministrareConturi,
    DialogCapacitateReteta,
)
from .export.fisa_limita import (
    gaseste_istoric_lot, calculeaza_consum_comanda,
    genereaza_fisa_limita_xlsx, genereaza_fisa_limita_pdf,
    OPENPYXL_DISPONIBIL, REPORTLAB_DISPONIBIL,
)
from .export.retdozare import genereaza_retdozare_xlsx, genereaza_retdozare_pdf


class ComboFaraScroll(QComboBox):
    """QComboBox care ignora scroll-ul rotitei mouse-ului, ca sa nu se
    schimbe lotul selectat accidental cand utilizatorul doar deruleaza
    pagina peste combo, fara sa dea click pe el."""

    def wheelEvent(self, event):
        event.ignore()


class FereastraDozareTitan(QWidget):
    def __init__(self, utilizator):
        super().__init__()
        self.setWindowTitle("Dozare lingouri titan")
        self.resize(1180, 780)
        self.setStyleSheet("background-color: white;")

        self.state = incarca_date()
        self.utilizator = utilizator
        self.poate_edita = bool(utilizator.get("editor"))
        self.este_admin = bool(utilizator.get("admin"))
        self.se_delogheaza = False
        self.mat_expanded = {m["id"]: True for m in MATERIALE}
        self.order_expanded = {}

        self._construieste_interfata()
        self.refresh()

    # ------------------------------------------------------------------
    def _construieste_interfata(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        layout.addLayout(self._construieste_bara_sus())

        self.tabs = QTabWidget()
        self.tab_stoc = QWidget()
        self.tab_retete = QWidget()
        self.tabs.addTab(self.tab_stoc, "Stoc loturi")
        self.tabs.addTab(self.tab_retete, "Comenzi si retete")
        layout.addWidget(self.tabs)

        stoc_scroll = QScrollArea()
        stoc_scroll.setWidgetResizable(True)
        stoc_scroll.setStyleSheet("border: none;")
        self.stoc_content = QWidget()
        self.stoc_layout = QVBoxLayout(self.stoc_content)
        self.stoc_layout.setSpacing(14)
        stoc_scroll.setWidget(self.stoc_content)
        stoc_tab_layout = QVBoxLayout(self.tab_stoc)
        stoc_tab_layout.setContentsMargins(0, 12, 0, 0)
        stoc_tab_layout.addWidget(stoc_scroll)

        retete_scroll = QScrollArea()
        retete_scroll.setWidgetResizable(True)
        retete_scroll.setStyleSheet("border: none;")
        self.retete_content = QWidget()
        self.retete_layout = QVBoxLayout(self.retete_content)
        self.retete_layout.setSpacing(14)
        retete_scroll.setWidget(self.retete_content)
        retete_tab_layout = QVBoxLayout(self.tab_retete)
        retete_tab_layout.setContentsMargins(0, 12, 0, 0)
        retete_tab_layout.addWidget(retete_scroll)

    def _construieste_bara_sus(self):
        bara = QHBoxLayout()
        text = QVBoxLayout()
        titlu = QLabel("Dozare lingouri titan")
        titlu.setFont(QFont("Sans Serif", 16, QFont.Bold))
        titlu.setStyleSheet(f"color: {CULOARE_BLEUMARIN};")
        subtitlu = QLabel("Stoc loturi & calcul retete de sarja")
        subtitlu.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 12px;")
        text.addWidget(titlu)
        text.addWidget(subtitlu)
        bara.addLayout(text)
        bara.addStretch(1)

        self.zona_admin = QHBoxLayout()
        bara.addLayout(self.zona_admin)
        return bara

    def _refresh_zona_admin(self):
        _clear_layout(self.zona_admin)
        rol = "Editor" if self.poate_edita else "Doar vizualizare"
        culoare_rol = CULOARE_SUCCES if self.poate_edita else CULOARE_GRI_TEXT
        eticheta = QLabel(f"{self.utilizator.get('alias', '')} \u00b7 {rol}")
        eticheta.setStyleSheet(
            f"color: {culoare_rol}; font-weight: 600; font-size: 11px; "
            f"padding: 4px 10px; border: 1px solid {culoare_rol}; border-radius: 10px;"
        )
        self.zona_admin.addWidget(eticheta)
        if self.este_admin:
            btn_admin = QPushButton("Administrare conturi")
            btn_admin.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_admin.clicked.connect(self._administrare_conturi)
            self.zona_admin.addWidget(btn_admin)
        btn = QPushButton("Delogare")
        btn.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn.clicked.connect(self._delogare)
        self.zona_admin.addWidget(btn)

    def _administrare_conturi(self):
        cheie_curenta = (self.utilizator.get("utilizator") or "").strip().lower()
        dialog = DialogAdministrareConturi(cheie_curenta, self)
        dialog.exec()

    def _delogare(self):
        if QMessageBox.question(self, "Confirmare", "Te deloghezi din aplicatie?") != QMessageBox.Yes:
            return
        self.se_delogheaza = True
        self.close()

    # ------------------------------------------------------------------
    def refresh(self):
        self._refresh_zona_admin()
        self._rebuild_stoc()
        self._rebuild_retete()

    # ------------------------ TAB: STOC LOTURI ------------------------
    def _rebuild_stoc(self):
        _clear_layout(self.stoc_layout)
        for mat in MATERIALE:
            self.stoc_layout.addWidget(self._construieste_sectiune_material(mat))
        self.stoc_layout.addStretch(1)

    def _construieste_sectiune_material(self, mat):
        lots = [l for l in self.state["lots"] if l["material"] == mat["id"]]
        total_intrare = sum(to_float(l.get("stocIntrare")) for l in lots)
        total_consum = sum(to_float(l.get("consum")) for l in lots)
        total_rest = total_intrare - total_consum
        expanded = self.mat_expanded.get(mat["id"], True)

        cadru = QFrame()
        cadru.setStyleSheet(f"QFrame {{ border: 1px solid {CULOARE_BORDURA}; border-radius: 8px; }}")
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        antet = QFrame()
        antet.setStyleSheet(f"background-color: {CULOARE_FUNDAL_SECTIUNE};")
        antet_layout = QHBoxLayout(antet)
        antet_layout.setContentsMargins(14, 10, 14, 10)

        buton_titlu = QPushButton(("\u25be " if expanded else "\u25b8 ") + f"{mat['nume']}  ({len(lots)} loturi)")
        buton_titlu.setFlat(True)
        buton_titlu.setCursor(Qt.PointingHandCursor)
        buton_titlu.setStyleSheet(
            "QPushButton { border: none; background: transparent; font-weight: 600; text-align: left; }"
        )
        buton_titlu.clicked.connect(lambda _, mid=mat["id"]: self._toggle_material(mid))
        antet_layout.addWidget(buton_titlu)
        antet_layout.addStretch(1)

        culoare_rest = CULOARE_EROARE if total_rest < 0 else CULOARE_SUCCES
        stat = QLabel(
            f"Intrare <b>{fmt(total_intrare)} kg</b> &nbsp;&nbsp; "
            f"Consum <b>{fmt(total_consum)} kg</b> &nbsp;&nbsp; "
            f"Rest <b style='color:{culoare_rest}'>{fmt(total_rest)} kg</b>"
        )
        stat.setTextFormat(Qt.RichText)
        stat.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 12px;")
        antet_layout.addWidget(stat)

        if self.poate_edita:
            btn_add = QPushButton("+ Lot nou")
            btn_add.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn_add.clicked.connect(lambda _, m=mat: self._adauga_lot(m))
            antet_layout.addWidget(btn_add)

        layout.addWidget(antet)
        if expanded:
            layout.addWidget(self._construieste_tabel_loturi(lots, mat["id"]))
        return cadru

    def _construieste_tabel_loturi(self, lots, material_id):
        coloane = ["Nr. LOT", "O %", "Fe %", "N %", "V %", "Al %", "Ti % (calc.)",
                   "Stoc intrare", "Consum", "Rest", "Rezervare", ""]

        tabel = QTableWidget(len(lots) if lots else 1, len(coloane))
        tabel.setHorizontalHeaderLabels(coloane)
        tabel.verticalHeader().setVisible(False)
        tabel.setEditTriggers(QTableWidget.NoEditTriggers)
        tabel.setSelectionMode(QTableWidget.NoSelection)
        tabel.setStyleSheet(
            f"QTableWidget {{ border: none; gridline-color: {CULOARE_BORDURA}; }}"
            f"QHeaderView::section {{ background-color: white; color: {CULOARE_GRI_TEXT}; "
            f"border: none; border-bottom: 1px solid {CULOARE_BORDURA}; padding: 6px; font-size: 10.5px; }}"
        )
        tabel.horizontalHeader().setStretchLastSection(True)

        if not lots:
            mesaj = "Niciun lot introdus inca pentru acest material."
            mesaj += " Foloseste \u201e+ Lot nou\u201d." if self.poate_edita else " Contul tau nu are drept de editare — contacteaza un utilizator cu drept de editare pentru a adauga loturi."
            item = QTableWidgetItem(mesaj)
            item.setTextAlignment(Qt.AlignCenter)
            tabel.setSpan(0, 0, 1, len(coloane))
            tabel.setItem(0, 0, item)
            tabel.setFixedHeight(60)
            return tabel

        for r, l in enumerate(lots):
            rest = lot_rest(l)
            ti_val = lot_ti(l)
            ti_display = f"{fmt(ti_val, 3)}%" if ti_val > 0 else "0%"

            valori = [
                l.get("nrLot", ""), fmt(l.get("dozaO"), 3), fmt(l.get("dozaFe"), 3),
                fmt(l.get("dozaN"), 3), fmt(l.get("dozaV"), 3), fmt(l.get("dozaAl"), 3),
                ti_display, fmt(l.get("stocIntrare"), 2), fmt(l.get("consum"), 2),
                f"{fmt(rest, 2)} kg",
            ]
            for c, val in enumerate(valori):
                item = QTableWidgetItem(str(val))
                if c > 0:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if c == 9:
                    if rest < 0:
                        item.setForeground(QColor(CULOARE_EROARE))
                    elif rest < to_float(l.get("stocIntrare")) * 0.1:
                        item.setForeground(QColor(CULOARE_AVERTISMENT))
                    else:
                        item.setForeground(QColor(CULOARE_SUCCES))
                tabel.setItem(r, c, item)

            rezervat_nume = l.get("rezervatComandaNume") if l.get("rezervatComandaId") else None
            item_rezervare = QTableWidgetItem(f"\U0001F512 {rezervat_nume}" if rezervat_nume else "")
            if rezervat_nume:
                item_rezervare.setForeground(QColor(CULOARE_AVERTISMENT))
                font_rezervare = item_rezervare.font()
                font_rezervare.setBold(True)
                item_rezervare.setFont(font_rezervare)
            tabel.setItem(r, 10, item_rezervare)

            cell_actiuni = QWidget()
            cell_layout = QHBoxLayout(cell_actiuni)
            cell_layout.setContentsMargins(0, 0, 0, 0)
            cell_layout.setSpacing(4)
            btn_istoric = QPushButton("Istoric")
            btn_istoric.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_istoric.setCursor(Qt.PointingHandCursor)
            btn_istoric.setToolTip("Vezi in ce comenzi/retete/bare a fost folosit acest lot")
            btn_istoric.clicked.connect(lambda _, lid=l["id"]: self._arata_istoric_lot(lid))
            cell_layout.addWidget(btn_istoric)
            if self.poate_edita:
                btn_edit = QPushButton("Editeaza")
                btn_edit.setStyleSheet(STIL_BUTON_SECUNDAR)
                btn_edit.setCursor(Qt.PointingHandCursor)
                btn_edit.clicked.connect(lambda _, lid=l["id"], mat_id=material_id: self._editeaza_lot(lid, mat_id))
                cell_layout.addWidget(btn_edit)
                btn = QPushButton("\u2715")
                btn.setStyleSheet(STIL_BUTON_PERICOL)
                btn.setCursor(Qt.PointingHandCursor)
                btn.clicked.connect(lambda _, lid=l["id"]: self._sterge_lot(lid))
                cell_layout.addWidget(btn)
            tabel.setCellWidget(r, len(coloane) - 1, cell_actiuni)

        tabel.resizeRowsToContents()
        tabel.setFixedHeight(min(320, 44 + 34 * len(lots)))
        return tabel

    def _toggle_material(self, mat_id):
        self.mat_expanded[mat_id] = not self.mat_expanded.get(mat_id, True)
        self._rebuild_stoc()

    def _adauga_lot(self, mat):
        dialog = DialogLotNou(mat["nume"], self, orders=self.state["orders"])
        if dialog.exec() == QDialog.Accepted and dialog.rezultat:
            lot = {"id": uid(), "material": mat["id"]}
            lot.update(dialog.rezultat)
            self.state["lots"].append(lot)
            salveaza_date(self.state)
            self._rebuild_stoc()
            self._rebuild_retete()

    def _editeaza_lot(self, lot_id, material_id):
        lot = next((l for l in self.state["lots"] if l["id"] == lot_id), None)
        if lot is None:
            return
        mat = next((m for m in MATERIALE if m["id"] == material_id), None)
        mat_nume = mat["nume"] if mat else ""
        dialog = DialogLotNou(mat_nume, self, lot=lot, orders=self.state["orders"])
        if dialog.exec() == QDialog.Accepted and dialog.rezultat:
            lot.update(dialog.rezultat)
            salveaza_date(self.state)
            self._rebuild_stoc()
            self._rebuild_retete()

    def _arata_istoric_lot(self, lot_id):
        lot = next((l for l in self.state["lots"] if l["id"] == lot_id), None)
        if lot is None:
            return
        mat = next((m for m in MATERIALE if m["id"] == lot.get("material")), None)
        mat_nume = mat["nume"] if mat else ""
        istoric = gaseste_istoric_lot(lot_id, self.state["orders"])
        dialog = DialogIstoricLot(lot, mat_nume, istoric, self)
        dialog.exec()

    def _sterge_lot(self, lot_id):
        if QMessageBox.question(self, "Confirmare",
                                 "Stergi acest lot din stoc? Actiunea nu poate fi anulata.") != QMessageBox.Yes:
            return
        self.state["lots"] = [l for l in self.state["lots"] if l["id"] != lot_id]
        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()

    # ------------------------ TAB: COMENZI SI RETETE ------------------------
    def _rebuild_retete(self):
        # Migrare pentru comenzile vechi, salvate inainte ca barele sa fie
        # mutate la nivel de comanda: totalul se deduce din barele deja
        # existente pe retete, ca sa nu se piarda nimic.
        for order in self.state.get("orders", []):
            if not str(order.get("nrBare", "")).strip():
                existente = sum(len(r.get("bare", [])) for r in order.get("retete", []))
                order["nrBare"] = str(existente) if existente else ""

        _clear_layout(self.retete_layout)

        antet_widget = QWidget()
        antet = QHBoxLayout(antet_widget)
        antet.setContentsMargins(0, 0, 0, 0)
        text = QVBoxLayout()
        titlu = QLabel("Comenzi")
        titlu.setFont(QFont("Sans Serif", 13, QFont.Bold))
        subtitlu = QLabel("O comanda poate contine mai multe retete; o reteta poate contine mai multe bare.")
        subtitlu.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11.5px;")
        subtitlu.setWordWrap(True)
        text.addWidget(titlu)
        text.addWidget(subtitlu)
        antet.addLayout(text)
        antet.addStretch(1)
        if self.poate_edita:
            btn_comanda = QPushButton("+ Comanda noua")
            btn_comanda.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn_comanda.clicked.connect(self._adauga_comanda)
            antet.addWidget(btn_comanda)
        self.retete_layout.addWidget(antet_widget)

        if not self.state["orders"]:
            gol = QLabel("Nicio comanda inca. Creeaza o comanda pentru a incepe sa definesti retete de sarja.")
            gol.setAlignment(Qt.AlignCenter)
            gol.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; padding: 30px;")
            self.retete_layout.addWidget(gol)
        else:
            for order in self.state["orders"]:
                self.retete_layout.addWidget(self._construieste_card_comanda(order))

        self.retete_layout.addStretch(1)

    def _adauga_comanda(self):
        nume, ok = QInputDialog.getText(self, "Comanda noua", "Nume comanda / client:")
        if not ok or not nume.strip():
            return
        order = {"id": uid(), "nume": nume.strip(), "data": date.today().strftime("%d.%m.%Y"),
                 "tipAliaj": ALIAJ_IMPLICIT, "nrBare": "", "retete": []}
        self.state["orders"].append(order)
        self.order_expanded[order["id"]] = True
        salveaza_date(self.state)
        self._rebuild_retete()

    def _sterge_comanda(self, order_id):
        if QMessageBox.question(self, "Confirmare",
                                 "Stergi aceasta comanda si toate retetele ei?") != QMessageBox.Yes:
            return
        self.state["orders"] = [o for o in self.state["orders"] if o["id"] != order_id]
        salveaza_date(self.state)
        self._rebuild_retete()

    def _genereaza_fisa_limita(self, order_id):
        order = self._gaseste_comanda(order_id)

        if not OPENPYXL_DISPONIBIL:
            QMessageBox.critical(
                self, "Lipseste o biblioteca",
                "Generarea fisei limita necesita biblioteca 'openpyxl'.\n\n"
                "Instaleaz-o cu: pip install openpyxl\nsi reporneste aplicatia."
            )
            return

        consum = calculeaza_consum_comanda(order, self.state["lots"])
        are_consum = any(date_mat["total"] > 1e-9 for date_mat in consum.values())
        if not are_consum:
            raspuns = QMessageBox.question(
                self, "Niciun consum aplicat",
                "Aceasta comanda nu are inca niciun consum aplicat pe stoc (nicio bara cu "
                "\u201eAplica consumul pe stoc\u201d), asa ca fisa limita va iesi cu coloanele "
                "LOT si Cantitate goale \u2014 cantitatile sunt luate STRICT din ce s-a consumat "
                "efectiv, nu din dozarea teoretica.\n\nContinui oricum?"
            )
            if raspuns != QMessageBox.Yes:
                return

        nume_implicit = f"Fisa_limita_{order['nume']}.xlsx"
        nume_implicit = "".join(c for c in nume_implicit if c not in '\\/:*?"<>|')
        cale, _ = QFileDialog.getSaveFileName(
            self, "Salveaza fisa limita", nume_implicit, "Excel (*.xlsx)"
        )
        if not cale:
            return
        if not cale.lower().endswith(".xlsx"):
            cale += ".xlsx"

        try:
            genereaza_fisa_limita_xlsx(order, consum, cale, intocmit_nume=self.utilizator.get("alias"))
        except Exception as e:
            QMessageBox.critical(self, "Eroare la generare", str(e))
            return

        cale_pdf = os.path.splitext(cale)[0] + ".pdf"
        pdf_ok, pdf_eroare = True, ""
        if REPORTLAB_DISPONIBIL:
            try:
                genereaza_fisa_limita_pdf(order, consum, cale_pdf, intocmit_nume=self.utilizator.get("alias"))
            except Exception as e:
                pdf_ok, pdf_eroare = False, str(e)
        else:
            pdf_ok = False
            pdf_eroare = "biblioteca 'reportlab' nu este instalata (pip install reportlab)"

        mesaj = f"Fisa limita a fost salvata:\n{cale}"
        if pdf_ok:
            mesaj += f"\n\nRaportul PDF a fost salvat:\n{cale_pdf}"
        else:
            mesaj += f"\n\nRaportul PDF NU a putut fi generat ({pdf_eroare})."
        QMessageBox.information(self, "Succes", mesaj)

    def _seteaza_tip_aliaj(self, order_id, tip_aliaj):
        order = self._gaseste_comanda(order_id)
        if order.get("tipAliaj") == tip_aliaj:
            return
        order["tipAliaj"] = tip_aliaj
        salveaza_date(self.state)
        self._rebuild_retete()

    def _seteaza_beneficiar(self, order_id, text):
        order = self._gaseste_comanda(order_id)
        text = text.strip()
        if order.get("beneficiar", BENEFICIAR_IMPLICIT) == text:
            return
        order["beneficiar"] = text
        salveaza_date(self.state)

    def _seteaza_nr_buc_lingouri(self, order_id, text):
        order = self._gaseste_comanda(order_id)
        text = text.strip()
        if str(order.get("nrBucLingouri", "")) == text:
            return
        order["nrBucLingouri"] = text
        salveaza_date(self.state)

    def _genereaza_retdozare(self, order_id):
        order = self._gaseste_comanda(order_id)
        tip_aliaj = order.get("tipAliaj", ALIAJ_IMPLICIT)

        if tip_aliaj not in ALIAJE_SPEC:
            QMessageBox.warning(
                self, "Aliaj neimplementat",
                f"Generarea RetDozare nu este inca implementata pentru tipul de aliaj \u201e{tip_aliaj}\u201d."
            )
            return

        if not OPENPYXL_DISPONIBIL:
            QMessageBox.critical(
                self, "Lipseste o biblioteca",
                "Generarea RetDozare necesita biblioteca 'openpyxl'.\n\n"
                "Instaleaz-o cu: pip install openpyxl\nsi reporneste aplicatia."
            )
            return

        if not order["retete"]:
            QMessageBox.warning(self, "Nicio reteta", "Aceasta comanda nu are inca nicio reteta definita.")
            return

        nume_implicit = f"RetDozare_{order['nume']}.xlsx"
        nume_implicit = "".join(c for c in nume_implicit if c not in '\\/:*?"<>|')
        cale, _ = QFileDialog.getSaveFileName(
            self, "Salveaza RetDozare", nume_implicit, "Excel (*.xlsx)"
        )
        if not cale:
            return
        if not cale.lower().endswith(".xlsx"):
            cale += ".xlsx"

        try:
            genereaza_retdozare_xlsx(order, self.state["lots"], cale, intocmit_nume=self.utilizator.get("alias"))
        except Exception as e:
            QMessageBox.critical(self, "Eroare la generare", str(e))
            return

        cale_pdf = os.path.splitext(cale)[0] + ".pdf"
        pdf_ok, pdf_eroare = True, ""
        if REPORTLAB_DISPONIBIL:
            try:
                genereaza_retdozare_pdf(order, self.state["lots"], cale_pdf, intocmit_nume=self.utilizator.get("alias"))
            except Exception as e:
                pdf_ok, pdf_eroare = False, str(e)
        else:
            pdf_ok = False
            pdf_eroare = "biblioteca 'reportlab' nu este instalata (pip install reportlab)"

        mesaj = f"RetDozare a fost salvat:\n{cale}"
        if pdf_ok:
            mesaj += f"\n\nRaportul PDF (o pagina per reteta) a fost salvat:\n{cale_pdf}"
        else:
            mesaj += f"\n\nRaportul PDF NU a putut fi generat ({pdf_eroare})."
        QMessageBox.information(self, "Succes", mesaj)

    def _toggle_comanda(self, order_id):
        self.order_expanded[order_id] = not self.order_expanded.get(order_id, False)
        self._rebuild_retete()

    def _gaseste_comanda(self, order_id):
        return next(o for o in self.state["orders"] if o["id"] == order_id)

    def _gaseste_reteta(self, order_id, recipe_id):
        order = self._gaseste_comanda(order_id)
        return next(r for r in order["retete"] if r["id"] == recipe_id)

    def _construieste_card_comanda(self, order):
        expanded = self.order_expanded.get(order["id"], False)
        cadru = QFrame()
        cadru.setStyleSheet(f"QFrame {{ border: 1px solid {CULOARE_BORDURA}; border-radius: 8px; }}")
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        antet = QFrame()
        antet.setStyleSheet(f"background-color: {CULOARE_FUNDAL_SECTIUNE};")
        antet_layout = QHBoxLayout(antet)
        antet_layout.setContentsMargins(14, 10, 14, 10)

        buton_titlu = QPushButton(("\u25be " if expanded else "\u25b8 ") + order["nume"])
        buton_titlu.setFlat(True)
        buton_titlu.setCursor(Qt.PointingHandCursor)
        buton_titlu.setStyleSheet(
            "QPushButton { border: none; background: transparent; font-weight: 600; text-align: left; }"
        )
        buton_titlu.clicked.connect(lambda _, oid=order["id"]: self._toggle_comanda(oid))
        antet_layout.addWidget(buton_titlu)

        combo_aliaj = QComboBox()
        combo_aliaj.setStyleSheet(STIL_CAMP)
        combo_aliaj.setToolTip("Tipul de aliaj al comenzii \u2014 determina limitele chimice si formatul RetDozare folosite")
        tip_curent = order.get("tipAliaj", ALIAJ_IMPLICIT)
        index_sel = 0
        for i, tip in enumerate(ALIAJE_DISPONIBILE):
            combo_aliaj.addItem(tip)
            if tip == tip_curent:
                index_sel = i
        combo_aliaj.setCurrentIndex(index_sel)
        combo_aliaj.currentTextChanged.connect(
            lambda text, oid=order["id"]: self._seteaza_tip_aliaj(oid, text)
        )
        combo_aliaj.setEnabled(self.poate_edita)
        antet_layout.addWidget(combo_aliaj)

        camp_beneficiar = QLineEdit(order.get("beneficiar", BENEFICIAR_IMPLICIT))
        camp_beneficiar.setPlaceholderText("Beneficiar")
        camp_beneficiar.setToolTip("Beneficiarul comenzii \u2014 apare in antetul \u201eFisa limita\u201d (ex. \u201eBeneficiar-Zirom\u201d)")
        camp_beneficiar.setFixedWidth(110)
        camp_beneficiar.setStyleSheet(STIL_CAMP)
        camp_beneficiar.editingFinished.connect(
            lambda oid=order["id"], c=camp_beneficiar: self._seteaza_beneficiar(oid, c.text())
        )
        camp_beneficiar.setReadOnly(not self.poate_edita)
        antet_layout.addWidget(camp_beneficiar)

        camp_buc_lingou = QLineEdit(str(order.get("nrBucLingouri", "")))
        camp_buc_lingou.setPlaceholderText("buc. lingou")
        camp_buc_lingou.setToolTip("Numarul de bucati lingou al comenzii \u2014 apare in antetul \u201eFisa limita\u201d (ex. \u201e3 buc lingou\u201d)")
        camp_buc_lingou.setFixedWidth(80)
        camp_buc_lingou.setStyleSheet(STIL_CAMP)
        camp_buc_lingou.editingFinished.connect(
            lambda oid=order["id"], c=camp_buc_lingou: self._seteaza_nr_buc_lingouri(oid, c.text())
        )
        camp_buc_lingou.setReadOnly(not self.poate_edita)
        antet_layout.addWidget(camp_buc_lingou)

        camp_nr_bare_cda = QLineEdit(str(order.get("nrBare", "") or ""))
        camp_nr_bare_cda.setPlaceholderText("bare")
        camp_nr_bare_cda.setToolTip(
            "Numarul TOTAL de bare al comenzii. Barele de pe retete NU se mai scriu de mana: "
            "se distribuie automat, in functie de cate incap in loturile selectate pe fiecare reteta."
        )
        camp_nr_bare_cda.setFixedWidth(70)
        camp_nr_bare_cda.setStyleSheet(STIL_CAMP)
        camp_nr_bare_cda.editingFinished.connect(
            lambda oid=order["id"], c=camp_nr_bare_cda: self._seteaza_nr_bare_comanda(oid, c.text())
        )
        camp_nr_bare_cda.setReadOnly(not self.poate_edita)
        antet_layout.addWidget(camp_nr_bare_cda)

        if self.poate_edita:
            btn_distribuie = QPushButton("\u21bb Distribuie barele")
            btn_distribuie.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_distribuie.setToolTip(
                "Recalculeaza cate bare incap pe fiecare reteta, in functie de loturile "
                "selectate, si creeaza automat retetele de continuare de care e nevoie."
            )
            btn_distribuie.clicked.connect(
                lambda _, oid=order["id"]: self._distribuie_bare_comanda(oid)
            )
            antet_layout.addWidget(btn_distribuie)

        n = len(order["retete"])
        asignate = sum(len(r["bare"]) for r in order["retete"])
        meta = QLabel(f"{n} reteta{'e' if n != 1 else ''} \u00b7 {asignate} bare plasate \u00b7 creata {order['data']}")
        meta.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
        antet_layout.addWidget(meta)
        antet_layout.addStretch(1)

        btn_document = QPushButton("\U0001F4C4 Fisa limita")
        btn_document.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_document.setToolTip("Genereaza \"Fisa limita -cda\" (.xlsx + raport .pdf) cu consumul REAL de materiale")
        btn_document.clicked.connect(lambda _, oid=order["id"]: self._genereaza_fisa_limita(oid))
        antet_layout.addWidget(btn_document)

        if order.get("tipAliaj", ALIAJ_IMPLICIT) in ALIAJE_SPEC:
            btn_retdozare = QPushButton("\U0001F4C4 RetDozare")
            btn_retdozare.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_retdozare.setToolTip(
                "Genereaza \"RetDozare\" (.xlsx + .pdf cu o pagina per reteta), "
                f"conform formatului pentru {ALIAJE_SPEC[order.get('tipAliaj', ALIAJ_IMPLICIT)]['nume']}"
            )
            btn_retdozare.clicked.connect(lambda _, oid=order["id"]: self._genereaza_retdozare(oid))
            antet_layout.addWidget(btn_retdozare)

        if self.poate_edita:
            btn_sterge = QPushButton("\u2715")
            btn_sterge.setStyleSheet(STIL_BUTON_PERICOL)
            btn_sterge.clicked.connect(lambda _, oid=order["id"]: self._sterge_comanda(oid))
            antet_layout.addWidget(btn_sterge)

        layout.addWidget(antet)

        if expanded:
            corp = QFrame()
            corp_layout = QVBoxLayout(corp)
            corp_layout.setContentsMargins(14, 12, 14, 12)
            corp_layout.setSpacing(12)
            for r in order["retete"]:
                corp_layout.addWidget(self._construieste_card_reteta(order, r))
            if self.poate_edita:
                btn_reteta = QPushButton("+ Reteta noua")
                btn_reteta.setStyleSheet(STIL_BUTON_SECUNDAR)
                btn_reteta.clicked.connect(lambda _, oid=order["id"]: self._adauga_reteta(oid))
                corp_layout.addWidget(btn_reteta)
            layout.addWidget(corp)

        return cadru

    def _adauga_reteta(self, order_id):
        order = self._gaseste_comanda(order_id)
        order["retete"].append({
            "id": uid(),
            "nume": f"Reteta {len(order['retete']) + 1}",
            "target": {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.17},
            "lotSel": {},
            "bare": [],
            "numarPresari": "1",
        })
        salveaza_date(self.state)
        self._rebuild_retete()

    def _sterge_reteta(self, order_id, recipe_id):
        if QMessageBox.question(self, "Confirmare", "Stergi aceasta reteta?") != QMessageBox.Yes:
            return
        order = self._gaseste_comanda(order_id)
        order["retete"] = [r for r in order["retete"] if r["id"] != recipe_id]
        salveaza_date(self.state)
        self._rebuild_retete()

    def _construieste_card_reteta(self, order, r):
        cadru = QFrame()
        cadru.setStyleSheet(
            f"QFrame {{ background-color: {CULOARE_FUNDAL_SECTIUNE}; "
            f"border: 1px solid {CULOARE_BORDURA}; border-radius: 8px; }}"
        )
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        rand_titlu = QHBoxLayout()
        camp_nume = QLineEdit(r["nume"])
        camp_nume.setStyleSheet(STIL_CAMP)
        camp_nume.editingFinished.connect(
            lambda oid=order["id"], rid=r["id"], c=camp_nume: self._redenumeste_reteta(oid, rid, c.text())
        )
        camp_nume.setReadOnly(not self.poate_edita)
        rand_titlu.addWidget(camp_nume)
        if self.poate_edita:
            btn_sterge = QPushButton("\u2715")
            btn_sterge.setStyleSheet(STIL_BUTON_PERICOL)
            btn_sterge.clicked.connect(lambda _, oid=order["id"], rid=r["id"]: self._sterge_reteta(oid, rid))
            rand_titlu.addWidget(btn_sterge)
        layout.addLayout(rand_titlu)

        layout.addWidget(self._eticheta_eyebrow("Compozitie chimica tinta in bara (%)"))
        layout.addWidget(QLabel("Modifica valorile pentru a stabili compozitia dorita. Ti% se calculeaza automat ca rest."))

        rand_tinta = QHBoxLayout()
        for camp, eticheta in [("al", "Al %"), ("v", "V %"), ("o", "O %"), ("fe", "Fe %") , ("mo", "Mo %"), ("si", "Si %"), ("zr", "Zr %")]:
            bloc = QVBoxLayout()
            bloc.addWidget(self._eticheta_mica(eticheta))
            camp_edit = QLineEdit(str(r["target"].get(camp, 0)))
            camp_edit.setStyleSheet(STIL_CAMP)
            camp_edit.editingFinished.connect(
                lambda oid=order["id"], rid=r["id"], f=camp, c=camp_edit:
                    self._seteaza_tinta(oid, rid, f, c.text())
            )
            camp_edit.setReadOnly(not self.poate_edita)
            bloc.addWidget(camp_edit)
            rand_tinta.addLayout(bloc)

        bloc_ti = QVBoxLayout()
        bloc_ti.addWidget(self._eticheta_mica("Ti % (rest)"))
        camp_ti = QLineEdit(fmt(target_ti(r["target"]), 2))
        camp_ti.setEnabled(False)
        camp_ti.setStyleSheet(STIL_CAMP)
        bloc_ti.addWidget(camp_ti)
        rand_tinta.addLayout(bloc_ti)
        layout.addLayout(rand_tinta)

        layout.addWidget(self._eticheta_eyebrow("Loturi folosite in amestec"))
        rand_loturi = QHBoxLayout()
        for mat in MATERIALE:
            necesar = material_necesar(mat["id"], r["target"])
            bloc = QVBoxLayout()
            eticheta_mat = mat["nume"] if necesar else f"{mat['nume']} (neutilizat, tinta 0%)"
            eticheta_widget = self._eticheta_mica(eticheta_mat)
            if not necesar:
                eticheta_widget.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10px; font-style: italic;")
            bloc.addWidget(eticheta_widget)
            combo = ComboFaraScroll()
            combo.setStyleSheet(STIL_CAMP)
            combo.setEnabled(necesar and self.poate_edita)
            combo.addItem("\u2014 alege lot \u2014" if necesar else "\u2014 neutilizat \u2014", "")
            lots_mat = [l for l in self.state["lots"] if l["material"] == mat["id"]]
            sel_actual = r["lotSel"].get(mat["id"], "")
            index_sel = 0
            for i, l in enumerate(lots_mat, start=1):
                text = f"{l['nrLot']} (rest {fmt(lot_rest(l), 2)} kg)"
                if l.get("rezervatComandaId") and l["rezervatComandaId"] != order["id"]:
                    text += f" \u26a0 REZERVAT: {l.get('rezervatComandaNume', '')}"
                combo.addItem(text, l["id"])
                if l["id"] == sel_actual:
                    index_sel = i
            combo.setCurrentIndex(index_sel)
            combo.currentIndexChanged.connect(
                lambda _idx, oid=order["id"], rid=r["id"], m=mat["id"], cb=combo:
                    self._seteaza_lot(oid, rid, m, cb.currentData())
            )
            bloc.addWidget(combo)
            rand_loturi.addLayout(bloc)
        layout.addLayout(rand_loturi)

        rand_multiplicatori = QHBoxLayout()
        bloc_portie = QVBoxLayout()
        bloc_portie.addWidget(self._eticheta_mica("Portie (kg) - aceeasi dozare pentru toate barele"))
        camp_portie = QLineEdit(str(r.get("portie") or ""))
        camp_portie.setStyleSheet(STIL_CAMP)
        camp_portie.setFixedWidth(100)
        camp_portie.editingFinished.connect(
            lambda oid=order["id"], rid=r["id"], c=camp_portie: self._seteaza_portie_reteta(oid, rid, c.text())
        )
        camp_portie.setReadOnly(not self.poate_edita)
        bloc_portie.addWidget(camp_portie)
        rand_multiplicatori.addLayout(bloc_portie)

        bloc_nr_bare = QVBoxLayout()
        bloc_nr_bare.addWidget(self._eticheta_mica("Bare pe reteta (automat)"))
        camp_nr_bare = QLineEdit(str(len(r["bare"])))
        camp_nr_bare.setStyleSheet(STIL_CAMP)
        camp_nr_bare.setFixedWidth(90)
        # Numarul de bare al retetei NU se mai introduce manual: el rezulta din
        # cate bare incap in loturile selectate aici. Totalul se scrie pe comanda.
        camp_nr_bare.setReadOnly(True)
        camp_nr_bare.setToolTip(
            "Rezulta automat din loturile selectate. Numarul total de bare se "
            "seteaza pe comanda, apoi apesi \u201eDistribuie barele\u201d."
        )
        bloc_nr_bare.addWidget(camp_nr_bare)
        rand_multiplicatori.addLayout(bloc_nr_bare)

        bloc_nr_presari = QVBoxLayout()
        bloc_nr_presari.addWidget(self._eticheta_mica("Numar de presari (acelasi pentru toate barele)"))
        camp_nr_presari = QLineEdit(str(r.get("numarPresari") or "1"))
        camp_nr_presari.setStyleSheet(STIL_CAMP)
        camp_nr_presari.setFixedWidth(90)
        camp_nr_presari.editingFinished.connect(
            lambda oid=order["id"], rid=r["id"], c=camp_nr_presari: self._seteaza_numar_presari(oid, rid, c.text())
        )
        camp_nr_presari.setReadOnly(not self.poate_edita)
        bloc_nr_presari.addWidget(camp_nr_presari)
        rand_multiplicatori.addLayout(bloc_nr_presari)
        rand_multiplicatori.addStretch(1)
        layout.addLayout(rand_multiplicatori)

        btn_capacitate = QPushButton("Capacitate reteta / desfasurare pe bare")
        btn_capacitate.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_capacitate.setToolTip(
            "Arata, bara cu bara, cat consuma fiecare si cat mai ramane din fiecare lot. "
            "De aici se aplica apoi consumul o singura data, pentru toata reteta."
        )
        btn_capacitate.clicked.connect(
            lambda _, oid=order["id"], rid=r["id"]: self._capacitate_reteta(oid, rid)
        )
        rand_capacitate = QHBoxLayout()
        rand_capacitate.addWidget(btn_capacitate)

        if self.poate_edita and self._are_consum_de_retras(order, r):
            btn_retrage = QPushButton("\u21a9 Retrage consumul retetei")
            btn_retrage.setStyleSheet(STIL_BUTON_PERICOL)
            btn_retrage.setToolTip(
                "Anuleaza dintr-o data tot consumul aplicat pe stoc de aceasta reteta "
                "(toate barele ei, inclusiv bara de report, daca exista) si repune "
                "cantitatile in loturi."
            )
            btn_retrage.clicked.connect(
                lambda _, oid=order["id"], rid=r["id"]: self._retrage_consum_total_reteta(oid, rid)
            )
            rand_capacitate.addWidget(btn_retrage)
        layout.addLayout(rand_capacitate)

        layout.addWidget(self._eticheta_eyebrow("Bare - urmarire consum pe stoc (toate folosesc aceeasi portie si numar de presari)"))
        offset_bare = _numar_bare_anterioare(order, r["id"])
        for i, bar in enumerate(r["bare"]):
            layout.addWidget(self._construieste_rand_bara(order, r, bar, offset_bare + i))

        layout.addWidget(self._construieste_bilant_reteta(order, r))

        return cadru

    def _construieste_bilant_reteta(self, order, r):
        cadru = QFrame()
        cadru.setStyleSheet(f"QFrame {{ background-color: {CULOARE_FUNDAL_SECTIUNE}; border-radius: 7px; }}")
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        layout.addWidget(self._eticheta_eyebrow("Bilant reteta (suma tuturor barelor, cu presarile lor)"))

        all_selected = all(r["lotSel"].get(k) for k in ORDINE_MAT if material_necesar(k, r["target"]))
        if not all_selected:
            layout.addWidget(self._banner("Selecteaza un lot pentru fiecare material necesar (cu tinta > 0%) ca sa vezi bilantul retetei.", "info"))
            return cadru
        if not r["bare"]:
            layout.addWidget(self._banner("Adauga cel putin o bara pentru a vedea bilantul retetei.", "info"))
            return cadru

        tip_aliaj = order.get("tipAliaj", ALIAJ_IMPLICIT)
        lot_sel = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        totaluri = {k: 0.0 for k in ORDINE_MAT}
        materiale_folosite = set()
        erori = []
        for bar in r["bare"]:
            if bar.get("calcSnapshot"):
                calc = bar["calcSnapshot"]
            else:
                calc = calculeaza_bara(tip_aliaj, r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
            if calc.get("eroare"):
                erori.append(calc["eroare"])
                continue
            materiale_folosite.update(calc["rezultatTotal"].keys())
            for k in ORDINE_MAT:
                totaluri[k] += calc["rezultatTotal"].get(k, 0.0)

        if erori:
            layout.addWidget(self._banner(erori[0], "err"))

        grid = QHBoxLayout()
        depaseste_stoc = False
        for mat in MATERIALE:
            if materiale_folosite and mat["id"] not in materiale_folosite:
                # Material care nu face parte din formula acestui tip de
                # aliaj (ex. Aliaj AlV la o reteta VT9) — nu se afiseaza.
                continue
            val = totaluri[mat["id"]]
            lot = lot_sel.get(mat["id"])
            rest = lot_rest(lot) if lot else 0
            culoare = "black"
            if val < -0.005:
                culoare = CULOARE_EROARE
            elif lot and val > rest:
                culoare = CULOARE_AVERTISMENT
                depaseste_stoc = True
            cell = QFrame()
            cell.setStyleSheet(f"background: white; border: 1px solid {CULOARE_BORDURA}; border-radius: 6px;")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(8, 6, 8, 6)
            eticheta = QLabel(mat["nume"])
            eticheta.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10px;")
            valoare = QLabel(f"{fmt(val)} kg")
            valoare.setStyleSheet(f"font-weight: 700; color: {culoare};")
            cell_layout.addWidget(eticheta)
            cell_layout.addWidget(valoare)
            if lot:
                rest_label = QLabel(f"rest: {fmt(rest)} kg")
                rest_label.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 9px;")
                cell_layout.addWidget(rest_label)
            grid.addWidget(cell)
        layout.addLayout(grid)

        eticheta_total = QLabel(f"Total materiale pentru intreaga reteta: {fmt(sum(totaluri.values()))} kg")
        eticheta_total.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10.5px;")
        layout.addWidget(eticheta_total)

        if depaseste_stoc:
            layout.addWidget(self._banner("⚠️ Bilantul retetei depaseste stocul disponibil pentru cel putin un material.", "err"))

        return cadru

    def _eticheta_eyebrow(self, text):
        l = QLabel(text.upper())
        l.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10px; font-weight: 700;")
        return l

    def _eticheta_mica(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10.5px;")
        return l

    def _redenumeste_reteta(self, order_id, recipe_id, text):
        r = self._gaseste_reteta(order_id, recipe_id)
        if text.strip():
            r["nume"] = text.strip()
            salveaza_date(self.state)

    def _seteaza_tinta(self, order_id, recipe_id, camp, text):
        r = self._gaseste_reteta(order_id, recipe_id)
        r["target"][camp] = to_float(text)
        salveaza_date(self.state)
        self._rebuild_retete()

    def _seteaza_lot(self, order_id, recipe_id, mat_id, lot_id):
        r = self._gaseste_reteta(order_id, recipe_id)
        if lot_id:
            lot = next((l for l in self.state["lots"] if l["id"] == lot_id), None)
            if lot and lot.get("rezervatComandaId") and lot["rezervatComandaId"] != order_id:
                QMessageBox.warning(
                    self, "Lot rezervat",
                    f"LOT REZERVAT PENTRU COMANDA {lot.get('rezervatComandaNume', '')}\n\n"
                    f"Lotul {lot.get('nrLot', '')} este rezervat exclusiv pentru comanda "
                    f"\u201e{lot.get('rezervatComandaNume', '')}\u201d si nu poate fi folosit in alta comanda. "
                    "Cere administratorului sa elibereze rezervarea daca vrei sa il folosesti aici."
                )
                self._rebuild_retete()  # revine la selectia anterioara, salvata
                return
        r["lotSel"][mat_id] = lot_id or None
        salveaza_date(self.state)
        self._rebuild_retete()

    def _seteaza_portie_reteta(self, order_id, recipe_id, text):
        r = self._gaseste_reteta(order_id, recipe_id)
        r["portie"] = text
        salveaza_date(self.state)
        self._rebuild_retete()

    def _seteaza_numar_presari(self, order_id, recipe_id, text):
        r = self._gaseste_reteta(order_id, recipe_id)
        r["numarPresari"] = text
        salveaza_date(self.state)
        self._rebuild_retete()

    def _construieste_rand_bara(self, order, r, bar, index):
        all_selected = all(r["lotSel"].get(k) for k in ORDINE_MAT if material_necesar(k, r["target"]))
        tip_aliaj = order.get("tipAliaj", ALIAJ_IMPLICIT)

        cadru = QFrame()
        cadru.setStyleSheet(f"QFrame {{ border: 1px solid {CULOARE_BORDURA}; border-radius: 7px; }}")
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        rand_sus = QHBoxLayout()
        eticheta_bara = QLabel(f"Bara #{index + 1}")
        eticheta_bara.setStyleSheet("font-weight: 600;")
        rand_sus.addWidget(eticheta_bara)

        # Bara de report vine din reteta anterioara si isi pastreaza dozarea
        # de acolo — trebuie sa se vada, ca sa nu fie confundata cu una noua.
        if bar.get("reportRamas"):
            sursa = bar.get("retetaSursaDozare") or "reteta anterioara"
            tag_report = QLabel(f"Report \u2014 dozare de pe {sursa}")
            tag_report.setStyleSheet(TAG_STYLES.get("warn", TAG_STYLES["good"]))
            tag_report.setToolTip(
                "Bara nu a incaput pe reteta anterioara. Pastreaza dozarea calculata "
                "acolo, consuma resturile ramase din loturile vechi si completeaza "
                "diferenta din loturile selectate aici."
            )
            rand_sus.addWidget(tag_report)
        elif bar.get("tipDozare") == "asteptare":
            tag_ast = QLabel("In asteptare \u2014 selecteaza loturile")
            tag_ast.setStyleSheet(TAG_STYLES.get("warn", TAG_STYLES["good"]))
            rand_sus.addWidget(tag_ast)

        rand_sus.addStretch(1)

        if bar.get("consumApplied"):
            tag = QLabel("Consum aplicat")
            tag.setStyleSheet(TAG_STYLES["good"])
            rand_sus.addWidget(tag)
            if self.poate_edita:
                btn_revert = QPushButton("Anuleaza consumul")
                btn_revert.setStyleSheet(STIL_BUTON_SECUNDAR)
                btn_revert.clicked.connect(
                    lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._revert_consum(oid, rid, bid)
                )
                rand_sus.addWidget(btn_revert)

        layout.addLayout(rand_sus)

        if not all_selected:
            layout.addWidget(self._banner("Selecteaza un lot pentru fiecare material necesar (cu tinta > 0%) ca sa poata fi calculat amestecul.", "info"))
        elif not r.get("portie") and not bar.get("calcSnapshot"):
            layout.addWidget(self._banner("Introdu portia (kg) a retetei pentru a calcula dozarea.", "info"))
        else:
            lot_sel = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
            if bar.get("calcSnapshot"):
                # Bara are un calcul inghetat (consum deja aplicat pe stoc SAU
                # bara a fost mutata pe alta reteta pastrandu-si dozarea
                # calculata initial) — nu-l recalculam dupa tinta/loturile
                # curente ale retetei in care se afla acum.
                calc = bar["calcSnapshot"]
            else:
                calc = calculeaza_bara(tip_aliaj, r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
            if calc.get("eroare"):
                layout.addWidget(self._banner(calc["eroare"], "err"))
            else:
                layout.addWidget(self._construieste_rezultate_bara(order, r, bar, calc, lot_sel))

        return cadru

    def _banner(self, text, tip):
        culori = {
            "err": (CULOARE_EROARE, "#fbeceb"),
            "ok": (CULOARE_SUCCES, "#e5f5ec"),
            "info": (CULOARE_BLEUMARIN, "#eef2f8"),
        }
        fg, bg = culori.get(tip, (CULOARE_GRI_TEXT, CULOARE_FUNDAL_SECTIUNE))
        l = QLabel(text)
        l.setWordWrap(True)
        l.setStyleSheet(f"color: {fg}; background-color: {bg}; border-radius: 6px; padding: 8px 10px; font-size: 12px;")
        return l

    def _construieste_rezultate_bara(self, order, r, bar, calc, lot_sel):
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        nr_presari = calc.get("nrPresari", 1)

        # Titlu rezultate (dozare pentru o singura bara)
        titlu_rezultate = QLabel("Cantitati necesare pentru compozitia tinta (per bara):")
        titlu_rezultate.setStyleSheet(f"font-weight: 600; color: {CULOARE_BLEUMARIN}; font-size: 11px;")
        layout.addWidget(titlu_rezultate)

        grid = QHBoxLayout()
        for mat in MATERIALE:
            if mat["id"] not in calc["rezultat"]:
                # Material care nu face parte din formula acestui tip de
                # aliaj (ex. Aliaj AlV la o reteta VT9) — nu se afiseaza.
                continue
            val = calc["rezultat"][mat["id"]]
            cell = QFrame()
            cell.setStyleSheet(f"background: white; border: 1px solid {CULOARE_BORDURA}; border-radius: 6px;")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(8, 6, 8, 6)
            eticheta = QLabel(mat["nume"])
            eticheta.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10px;")
            valoare = QLabel(f"{fmt(val)} kg")
            culoare = CULOARE_EROARE if val < -0.005 else "black"
            valoare.setStyleSheet(f"font-weight: 700; color: {culoare};")
            cell_layout.addWidget(eticheta)
            cell_layout.addWidget(valoare)
            grid.addWidget(cell)
        layout.addLayout(grid)

        # Total consumat din stoc (dozare x numar de bare din reteta)
        titlu_total = QLabel(f"Total consum din stoc \u2014 {fmt(nr_presari, 3)} bare:")
        titlu_total.setStyleSheet(f"font-weight: 600; color: {CULOARE_BLEUMARIN}; font-size: 11px; margin-top: 4px;")
        layout.addWidget(titlu_total)

        grid_total = QHBoxLayout()
        for mat in MATERIALE:
            if mat["id"] not in calc["rezultatTotal"]:
                continue
            val_total = calc["rezultatTotal"][mat["id"]]
            lot = lot_sel.get(mat["id"])
            rest = lot_rest(lot) if lot else 0
            culoare = "black"
            if val_total < -0.005:
                culoare = CULOARE_EROARE
            elif lot and val_total > rest:
                culoare = CULOARE_AVERTISMENT
            cell = QFrame()
            cell.setStyleSheet(f"background: white; border: 1px solid {CULOARE_BORDURA}; border-radius: 6px;")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(8, 6, 8, 6)
            eticheta = QLabel(mat["nume"])
            eticheta.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 10px;")
            valoare = QLabel(f"{fmt(val_total)} kg")
            valoare.setStyleSheet(f"font-weight: 700; color: {culoare};")
            cell_layout.addWidget(eticheta)
            cell_layout.addWidget(valoare)
            if lot:
                rest_label = QLabel(f"rest: {fmt(rest)} kg")
                rest_label.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 9px;")
                cell_layout.addWidget(rest_label)
            grid_total.addWidget(cell)
        layout.addLayout(grid_total)

        # Compozitia rezultata
        comp_rezultata = QLabel("Compozitia chimica rezultata in bara (verificare):")
        comp_rezultata.setStyleSheet(f"font-weight: 600; color: {CULOARE_BLEUMARIN}; font-size: 11px; margin-top: 4px;")
        layout.addWidget(comp_rezultata)

        comp_grid = QHBoxLayout()
        # Afisam compozitia rezultata din calc
        comp_rez = calc.get("compozitie_rezultata", {})
        for elem, eticheta in [("Al", "Al"), ("V", "V"), ("O", "O"), ("Fe", "Fe"), ("Ti", "Ti")]:
            if elem == "Ti":
                val = calc["tiRezultat"]
                tinta = calc["tiTarget"]
            else:
                val = comp_rez.get(elem, 0)
                tinta = to_float(r["target"].get(elem.lower(), 0))

            cell = QFrame()
            if abs(val - tinta) < 0.001:
                cell.setStyleSheet(f"background: #e5f5ec; border: 1px solid {CULOARE_SUCCES}; border-radius: 4px;")
            else:
                cell.setStyleSheet(f"background: #fbeceb; border: 1px solid {CULOARE_EROARE}; border-radius: 4px;")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(6, 4, 6, 4)
            etic = QLabel(eticheta)
            etic.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 9px;")
            valoare = QLabel(f"{fmt(val, 3)}%")
            valoare.setStyleSheet("font-weight: 700;")
            tinta_label = QLabel(f"(tinta: {fmt(tinta, 2)}%)")
            tinta_label.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 9px;")
            cell_layout.addWidget(etic)
            cell_layout.addWidget(valoare)
            cell_layout.addWidget(tinta_label)
            comp_grid.addWidget(cell)
        layout.addLayout(comp_grid)

        if calc["negativ"]:
            layout.addWidget(self._banner(
                "⚠️ Una sau mai multe cantitati depasesc stocul disponibil sau sunt negative. "
                "Poti aplica in continuare consumul — stocul va deveni negativ pentru materialele afectate.", "err"
            ))
            if not bar.get("consumApplied") and self.poate_edita:
                btn_muta = QPushButton("\u2192 Muta bara pe reteta urmatoare")
                btn_muta.setStyleSheet(STIL_BUTON_SECUNDAR)
                btn_muta.clicked.connect(
                    lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._muta_bara_pe_urmatoarea_reteta(oid, rid, bid)
                )
                layout.addWidget(btn_muta)

        if not bar.get("consumApplied"):
            if self.poate_edita:
                btn = QPushButton("✅ Aplica consumul pe stoc")
                btn.setStyleSheet(STIL_BUTON_PRINCIPAL)
                btn.clicked.connect(
                    lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._aplica_consum(oid, rid, bid)
                )
                layout.addWidget(btn)
            else:
                layout.addWidget(self._banner("Consumul nu a fost inca aplicat pe stoc. Contul tau nu are drept de editare.", "info"))
        else:
            layout.addWidget(self._banner("✅ Consumul a fost aplicat pe stoc.", "ok"))

        return container

    def _aplica_consum(self, order_id, recipe_id, bar_id):
        order = self._gaseste_comanda(order_id)
        r = self._gaseste_reteta(order_id, recipe_id)
        bar = next(b for b in r["bare"] if b["id"] == bar_id)
        lot_sel = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        calc = calculeaza_bara(order.get("tipAliaj", ALIAJ_IMPLICIT), r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
        if calc.get("eroare"):
            QMessageBox.warning(self, "Eroare", calc["eroare"])
            return
        snapshot = {}
        for k in ORDINE_MAT:
            lot = lot_sel[k]
            if lot is None:
                if material_necesar(k, r["target"]):
                    QMessageBox.warning(self, "Eroare", f"Lotul pentru {k} nu a fost selectat.")
                    return
                snapshot[k] = []
                continue
            lot["consum"] = to_float(lot.get("consum")) + calc["rezultatTotal"].get(k, 0.0)
            snapshot[k] = [{"lotId": lot["id"], "kg": calc["rezultatTotal"].get(k, 0.0)}]
        bar["consumApplied"] = True
        bar["consumSnapshot"] = snapshot
        bar["calcSnapshot"] = calc
        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()
        QMessageBox.information(self, "Succes", "Consumul a fost aplicat pe stoc!")

    def _revert_consum(self, order_id, recipe_id, bar_id):
        if QMessageBox.question(self, "Confirmare",
                                 "Anulezi consumul aplicat de aceasta bara? Stocul va fi repus.") != QMessageBox.Yes:
            return
        r = self._gaseste_reteta(order_id, recipe_id)
        bar = next(b for b in r["bare"] if b["id"] == bar_id)
        if bar.get("consumSnapshot"):
            for parti in bar["consumSnapshot"].values():
                intrari = parti if isinstance(parti, list) else [parti]
                for snap in intrari:
                    lot = next((l for l in self.state["lots"] if l["id"] == snap["lotId"]), None)
                    if lot:
                        lot["consum"] = max(0, to_float(lot.get("consum")) - snap["kg"])
        bar["consumApplied"] = False
        bar["consumSnapshot"] = None
        bar["calcSnapshot"] = None
        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()
        QMessageBox.information(self, "Succes", "Consumul a fost anulat!")

    def _muta_bara_pe_urmatoarea_reteta(self, order_id, recipe_id, bar_id):
        order = self._gaseste_comanda(order_id)
        r = self._gaseste_reteta(order_id, recipe_id)
        bar = next(b for b in r["bare"] if b["id"] == bar_id)

        if bar.get("consumApplied"):
            QMessageBox.warning(
                self, "Eroare",
                "Aceasta bara are deja consum aplicat pe stoc. Anuleaza consumul mai intai daca vrei sa o muti."
            )
            return

        try:
            idx = order["retete"].index(r)
        except ValueError:
            return

        if idx + 1 >= len(order["retete"]):
            QMessageBox.warning(
                self, "Nu exista reteta urmatoare",
                "Aceasta este ultima reteta din comanda \u2014 nu mai e nicio reteta pe care sa muti bara. "
                "Adauga mai intai o reteta noua in aceasta comanda daca vrei sa muti bara pe ea."
            )
            return

        reteta_urmatoare = order["retete"][idx + 1]
        tip_aliaj = order.get("tipAliaj", ALIAJ_IMPLICIT)

        # Dozarea (portia) barei se calculeaza o singura data, pe baza retetei
        # curente — aceasta ramane fixa indiferent unde ajunge bara.
        lot_sel_curent = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        calc = calculeaza_bara(tip_aliaj, r["target"], lot_sel_curent, r.get("portie"), r.get("numarPresari") or 1)
        if calc.get("eroare"):
            QMessageBox.warning(self, "Eroare", calc["eroare"])
            return

        lot_sel_urmator = {k: next((l for l in self.state["lots"] if l["id"] == reteta_urmatoare["lotSel"].get(k)), None) for k in ORDINE_MAT}

        # Verificam dinainte daca vreun material are nevoie de un lot "de
        # rezerva" (stocul din lotul curent nu ajunge) si daca acel lot e
        # deja selectat in reteta urmatoare — altfel oprim mutarea aici.
        lipsuri = []
        for k in ORDINE_MAT:
            necesar = calc["rezultatTotal"].get(k, 0.0)
            rest_vechi = lot_rest(lot_sel_curent[k]) if lot_sel_curent[k] else 0
            if necesar > rest_vechi + 0.001 and lot_sel_urmator[k] is None:
                lipsuri.append(next(m["nume"] for m in MATERIALE if m["id"] == k))
        if lipsuri:
            QMessageBox.warning(
                self, "Lot lipsa in reteta urmatoare",
                f"Stocul din lotul curent nu ajunge pentru: {', '.join(lipsuri)}. "
                f"Selecteaza mai intai un lot pentru aceste materiale in reteta \u201e{reteta_urmatoare['nume']}\u201d, "
                "apoi incearca din nou mutarea."
            )
            return

        raspuns = QMessageBox.question(
            self, "Confirmare",
            f"Muti aceasta bara pe reteta \u201e{reteta_urmatoare['nume']}\u201d? "
            "Dozarea calculata initial pentru aceasta bara ramane neschimbata. "
            "Ce nu incape in lotul curent va fi consumat automat din lotul selectat in reteta urmatoare."
        )
        if raspuns != QMessageBox.Yes:
            return

        # Aplicam efectiv consumul: pentru fiecare material, intai golim ce
        # mai e disponibil in lotul curent, apoi luam diferenta din lotul
        # selectat in reteta urmatoare.
        snapshot = {}
        for k in ORDINE_MAT:
            necesar = calc["rezultatTotal"].get(k, 0.0)
            lot_vechi = lot_sel_curent[k]
            if lot_vechi is None:
                # Material fara lot selectat, deoarece nu e necesar in
                # aceasta reteta (tinta elementului asociat e 0%).
                snapshot[k] = []
                continue
            rest_vechi = lot_rest(lot_vechi)
            parti = []
            if necesar <= rest_vechi + 0.001 or lot_sel_urmator[k] is None:
                # Incape tot in lotul curent — consum normal, ca pana acum.
                lot_vechi["consum"] = to_float(lot_vechi.get("consum")) + necesar
                parti.append({"lotId": lot_vechi["id"], "kg": necesar})
            else:
                din_vechi = max(0.0, rest_vechi)
                din_nou = necesar - din_vechi
                if din_vechi > 1e-9:
                    lot_vechi["consum"] = to_float(lot_vechi.get("consum")) + din_vechi
                    parti.append({"lotId": lot_vechi["id"], "kg": din_vechi})
                lot_nou = lot_sel_urmator[k]
                lot_nou["consum"] = to_float(lot_nou.get("consum")) + din_nou
                parti.append({"lotId": lot_nou["id"], "kg": din_nou})
            snapshot[k] = parti

        bar["calcSnapshot"] = calc
        bar["consumApplied"] = True
        bar["consumSnapshot"] = snapshot

        r["bare"] = [b for b in r["bare"] if b["id"] != bar_id]
        reteta_urmatoare["bare"].append(bar)
        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()
        QMessageBox.information(
            self, "Succes",
            f"Bara a fost mutata pe reteta \u201e{reteta_urmatoare['nume']}\u201d si consumul a fost aplicat "
            "(impartit intre lotul curent si lotul din reteta noua, acolo unde a fost cazul)."
        )

    # ------------------------------------------------------------------
    # Planificarea barelor: barele apartin COMENZII, distributia pe retete
    # rezulta din loturi
    # ------------------------------------------------------------------
    # Regula, pe scurt:
    #   - numarul total de bare e cunoscut preliminar, pe COMANDA;
    #   - fiecare reteta primeste atatea bare cate incap in loturile ei;
    #   - PRIMA bara care nu mai incape trece pe reteta urmatoare, dar isi
    #     pastreaza DOZAREA VECHE (calculata pe loturile retetei anterioare)
    #     si consuma resturile vechi + diferenta din loturile noi;
    #   - TOATE celelalte bare ramase se dozeaza de la zero, pe loturile noi;
    #   - mecanismul se repeta pentru fiecare reteta urmatoare.

    def _loturi_selectate(self, r):
        """Loturile efectiv selectate in reteta, indexate pe material."""
        return {
            k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None)
            for k in ORDINE_MAT
        }

    def _reteta_continuare(self, order, model):
        """Reteta noua, goala, in care utilizatorul introduce loturile noi."""
        return {
            "id": uid(),
            "nume": f"Reteta {len(order['retete']) + 1}",
            "target": dict(model["target"]),
            "lotSel": {},
            "bare": [],
            "numarPresari": model.get("numarPresari") or "1",
            "portie": model.get("portie"),
        }

    def _planifica_comanda(self, order, creeaza_retete=True):
        """Distribuie barele comenzii pe retete, in functie de loturi.

        Nu modifica barele si nu atinge stocul — doar calculeaza planul.
        Singurul efect lateral (daca creeaza_retete=True) e adaugarea
        retetelor de continuare de care e nevoie ca sa incapa toate barele.
        """
        total = max(0, int(to_float(order.get("nrBare"))))
        plan = []
        report = None      # bara de report care coboara in reteta urmatoare
        asignate = 0
        idx = 0
        limita_retete = len(order["retete"]) + 50   # plasa de siguranta

        while idx < len(order["retete"]) and idx < limita_retete:
            r = order["retete"][idx]
            lot_sel = self._loturi_selectate(r)

            # Barele cu consum deja aplicat sunt batute in cuie: raman unde
            # sunt si se scad din totalul comenzii.
            aplicate = [b for b in r["bare"] if b.get("consumApplied")]
            asignate += len(aplicate)
            ramase = total - asignate

            intrare = {
                "reteta": r, "report": report, "bareNoi": 0, "bareAsteptare": 0,
                "reportOut": None, "desf": None, "calc": None, "aplicate": len(aplicate),
                "stare": "ok",
            }

            if ramase <= 0:
                intrare["stare"] = "fara bare ramase"
                plan.append(intrare)
                report = None
                idx += 1
                continue

            calc = calculeaza_bara(
                order.get("tipAliaj", ALIAJ_IMPLICIT), r["target"], lot_sel,
                r.get("portie"), r.get("numarPresari") or 1,
            )
            if calc.get("eroare"):
                # Reteta inca nu are loturile complete: barele stau aici in
                # asteptare, inclusiv bara de report venita de mai sus.
                intrare["stare"] = "loturi incomplete"
                intrare["eroare"] = calc["eroare"]
                intrare["bareAsteptare"] = max(0, ramase - (1 if report else 0))
                plan.append(intrare)
                report = None
                break

            disponibile = max(0, ramase - (1 if report else 0))
            resturi = {k: (lot_rest(lot_sel[k]) if lot_sel[k] else 0.0) for k in ORDINE_MAT}
            desf = desfasoara_reteta(
                calc["rezultatTotal"], resturi, disponibile,
                report["ramas"] if report else None,
            )

            intrare["calc"] = calc
            intrare["desf"] = desf
            intrare["bareNoi"] = desf["bareIntregi"]
            asignate += (1 if report else 0) + desf["bareIntregi"]

            if desf["baraReport"]:
                intrare["reportOut"] = {
                    "ramas": {k: v for k, v in desf["baraReport"]["dinLotNou"].items() if v > 1e-9},
                    "dinLotAnterior": {k: v for k, v in desf["baraReport"]["dinLotCurent"].items() if v > 1e-9},
                    "calc": calc,
                    "retetaSursa": r["nume"],
                }

            plan.append(intrare)
            report = intrare["reportOut"]
            idx += 1

            # Mai sunt bare de plasat dar nu mai exista reteta urmatoare?
            if idx >= len(order["retete"]) and (total - asignate) > 0 and creeaza_retete:
                order["retete"].append(self._reteta_continuare(order, r))

        return plan

    def _aplica_plan(self, order, plan):
        """Rescrie listele de bare ale retetelor conform planului.

        Barele existente sunt REFOLOSITE (se pastreaza id-ul si eventualele
        date deja inregistrate), ca sa nu se piarda istoricul la fiecare
        redistribuire.
        """
        for intrare in plan:
            r = intrare["reteta"]
            aplicate = [b for b in r["bare"] if b.get("consumApplied")]
            libere = [b for b in r["bare"] if not b.get("consumApplied")]

            def ia_bara():
                return libere.pop(0) if libere else {"id": uid(), "consumApplied": False}

            noi = []

            # 1. Bara de report: pastreaza dozarea retetei ANTERIOARE.
            if intrare["report"]:
                bar = ia_bara()
                bar["tipDozare"] = "report"
                bar["reportRamas"] = intrare["report"]["ramas"]
                bar["reportDinLotAnterior"] = intrare["report"]["dinLotAnterior"]
                bar["calcSnapshot"] = intrare["report"]["calc"]
                bar["retetaSursaDozare"] = intrare["report"]["retetaSursa"]
                noi.append(bar)

            # 2. Barele normale: se dozeaza pe loturile ACESTEI retete.
            for _ in range(intrare["bareNoi"]):
                bar = ia_bara()
                bar["tipDozare"] = "normal"
                for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare"):
                    bar.pop(cheie, None)
                noi.append(bar)

            # 3. Bare in asteptare (reteta inca fara loturi selectate).
            for _ in range(intrare["bareAsteptare"]):
                bar = ia_bara()
                bar["tipDozare"] = "asteptare"
                for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare"):
                    bar.pop(cheie, None)
                noi.append(bar)

            r["bare"] = aplicate + noi

    def _distribuie_bare_comanda(self, order_id):
        order = self._gaseste_comanda(order_id)
        total = max(0, int(to_float(order.get("nrBare"))))
        if total <= 0:
            QMessageBox.information(
                self, "Numar de bare",
                "Seteaza mai intai numarul total de bare al comenzii."
            )
            return
        if not order["retete"]:
            QMessageBox.information(
                self, "Nicio reteta",
                "Adauga mai intai o reteta si selecteaza loturile de start."
            )
            return

        plan = self._planifica_comanda(order)
        self._aplica_plan(order, plan)
        salveaza_date(self.state)
        self._rebuild_retete()

        # --- Rezumat pentru utilizator --------------------------------
        linii = []
        plasate = 0
        for intrare in plan:
            r = intrare["reteta"]
            bucati = []
            if intrare["aplicate"]:
                bucati.append(f"{intrare['aplicate']} cu consum aplicat")
            if intrare["report"]:
                bucati.append(f"1 de report (dozare de pe {intrare['report']['retetaSursa']})")
            if intrare["bareNoi"]:
                bucati.append(f"{intrare['bareNoi']} dozate pe loturile de aici")
            if intrare["bareAsteptare"]:
                bucati.append(f"{intrare['bareAsteptare']} in asteptare (lipsesc loturile)")
            plasate += len(r["bare"])
            descriere = ", ".join(bucati) if bucati else "nicio bara"
            if intrare["stare"] == "loturi incomplete":
                descriere += f" \u2014 {intrare.get('eroare', '')}"
            linii.append(f"\u2022 {r['nume']}: {descriere}")

        mesaj = f"Comanda are {total} bare. Distributie:\n\n" + "\n".join(linii)
        neplasate = total - plasate
        if neplasate > 0:
            mesaj += (
                f"\n\nMai raman {neplasate} bare neplasate \u2014 selecteaza loturile "
                "pe ultima reteta si apasa din nou \u201eDistribuie barele\u201d."
            )
        QMessageBox.information(self, "Distributie bare", mesaj)

    def _seteaza_nr_bare_comanda(self, order_id, text):
        order = self._gaseste_comanda(order_id)
        order["nrBare"] = str(max(0, int(to_float(text))))
        salveaza_date(self.state)
        self._rebuild_retete()

    # ------------------------------------------------------------------
    # Desfasurarea / aplicarea consumului pe o reteta
    # ------------------------------------------------------------------
    def _desfasurare_reteta(self, order, r):
        """(calc, desf, lot_sel) pentru o singura reteta, sau (None, None, None)."""
        lot_sel = self._loturi_selectate(r)
        calc = calculeaza_bara(
            order.get("tipAliaj", ALIAJ_IMPLICIT), r["target"], lot_sel,
            r.get("portie"), r.get("numarPresari") or 1,
        )
        if calc.get("eroare"):
            QMessageBox.warning(self, "Eroare", calc["eroare"])
            return None, None, None

        # Bara de report nu consuma o doza intreaga, ci doar diferenta
        # ramasa; e numarata separat, ca sa nu iasa dubla.
        report = {k: 0.0 for k in ORDINE_MAT}
        for bar in r["bare"]:
            if bar.get("consumApplied"):
                continue
            for k, kg in (bar.get("reportRamas") or {}).items():
                report[k] = report.get(k, 0.0) + to_float(kg)

        normale = [
            b for b in r["bare"]
            if not b.get("consumApplied") and not b.get("reportRamas")
        ]
        resturi = {k: (lot_rest(lot_sel[k]) if lot_sel[k] else 0.0) for k in ORDINE_MAT}
        desf = desfasoara_reteta(calc["rezultatTotal"], resturi, len(normale), report)
        return calc, desf, lot_sel

    def _capacitate_reteta(self, order_id, recipe_id):
        order = self._gaseste_comanda(order_id)
        r = self._gaseste_reteta(order_id, recipe_id)
        if not r["bare"]:
            QMessageBox.information(
                self, "Nicio bara",
                "Seteaza numarul de bare pe comanda si apasa \u201eDistribuie barele\u201d."
            )
            return
        calc, desf, lot_sel = self._desfasurare_reteta(order, r)
        if desf is None:
            return
        dlg = DialogCapacitateReteta(desf, self, poate_aplica=self.poate_edita)
        dlg.exec()
        if dlg.actiune == "aplica":
            self._aplica_consum_total_reteta(order_id, recipe_id)

    def _aplica_consum_total_reteta(self, order_id, recipe_id):
        """Aplica O SINGURA DATA consumul intregii retete pe loturi.

        Bara de report (mostenita din reteta anterioara) e un caz special:
        materialul care s-a terminat acolo isi completeaza acum diferenta
        din loturile de aici (deja calculat in desf["consumTotalReteta"]).
        Restul materialelor acelei bare INCA NU au fost consumate nicaieri
        (au fost lasate neatinse in reteta anterioara) — se consuma abia
        ACUM, din loturile acestei retete, dar cu DOZAREA VECHE
        (calcSnapshot al barei), nu cu dozarea recalculata aici.
        """
        order = self._gaseste_comanda(order_id)
        r = self._gaseste_reteta(order_id, recipe_id)
        calc, desf, lot_sel = self._desfasurare_reteta(order, r)
        if desf is None:
            return

        neaplicate = [b for b in r["bare"] if not b.get("consumApplied")]
        if not neaplicate:
            QMessageBox.information(
                self, "Nimic de aplicat",
                "Toate barele acestei retete au deja consum aplicat."
            )
            return

        mesaj = (
            f"Se aplica pe stoc consumul retetei \u201e{r['nume']}\u201d "
            f"({desf['bareIntregi']} bare dozate aici"
            + (" + bara de report" if any(desf["reportMostenit"].values()) else "")
            + ")."
        )
        if QMessageBox.question(self, "Confirmare", mesaj) != QMessageBox.Yes:
            return

        # --- 0. Cat mai trebuie consumat, in plus fata de desf, pentru
        #        materialele NElimitante ale barei de report care se
        #        rezolva acum (vezi docstring). ---------------------------
        extra_consum = {k: 0.0 for k in ORDINE_MAT}
        bara_report_curenta = next(
            (b for b in r["bare"] if b.get("reportRamas") and not b.get("consumApplied")), None
        )
        report_rezolvat = bara_report_curenta is not None and not (
            {k: v for k, v in desf["reportLipsa"].items() if v > 1e-9}
        )
        if report_rezolvat:
            calc_vechi = bara_report_curenta.get("calcSnapshot") or calc
            doza_veche = calc_vechi["rezultatTotal"]
            for k in ORDINE_MAT:
                if desf["reportAcoperit"].get(k, 0.0) > 1e-9:
                    continue  # material limitant, deja acoperit in desf
                kg = r2(to_float(doza_veche.get(k, 0.0)))
                if kg > 1e-9:
                    extra_consum[k] = kg

        # --- 1. Scaderea efectiva din loturile curente -----------------
        for k in ORDINE_MAT:
            kg = r2(desf["consumTotalReteta"].get(k, 0.0) + extra_consum[k])
            lot = lot_sel[k]
            if lot is None:
                if kg > TOLERANTA_KG and material_necesar(k, r["target"]):
                    QMessageBox.warning(self, "Eroare", f"Lotul pentru {k} nu a fost selectat.")
                    return
                continue
            lot["consum"] = r2(to_float(lot.get("consum")) + kg)

        # --- 2. Inregistrarea pe bare (istoric, Fisa limita) ------------
        doza = calc["rezultatTotal"]
        index_pas = 0
        for bar in r["bare"]:
            if bar.get("consumApplied"):
                continue

            # Bara de report isi stinge datoria fata de loturile de aici.
            if bar.get("reportRamas"):
                calc_vechi = bar.get("calcSnapshot") or calc
                doza_veche = calc_vechi["rezultatTotal"]
                snapshot = bar.get("consumSnapshot") or {}

                # Materialul limitant: doar diferenta acoperita acum.
                for k in ORDINE_MAT:
                    kg = desf["reportAcoperit"].get(k, 0.0)
                    if kg <= 1e-9 or lot_sel[k] is None:
                        continue
                    snapshot.setdefault(k, [])
                    snapshot[k].append({"lotId": lot_sel[k]["id"], "kg": kg})

                ramas_dupa = {k: v for k, v in desf["reportLipsa"].items() if v > 1e-9}
                if ramas_dupa:
                    # Inca nu ajunge nici din loturile de aici -> ramane
                    # in asteptare, coboara mai departe (caz rar).
                    bar["consumSnapshot"] = snapshot
                    bar["reportRamas"] = ramas_dupa
                    continue

                # S-a rezolvat: acum se consuma, cu dozarea VECHE, si
                # restul materialelor (cele care nu au limitat).
                for k in ORDINE_MAT:
                    if desf["reportAcoperit"].get(k, 0.0) > 1e-9:
                        continue
                    kg = r2(to_float(doza_veche.get(k, 0.0)))
                    if kg <= 1e-9 or lot_sel[k] is None:
                        continue
                    snapshot.setdefault(k, [])
                    snapshot[k].append({"lotId": lot_sel[k]["id"], "kg": kg})

                bar["consumSnapshot"] = snapshot
                bar.pop("reportRamas", None)
                bar["consumApplied"] = True
                continue

            pas = desf["pasi"][index_pas] if index_pas < len(desf["pasi"]) else None
            index_pas += 1
            if pas is None or not pas["incape"]:
                continue

            bar["consumSnapshot"] = {
                k: ([{"lotId": lot_sel[k]["id"], "kg": r2(doza.get(k, 0.0))}]
                    if lot_sel[k] and doza.get(k, 0.0) > 1e-9 else [])
                for k in ORDINE_MAT
            }
            bar["consumApplied"] = True
            bar["calcSnapshot"] = calc

        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()
        QMessageBox.information(self, "Succes", "Consumul a fost aplicat pe toata reteta.")

    def _bara_report_trimisa(self, order, r):
        """(reteta_urmatoare, bara) daca ACEASTA reteta a trimis mai
        departe o bara de report (adica o bara de-a ei nu a incaput toata
        aici si a fost mutata, cu dozarea veche, pe reteta urmatoare).
        Returneaza (None, None) daca nu exista asa ceva sau nu exista
        reteta urmatoare.
        """
        try:
            idx = order["retete"].index(r)
        except ValueError:
            return None, None
        if idx + 1 >= len(order["retete"]):
            return None, None
        r_urm = order["retete"][idx + 1]
        bar = next(
            (b for b in r_urm["bare"] if b.get("retetaSursaDozare") == r["nume"]),
            None,
        )
        return r_urm, bar

    def _are_consum_de_retras(self, order, r):
        """True daca exista ceva de retras pentru aceasta reteta: fie bare
        cu consum aplicat aici, fie o bara de report trimisa mai departe
        care a mai luat deja materialul limitant din loturile de aici."""
        if any(b.get("consumApplied") for b in r["bare"]):
            return True
        _, bar_out = self._bara_report_trimisa(order, r)
        return bool(bar_out and bar_out.get("reportDinLotAnterior"))

    def _retrage_consum_total_reteta(self, order_id, recipe_id):
        """Retrage O SINGURA DATA tot consumul aplicat pe stoc pentru o
        reteta, simetric cu ``_aplica_consum_total_reteta``: repune in
        loturi tot ce s-a scazut de acolo pentru barele acestei retete.

        Cazul special e bara de report: daca aceasta reteta a "cedat" o
        bara pe reteta urmatoare (materialul limitant s-a terminat aici),
        acea bara a consumat deja, direct din loturile ACESTEI retete,
        partea care s-a mai gasit (``reportDinLotAnterior``) — fara sa
        fie inregistrata pe nicio bara de-a acestei retete. La retragere,
        acea cantitate trebuie repusa aici, iar bara redevine "in
        asteptare" pe reteta urmatoare (isi pastreaza dozarea veche, dar
        asteapta din nou sa fie rezolvata).

        Daca bara de report a fost deja REZOLVATA pe reteta urmatoare
        (adica s-a aplicat consumul si acolo), retragerea se blocheaza —
        trebuie retras mai intai consumul de acolo, in ordine inversa.
        """
        order = self._gaseste_comanda(order_id)
        r = self._gaseste_reteta(order_id, recipe_id)

        aplicate = [b for b in r["bare"] if b.get("consumApplied")]
        r_urm, bar_out = self._bara_report_trimisa(order, r)
        out_are_consum = bool(bar_out and bar_out.get("reportDinLotAnterior"))
        out_rezolvata = bool(bar_out and bar_out.get("consumApplied"))

        if not aplicate and not out_are_consum:
            QMessageBox.information(
                self, "Nimic de retras",
                "Aceasta reteta nu are niciun consum aplicat pe stoc."
            )
            return

        if out_rezolvata:
            QMessageBox.warning(
                self, "Nu se poate retrage",
                f"Bara de report trimisa de aceasta reteta a fost deja consumata "
                f"pe reteta \u201e{r_urm['nume']}\u201d. Retrage mai intai consumul "
                "de acolo, apoi revino la aceasta reteta."
            )
            return

        mesaj = (
            f"Retragi tot consumul aplicat pe stoc de reteta \u201e{r['nume']}\u201d "
            f"({len(aplicate)} bare cu consum aplicat"
            + (" + bara de report trimisa mai departe" if out_are_consum else "")
            + ")? Stocul va fi repus la loc."
        )
        if QMessageBox.question(self, "Confirmare", mesaj) != QMessageBox.Yes:
            return

        # --- 1. Bare cu consum aplicat AICI (normale, mutate manual, sau
        #        bare de report inca din reteta anterioara si rezolvate
        #        chiar in aceasta reteta). ------------------------------
        for bar in aplicate:
            era_report_aici = bar.get("tipDozare") == "report" and bar.get("reportDinLotAnterior")
            for parti in (bar.get("consumSnapshot") or {}).values():
                intrari = parti if isinstance(parti, list) else [parti]
                for snap in intrari:
                    lot = next((l for l in self.state["lots"] if l["id"] == snap["lotId"]), None)
                    if lot:
                        lot["consum"] = max(0.0, r2(to_float(lot.get("consum")) - to_float(snap.get("kg"))))
            bar["consumApplied"] = False
            bar["consumSnapshot"] = None
            if era_report_aici:
                # Redevine bara de report NEREZOLVATA: pastreaza dozarea
                # veche (calcSnapshot) si reconstruieste ce mai are de
                # luat din lotul nou, ca sa astepte din nou rezolvarea.
                doza_veche = (bar.get("calcSnapshot") or {}).get("rezultatTotal", {})
                din_anterior = bar.get("reportDinLotAnterior") or {}
                ramas = {
                    k: r2(to_float(doza_veche.get(k, 0.0)) - to_float(din_anterior.get(k, 0.0)))
                    for k in din_anterior
                }
                bar["reportRamas"] = {k: v for k, v in ramas.items() if v > 1e-9}
            else:
                bar.pop("calcSnapshot", None)

        # --- 2. Bara de report TRIMISA mai departe de aceasta reteta:
        #        repune partea deja luata din loturile de aici si o
        #        redeschide (nu mai e nici "report rezolvata", nici
        #        "consum aplicat" - revine "in asteptare"). -------------
        if out_are_consum:
            lot_sel = self._loturi_selectate(r)
            for k, kg in bar_out["reportDinLotAnterior"].items():
                kg = to_float(kg)
                lot = lot_sel.get(k)
                if lot and kg > 1e-9:
                    lot["consum"] = max(0.0, r2(to_float(lot.get("consum")) - kg))
            for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare", "calcSnapshot"):
                bar_out.pop(cheie, None)
            bar_out["tipDozare"] = "asteptare"
            bar_out["consumApplied"] = False
            bar_out["consumSnapshot"] = None

        salveaza_date(self.state)
        self._rebuild_stoc()
        self._rebuild_retete()
        QMessageBox.information(self, "Succes", "Consumul retetei a fost retras de pe stoc.")
