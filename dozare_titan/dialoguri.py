"""Ferestre de dialog: autentificare admin, adaugare/editare lot, istoric lot.

Daca un aliaj nou are elemente in plus in compozitia unui lot (ex. Si, Zr,
Mo pentru Ti-VT9), locul de adaugat campurile noi e in DialogLotNou,
variabila `etichete_doza` din __init__.
"""

from datetime import date

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QGridLayout, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QListWidget, QMessageBox, QPushButton, QScrollArea,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from . import standarde
from .config import MATERIALE, ORDINE_MAT, ALIAJE_SPEC
from . import conturi as conturi_mod
from .stiluri import STIL_BUTON_PRINCIPAL, STIL_BUTON_SECUNDAR, STIL_BUTON_PERICOL, STIL_CAMP, CULOARE_EROARE, CULOARE_SUCCES, CULOARE_GRI_TEXT, CULOARE_BORDURA, CULOARE_FUNDAL_SECTIUNE
from .utils import fmt, to_float, r2
from .calcule import lot_ti, lot_rest, nume_material


class DialogLogin(QDialog):
    """Fereastra de login, obligatorie la pornirea aplicatiei (blocheaza
    fereastra principala pana la o autentificare reusita).

    Conturile sunt gestionate in dozare_titan/conturi.json (parole
    hash-uite) — vezi modulul .conturi. Oricine isi poate cere cont nou
    din acest ecran ("Creeaza cont nou"); contul ramane in asteptare pana
    cand un administrator il aproba si ii acorda drepturi (vezi
    DialogAdministrareConturi, disponibil in fereastra principala pentru
    conturile cu "admin": True).

    Privilegiile de editare (adaugare/editare loturi, creare/stergere de
    comenzi si retete, aplicare consum pe stoc) sunt acordate per cont
    ("editor": True/False). Un cont fara acest drept poate doar vizualiza.

    La succes, self.rezultat contine dict-ul contului logat (utilizator,
    alias, editor, admin, ...) — "alias" e numele complet folosit apoi la
    rubrica "Intocmit" pe documentele generate; ramane pe documente
    indiferent cat de scurt e numele de utilizator (login)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Autentificare \u2014 Dozare lingouri titan")
        self.setFixedWidth(340)
        self.rezultat = None
        self.conturi = conturi_mod.incarca_conturi()

        layout = QVBoxLayout(self)
        info = QLabel("Introdu contul tau pentru a intra in aplicatie.")
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {CULOARE_GRI_TEXT};")
        layout.addWidget(info)

        layout.addWidget(QLabel("Utilizator"))
        self.camp_utilizator = QLineEdit()
        self.camp_utilizator.setStyleSheet(STIL_CAMP)
        self.camp_utilizator.returnPressed.connect(self._verifica)
        layout.addWidget(self.camp_utilizator)

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
        btn_iesire = QPushButton("Iesire")
        btn_iesire.setStyleSheet(STIL_BUTON_SECUNDAR)
        # Explicit FARA autoDefault — altfel, in anumite conditii, o
        # apasare de Enter in campurile de mai sus putea activa acest
        # buton (deci inchidea toata aplicatia) in loc sa doar verifice
        # datele introduse. Butonul de iesire trebuie apasat manual, cu
        # mouse-ul/tab+Enter direct pe el — niciodata din campurile de
        # utilizator/parola.
        btn_iesire.setAutoDefault(False)
        btn_iesire.setDefault(False)
        btn_iesire.clicked.connect(self.reject)
        btn_intra = QPushButton("Intra")
        btn_intra.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_intra.setAutoDefault(True)
        btn_intra.setDefault(True)
        btn_intra.clicked.connect(self._verifica)
        rand.addWidget(btn_iesire)
        rand.addWidget(btn_intra)
        layout.addLayout(rand)

        btn_cont_nou = QPushButton("Creeaza cont nou")
        btn_cont_nou.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_cont_nou.setAutoDefault(False)
        btn_cont_nou.clicked.connect(self._deschide_creare_cont)
        layout.addWidget(btn_cont_nou)

        self.camp_utilizator.setFocus()

    def _deschide_creare_cont(self):
        dialog = DialogContNou(self.conturi, self)
        if dialog.exec() == QDialog.Accepted:
            self.eticheta_eroare.setStyleSheet(f"color: {CULOARE_SUCCES}; font-size: 12px;")
            self.eticheta_eroare.setText(
                "Cererea de cont a fost trimisa. Poti intra dupa ce un "
                "administrator o aproba."
            )
            self.camp_utilizator.setText(dialog.utilizator_creat)
            self.camp_parola.clear()
            self.camp_parola.setFocus()

    def _verifica(self):
        utilizator = self.camp_utilizator.text().strip()
        parola = self.camp_parola.text()
        self.eticheta_eroare.setStyleSheet(f"color: {CULOARE_EROARE}; font-size: 12px;")
        if not utilizator:
            self.eticheta_eroare.setText("Introdu numele de utilizator.")
            self.camp_utilizator.setFocus()
            return

        # Reincarcam conturile la fiecare incercare, ca sa vedem imediat o
        # eventuala aprobare facuta intre timp de un administrator.
        self.conturi = conturi_mod.incarca_conturi()
        cont = conturi_mod.gaseste_cont(self.conturi, utilizator)

        if cont is None or not conturi_mod.verifica_parola(parola, cont.get("hash"), cont.get("sare")):
            self.eticheta_eroare.setText("Nume de utilizator sau parola gresita.")
            self.camp_parola.clear()
            self.camp_parola.setFocus()
            return

        stare = cont.get("stare", conturi_mod.STARE_APROBAT)
        if stare == conturi_mod.STARE_ASTEPTARE:
            self.eticheta_eroare.setText(
                "Contul asteapta aprobarea administratorului."
            )
            return
        if stare == conturi_mod.STARE_RESPINS:
            self.eticheta_eroare.setText(
                "Cererea de cont a fost respinsa. Contacteaza administratorul."
            )
            return

        self.rezultat = cont
        self.accept()


class DialogContNou(QDialog):
    """Cerere de cont nou, facuta chiar de utilizator din ecranul de
    autentificare. Numele de utilizator (login) poate fi scurt — separat
    de alias, numele complet care ramane afisat pe documentele generate.
    Contul creat asteapta aprobarea unui administrator inainte de a putea
    fi folosit (vezi DialogAdministrareConturi)."""

    def __init__(self, conturi, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cont nou")
        self.setFixedWidth(340)
        self.conturi = conturi
        self.utilizator_creat = None

        layout = QVBoxLayout(self)
        info = QLabel(
            "Contul creat asteapta aprobarea unui administrator inainte "
            "de a putea fi folosit."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 12px;")
        layout.addWidget(info)

        layout.addWidget(QLabel("Utilizator (scurt, folosit doar la login)"))
        self.camp_utilizator = QLineEdit()
        self.camp_utilizator.setStyleSheet(STIL_CAMP)
        layout.addWidget(self.camp_utilizator)

        layout.addWidget(QLabel("Alias (nume complet, afisat pe documente)"))
        self.camp_alias = QLineEdit()
        self.camp_alias.setStyleSheet(STIL_CAMP)
        layout.addWidget(self.camp_alias)

        layout.addWidget(QLabel("Parola"))
        self.camp_parola = QLineEdit()
        self.camp_parola.setEchoMode(QLineEdit.Password)
        self.camp_parola.setStyleSheet(STIL_CAMP)
        layout.addWidget(self.camp_parola)

        layout.addWidget(QLabel("Confirma parola"))
        self.camp_parola2 = QLineEdit()
        self.camp_parola2.setEchoMode(QLineEdit.Password)
        self.camp_parola2.setStyleSheet(STIL_CAMP)
        layout.addWidget(self.camp_parola2)

        self.eticheta_eroare = QLabel("")
        self.eticheta_eroare.setStyleSheet(f"color: {CULOARE_EROARE}; font-size: 12px;")
        self.eticheta_eroare.setWordWrap(True)
        layout.addWidget(self.eticheta_eroare)

        rand = QHBoxLayout()
        btn_anuleaza = QPushButton("Anuleaza")
        btn_anuleaza.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_anuleaza.clicked.connect(self.reject)
        btn_trimite = QPushButton("Trimite cererea")
        btn_trimite.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_trimite.clicked.connect(self._trimite)
        rand.addWidget(btn_anuleaza)
        rand.addWidget(btn_trimite)
        layout.addLayout(rand)

        self.camp_utilizator.setFocus()

    def _trimite(self):
        utilizator = self.camp_utilizator.text().strip()
        alias = self.camp_alias.text().strip()
        parola = self.camp_parola.text()
        parola2 = self.camp_parola2.text()

        if not utilizator or not alias or not parola:
            self.eticheta_eroare.setText("Completeaza toate campurile.")
            return
        if " " in utilizator:
            self.eticheta_eroare.setText("Numele de utilizator nu poate contine spatii.")
            return
        if conturi_mod.utilizator_exista(self.conturi, utilizator):
            self.eticheta_eroare.setText("Exista deja un cont cu acest nume de utilizator.")
            return
        if parola != parola2:
            self.eticheta_eroare.setText("Parolele introduse nu coincid.")
            return
        if len(parola) < 4:
            self.eticheta_eroare.setText("Parola trebuie sa aiba cel putin 4 caractere.")
            return

        conturi_mod.creeaza_cererecont(self.conturi, utilizator, parola, alias)
        self.utilizator_creat = utilizator
        self.accept()


class DialogAdministrareConturi(QDialog):
    """Panou de administrare a conturilor, disponibil doar pentru conturile
    cu drept de administrare ("admin": True). Permite aprobarea/respingerea
    cererilor de cont noi si acordarea drepturilor de editare/administrare;
    aliasul (numele afisat pe documente) ramane editabil aici, separat de
    numele de utilizator (login)."""

    STARI = [conturi_mod.STARE_APROBAT, conturi_mod.STARE_ASTEPTARE, conturi_mod.STARE_RESPINS]
    STARI_ETICHETE = {
        conturi_mod.STARE_APROBAT: "Aprobat",
        conturi_mod.STARE_ASTEPTARE: "In asteptare",
        conturi_mod.STARE_RESPINS: "Respins",
    }

    def __init__(self, utilizator_curent, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Administrare conturi")
        self.resize(760, 420)
        self.utilizator_curent = utilizator_curent  # cheia (lowercase) contului logat
        self.conturi = conturi_mod.incarca_conturi()
        self.linii = {}  # cheie cont -> dict de widget-uri ale randului

        layout = QVBoxLayout(self)
        info = QLabel(
            "Aproba/respinge cererile de cont si acorda drepturi de editare "
            "sau de administrare. Aliasul e numele care ramane afisat pe "
            "documentele generate (\u201eIntocmit\u201d), indiferent cat de "
            "scurt e numele de utilizator."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 12px;")
        layout.addWidget(info)

        self.tabel = QTableWidget()
        self.tabel.setColumnCount(7)
        self.tabel.setHorizontalHeaderLabels(
            ["Utilizator", "Alias (pe documente)", "Stare", "Editor", "Admin", "", ""]
        )
        self.tabel.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabel.verticalHeader().setVisible(False)
        layout.addWidget(self.tabel)

        rand_jos = QHBoxLayout()
        rand_jos.addStretch(1)
        btn_salveaza = QPushButton("Salveaza modificarile")
        btn_salveaza.setStyleSheet(STIL_BUTON_PRINCIPAL)
        btn_salveaza.clicked.connect(self._salveaza)
        btn_inchide = QPushButton("Inchide")
        btn_inchide.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_inchide.clicked.connect(self.accept)
        rand_jos.addWidget(btn_inchide)
        rand_jos.addWidget(btn_salveaza)
        layout.addLayout(rand_jos)

        self._populeaza_tabel()

    def _populeaza_tabel(self):
        conturi_sortate = sorted(
            self.conturi.items(),
            key=lambda kv: (kv[1].get("stare") != conturi_mod.STARE_ASTEPTARE, kv[1].get("utilizator", "").lower()),
        )
        self.tabel.setRowCount(len(conturi_sortate))
        for rand, (cheie, cont) in enumerate(conturi_sortate):
            item_utilizator = QTableWidgetItem(cont.get("utilizator", cheie))
            item_utilizator.setFlags(item_utilizator.flags() & ~Qt.ItemIsEditable)
            self.tabel.setItem(rand, 0, item_utilizator)

            camp_alias = QLineEdit(cont.get("alias", ""))
            camp_alias.setStyleSheet(STIL_CAMP)
            self.tabel.setCellWidget(rand, 1, camp_alias)

            combo_stare = QComboBox()
            for s in self.STARI:
                combo_stare.addItem(self.STARI_ETICHETE[s], s)
            combo_stare.setCurrentIndex(self.STARI.index(cont.get("stare", conturi_mod.STARE_APROBAT)))
            self.tabel.setCellWidget(rand, 2, combo_stare)

            check_editor = QCheckBox()
            check_editor.setChecked(bool(cont.get("editor")))
            self.tabel.setCellWidget(rand, 3, self._centrat(check_editor))

            check_admin = QCheckBox()
            check_admin.setChecked(bool(cont.get("admin")))
            self.tabel.setCellWidget(rand, 4, self._centrat(check_admin))

            btn_sterge = QPushButton("Sterge")
            btn_sterge.setStyleSheet(STIL_BUTON_PERICOL)
            este_contul_propriu = cheie == self.utilizator_curent
            btn_sterge.setEnabled(not este_contul_propriu)
            btn_sterge.setToolTip("Nu iti poti sterge propriul cont." if este_contul_propriu else "")
            btn_sterge.clicked.connect(lambda _=False, c=cheie: self._sterge(c))
            self.tabel.setCellWidget(rand, 5, btn_sterge)

            if cont.get("stare") == conturi_mod.STARE_ASTEPTARE:
                btn_aproba = QPushButton("Aproba")
                btn_aproba.setStyleSheet(STIL_BUTON_PRINCIPAL)
                btn_aproba.clicked.connect(lambda _=False, s=combo_stare: s.setCurrentIndex(self.STARI.index(conturi_mod.STARE_APROBAT)))
                self.tabel.setCellWidget(rand, 6, btn_aproba)

            self.linii[cheie] = {
                "alias": camp_alias,
                "stare": combo_stare,
                "editor": check_editor,
                "admin": check_admin,
            }

    @staticmethod
    def _centrat(widget):
        w = QWidget()
        container = QVBoxLayout(w)
        container.setAlignment(Qt.AlignCenter)
        container.setContentsMargins(0, 0, 0, 0)
        container.addWidget(widget)
        return w

    def _sterge(self, cheie):
        if cheie == self.utilizator_curent:
            return
        nume = self.conturi.get(cheie, {}).get("utilizator", cheie)
        if QMessageBox.question(
            self, "Confirmare", f"Stergi definitiv contul \u201e{nume}\u201d?"
        ) != QMessageBox.Yes:
            return
        self.conturi.pop(cheie, None)
        conturi_mod.salveaza_conturi(self.conturi)
        self._populeaza_tabel()

    def _salveaza(self):
        # Nu lasam ultimul admin aprobat sa-si ia singur dreptul de admin,
        # ca sa nu ramana aplicatia fara niciun cont care poate aproba
        # conturi noi / acorda drepturi.
        va_ramane_admin = False
        for cheie, linie in self.linii.items():
            stare = linie["stare"].currentData()
            e_admin = linie["admin"].isChecked()
            if e_admin and stare == conturi_mod.STARE_APROBAT:
                va_ramane_admin = True

        if not va_ramane_admin:
            QMessageBox.warning(
                self, "Nu se poate salva",
                "Trebuie sa ramana cel putin un cont aprobat cu drept de "
                "administrare, altfel nimeni nu ar mai putea aproba conturi "
                "noi sau acorda drepturi."
            )
            return

        for cheie, linie in self.linii.items():
            cont = self.conturi.get(cheie)
            if cont is None:
                continue
            alias = linie["alias"].text().strip()
            cont["alias"] = alias or cont.get("utilizator", cheie)
            cont["stare"] = linie["stare"].currentData()
            cont["editor"] = linie["editor"].isChecked()
            cont["admin"] = linie["admin"].isChecked()

        conturi_mod.salveaza_conturi(self.conturi)
        QMessageBox.information(self, "Salvat", "Modificarile au fost salvate.")
        self._populeaza_tabel()


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
        ("dozaSn", "Doza Sn (%)"),
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
        self.camp_stoc = QLineEdit(fmt(lot.get("stocIntrare"), 2) if self.editare else "")
        self.camp_stoc.setPlaceholderText("0")

        grid.addWidget(QLabel("Nr. LOT"), 0, 0)
        grid.addWidget(self.camp_nr_lot, 1, 0)
        grid.addWidget(QLabel("Stoc intrare (kg)"), 0, 1)
        grid.addWidget(self.camp_stoc, 1, 1)

        self.camp_consum = None
        if self.editare:
            self.camp_consum = QLineEdit(fmt(lot.get("consum"), 2))
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

        # Concentratia de Ti se poate introduce DIRECT de la tastatura (din
        # buletinul de analiza al furnizorului). Daca campul ramane gol, Ti
        # se calculeaza automat ca pana acum (100 - suma elementelor, sau
        # 60% fix la TiO2, sau 0% la materialele fara Ti).
        rand_ti = QHBoxLayout()
        rand_ti.addWidget(QLabel("Ti % (manual, optional)"))
        ti_initial = lot.get("dozaTi") if self.editare else None
        self.camp_ti = QLineEdit(fmt(ti_initial, 3) if ti_initial not in (None, "", False) else "")
        self.camp_ti.setPlaceholderText("gol = calculat automat")
        self.camp_ti.textChanged.connect(self._actualizeaza_preview_ti)
        rand_ti.addWidget(self.camp_ti)
        layout.addLayout(rand_ti)

        campuri_stil = [self.camp_nr_lot, self.camp_stoc, self.camp_ti] + list(self.campuri_doza.values())
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
        manual = self.camp_ti.text().strip() if hasattr(self, "camp_ti") else ""
        if manual:
            self.eticheta_ti.setText(
                f"Ti introdus manual: {to_float(manual):.3f}% (are prioritate)"
            )
            return
        total = sum(to_float(c.text()) for c in self.campuri_doza.values())
        titlu = self.windowTitle()
        if "TiO2" in titlu or "TiO\u2082" in titlu:
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
        manual_ti = self.camp_ti.text().strip()
        if manual_ti and not (0 <= to_float(manual_ti) <= 100):
            self.eticheta_eroare.setText("Concentratia de Ti trebuie sa fie intre 0 si 100%.")
            return
        rezervare_id = self.combo_rezervare.currentData()
        self.rezultat = {
            "nrLot": nr_lot, "stocIntrare": r2(stoc), "consum": r2(consum),
            "dozaTi": r2(manual_ti) if manual_ti else None,
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


class DialogCapacitateReteta(QDialog):
    """Arata desfasurarea bara cu bara a unei retete: ce consuma fiecare
    bara la pasul ei si cat mai ramane din fiecare lot dupa acel pas.

    Nu aplica nimic pe stoc — doar afiseaza. Aplicarea se face din
    fereastra principala, o singura data, pentru toata reteta.
    """

    def __init__(self, desf, parent=None, poate_aplica=True):
        super().__init__(parent)
        self.setWindowTitle("Capacitate reteta \u2014 desfasurare pe bare")
        self.resize(940, 580)
        self.actiune = None
        self.desf = desf

        layout = QVBoxLayout(self)

        # Doar materialele care chiar intra in amestec (doza > 0).
        self.materiale = [k for k in ORDINE_MAT if desf["dozareBara"].get(k, 0) > 0]

        # --- Rezumatul de sus -----------------------------------------
        if desf["bareRamase"] > 0 and desf["baraReport"]:
            limitante = ", ".join(nume_material(k) for k in desf["materialeLimitante"])
            rezumat = (
                f"<b>Incap {desf['bareIntregi']} bare intregi</b> din cele "
                f"{desf['nrBareCerut']} cerute. Se termina: <b>{limitante}</b>.<br>"
                f"Bara {desf['baraReport']['index']} ia tot ce mai ramane din lotul care "
                f"s-a terminat si trece pe reteta urmatoare CU DOZAREA DE AICI; acolo "
                f"isi completeaza diferenta si consuma integral celelalte materiale. "
                f"Dupa ea mai raman {desf['bareRamase'] - 1} bare de turnat acolo, "
                f"dozate pe loturile noi."
            )
            culoare = CULOARE_EROARE
        else:
            rezumat = (
                f"<b>Toate cele {desf['nrBareCerut']} bare incap</b> in loturile "
                f"selectate. Capacitatea maxima a loturilor curente este de "
                f"{desf['capacitateMaxima']} bare."
            )
            culoare = CULOARE_SUCCES
        et = QLabel(rezumat)
        et.setWordWrap(True)
        et.setStyleSheet(f"color: {culoare}; font-size: 12px;")
        layout.addWidget(et)

        if any(desf["reportMostenit"].get(k, 0) > 0 for k in ORDINE_MAT):
            et_rep = QLabel(
                "Reteta contine o bara de report din reteta anterioara, cu dozarea "
                "ei veche \u2014 consumul ei se scade primul din loturile de aici, "
                "inaintea barelor dozate pe loturile de aici."
            )
            et_rep.setWordWrap(True)
            et_rep.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
            layout.addWidget(et_rep)

        # --- Tabelul de desfasurare -----------------------------------
        coloane = ["Bara"]
        for k in self.materiale:
            coloane += [f"{nume_material(k)}\nconsum (kg)", f"{nume_material(k)}\nrest lot (kg)"]
        coloane.append("Stare")

        tabel = QTableWidget(len(desf["pasi"]), len(coloane))
        tabel.setHorizontalHeaderLabels(coloane)
        tabel.verticalHeader().setVisible(False)
        tabel.setEditTriggers(QTableWidget.NoEditTriggers)
        tabel.setStyleSheet(
            f"QTableWidget {{ border: none; gridline-color: {CULOARE_BORDURA}; }}"
            f"QHeaderView::section {{ background-color: white; color: {CULOARE_GRI_TEXT}; "
            f"border: none; border-bottom: 1px solid {CULOARE_BORDURA}; padding: 6px; font-size: 10.5px; }}"
        )

        for rand, pas in enumerate(desf["pasi"]):
            tabel.setItem(rand, 0, QTableWidgetItem(f"Bara {pas['bara']}"))
            col = 1
            for k in self.materiale:
                m = pas["materiale"].get(k, {})
                if m.get("dinLotNou", 0) > 0:
                    it_consum = QTableWidgetItem(
                        f"{fmt(m.get('dinLotCurent'), 2)} + {fmt(m.get('dinLotNou'), 2)} "
                        "din reteta urmatoare"
                    )
                    it_consum.setForeground(QColor(CULOARE_EROARE))
                else:
                    it_consum = QTableWidgetItem(fmt(m.get("dinLotCurent"), 2))
                it_consum.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                tabel.setItem(rand, col, it_consum)

                it_rest = QTableWidgetItem(fmt(m.get("restDupa"), 2))
                it_rest.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if not m.get("incape", True):
                    it_rest.setForeground(QColor(CULOARE_EROARE))
                tabel.setItem(rand, col + 1, it_rest)
                col += 2

            if pas["incape"]:
                stare = QTableWidgetItem("OK")
                stare.setForeground(QColor(CULOARE_SUCCES))
            else:
                lipsa = ", ".join(nume_material(k) for k in pas["materialeLipsa"])
                stare = QTableWidgetItem(f"Nu incape \u2014 {lipsa}")
                stare.setForeground(QColor(CULOARE_EROARE))
            tabel.setItem(rand, col, stare)

        tabel.resizeColumnsToContents()
        tabel.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(tabel)

        # --- Butoane ---------------------------------------------------
        rand_btn = QHBoxLayout()
        btn_inchide = QPushButton("Inchide")
        btn_inchide.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_inchide.clicked.connect(self.reject)
        rand_btn.addWidget(btn_inchide)

        if poate_aplica:
            eticheta_btn = (
                "Aplica consumul pe toata reteta si creeaza reteta urmatoare"
                if desf["bareRamase"] > 0 else
                "Aplica consumul pe toata reteta"
            )
            btn_aplica = QPushButton(eticheta_btn)
            btn_aplica.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn_aplica.clicked.connect(self._aplica)
            rand_btn.addWidget(btn_aplica)

        layout.addLayout(rand_btn)

    def _aplica(self):
        self.actiune = "aplica"
        self.accept()


# Elementele de compozitie tinta afisate in configurarea comenzii, in
# functie de ce limite are aliajul in config.ALIAJE_SPEC.
_ELEMENTE_TINTA = [
    ("al", "Al [%]"), ("v", "V [%]"), ("mo", "Mo [%]"), ("sn", "Sn [%]"), ("zr", "Zr [%]"),
    ("si", "Si [%]"), ("o", "O [%]"), ("fe", "Fe [%]"),
]


def _elemente_pentru_aliaj(tip_aliaj, standard=None):
    """Ce campuri de compozitie tinta au sens pentru aliajul dat: cele
    care apar in limitele lui chimice (din config sau din standardul ales
    pe comanda), plus O si Fe (mereu prezente)."""
    spec = ALIAJE_SPEC.get(tip_aliaj, {})
    din_standard = {e.lower() for e in (standard or {}).get("limite", {})}
    elemente = []
    for cheie, eticheta in _ELEMENTE_TINTA:
        if (cheie in ("o", "fe") or f"{cheie}_min" in spec or f"{cheie}_max" in spec
                or cheie in din_standard):
            elemente.append((cheie, eticheta))
    return elemente


class DialogConfigurareComanda(QDialog):
    """Configurarea completa a unei comenzi, dintr-un singur loc: portia,
    numarul de presari, numarul total de bare, compozitia tinta si —
    partea noua — LISTA ORDONATA DE LOTURI alocate comenzii, pe material.

    Din listele astea programul genereaza singur retetele: toarna bare din
    loturile de pe primul rand pana cand unul se termina, apoi intra
    automat in urmatorul lot din lista aceluiasi material si deschide o
    reteta noua. Utilizatorul nu mai creeaza retete de mana si nu mai
    alege loturi pe fiecare reteta in parte.

    La OK, self.rezultat contine:
        {"portie", "numarPresari", "nrBare", "target", "loturiAlocate"}
    """

    def __init__(self, order, lots, parent=None, poate_edita=True):
        super().__init__(parent)
        self.setWindowTitle(f"Configurare comanda \u2014 {order.get('nume', '')}")
        self.resize(860, 720)
        self.rezultat = None
        self.poate_edita = poate_edita
        self.lots = lots
        self.tip_aliaj = order.get("tipAliaj", "")
        self.standard = standarde.standard_ales(order)

        prima = (order.get("retete") or [{}])[0]
        target = order.get("target") or prima.get("target") or {}
        alocate = order.get("loturiAlocate") or {}

        radacina = QVBoxLayout(self)

        intro = QLabel(
            "Pune aici, pe fiecare material, <b>loturile in ordinea in care intra "
            "in comanda</b>. Programul toarna barele din primul lot; cand unul se "
            "termina, bara care nu mai incape trece pe reteta urmatoare cu dozarea "
            "veche, iar materialul epuizat continua automat cu urmatorul lot din "
            "lista. Retetele se scriu singure."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
        radacina.addWidget(intro)

        # --- Datele de sarja -------------------------------------------
        grila = QGridLayout()
        self.camp_portie = QLineEdit(str(order.get("portie") or prima.get("portie") or ""))
        self.camp_presari = QLineEdit(str(order.get("numarPresari") or prima.get("numarPresari") or "1"))
        self.camp_bare = QLineEdit(str(order.get("nrBare") or ""))
        for col, (eticheta, camp, latime) in enumerate((
            ("Portie [kg]", self.camp_portie, 90),
            ("Numar de presari", self.camp_presari, 90),
            ("Numar total de bare", self.camp_bare, 90),
        )):
            et = QLabel(eticheta)
            et.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
            grila.addWidget(et, 0, col)
            camp.setStyleSheet(STIL_CAMP)
            camp.setFixedWidth(latime)
            camp.setReadOnly(not poate_edita)
            grila.addWidget(camp, 1, col)
        radacina.addLayout(grila)

        # --- Compozitia tinta ------------------------------------------
        et_tinta = QLabel("Compozitie tinta in bara")
        et_tinta.setStyleSheet("font-weight: 600; font-size: 11px;")
        radacina.addWidget(et_tinta)

        grila_tinta = QGridLayout()
        self.campuri_tinta = {}
        for col, (cheie, eticheta) in enumerate(_elemente_pentru_aliaj(self.tip_aliaj, self.standard)):
            et = QLabel(eticheta)
            et.setStyleSheet(f"color: {CULOARE_GRI_TEXT}; font-size: 11px;")
            grila_tinta.addWidget(et, 0, col)
            camp = QLineEdit(str(target.get(cheie, "")))
            camp.setStyleSheet(STIL_CAMP)
            camp.setFixedWidth(80)
            camp.setReadOnly(not poate_edita)
            grila_tinta.addWidget(camp, 1, col)
            self.campuri_tinta[cheie] = camp
        radacina.addLayout(grila_tinta)

        # --- Cozile de loturi, pe material -----------------------------
        et_loturi = QLabel("Loturi alocate comenzii (in ordinea folosirii)")
        et_loturi.setStyleSheet("font-weight: 600; font-size: 11px;")
        radacina.addWidget(et_loturi)

        zona = QScrollArea()
        zona.setWidgetResizable(True)
        zona.setStyleSheet("QScrollArea { border: none; }")
        interior = QWidget()
        layout_interior = QVBoxLayout(interior)
        layout_interior.setSpacing(8)

        self.liste = {}
        for mat in MATERIALE:
            mid = mat["id"]
            disponibile = [l for l in lots if l.get("material") == mid]
            if not disponibile and not alocate.get(mid):
                continue    # material fara niciun lot in stoc: nu-l aratam

            cadru = QWidget()
            rand = QHBoxLayout(cadru)
            rand.setContentsMargins(0, 0, 0, 0)

            et = QLabel(mat["nume"])
            et.setFixedWidth(130)
            et.setStyleSheet("font-size: 11px;")
            rand.addWidget(et)

            lista = QListWidget()
            lista.setFixedHeight(74)
            lista.setStyleSheet(f"QListWidget {{ border: 1px solid {CULOARE_BORDURA}; "
                                "border-radius: 5px; font-size: 11px; }")
            for lot_id in alocate.get(mid) or []:
                lot = next((l for l in lots if l["id"] == lot_id), None)
                if lot:
                    lista.addItem(self._eticheta_lot(lot))
                    lista.item(lista.count() - 1).setData(Qt.UserRole, lot_id)
            rand.addWidget(lista, 1)

            combo = QComboBox()
            combo.setStyleSheet(STIL_CAMP)
            combo.setFixedWidth(220)
            for lot in disponibile:
                combo.addItem(self._eticheta_lot(lot), lot["id"])
            rand.addWidget(combo)

            coloana_btn = QVBoxLayout()
            coloana_btn.setSpacing(3)
            for text, tooltip, functie in (
                ("+", "Adauga lotul ales la sfarsitul listei",
                 lambda _, m=mid: self._adauga(m)),
                ("\u2191", "Mut lotul selectat mai sus (se foloseste mai devreme)",
                 lambda _, m=mid: self._muta(m, -1)),
                ("\u2193", "Mut lotul selectat mai jos",
                 lambda _, m=mid: self._muta(m, 1)),
                ("\u2715", "Scot lotul selectat din lista",
                 lambda _, m=mid: self._sterge(m)),
            ):
                btn = QPushButton(text)
                btn.setFixedWidth(30)
                btn.setStyleSheet(STIL_BUTON_PERICOL if text == "\u2715" else STIL_BUTON_SECUNDAR)
                btn.setToolTip(tooltip)
                btn.clicked.connect(functie)
                btn.setEnabled(poate_edita)
                coloana_btn.addWidget(btn)
            rand.addLayout(coloana_btn)

            layout_interior.addWidget(cadru)
            self.liste[mid] = {"lista": lista, "combo": combo}

        layout_interior.addStretch(1)
        zona.setWidget(interior)
        radacina.addWidget(zona, 1)

        # --- Butoane ----------------------------------------------------
        rand_btn = QHBoxLayout()
        btn_renunta = QPushButton("Renunta")
        btn_renunta.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_renunta.clicked.connect(self.reject)
        rand_btn.addWidget(btn_renunta)
        rand_btn.addStretch(1)
        if poate_edita:
            btn_ok = QPushButton("Salveaza si genereaza retetele")
            btn_ok.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn_ok.setToolTip(
                "Salveaza configurarea si construieste automat toate retetele "
                "comenzii din listele de loturi de mai sus."
            )
            btn_ok.clicked.connect(self._accepta)
            rand_btn.addWidget(btn_ok)
        radacina.addLayout(rand_btn)

    # -- helpere ---------------------------------------------------------
    @staticmethod
    def _eticheta_lot(lot):
        nume = lot.get("nrLot") or lot.get("lot") or lot.get("id")
        return f"{nume}  \u2014  rest {fmt(lot_rest(lot), 2)} kg"

    def _lot_ales(self, mid):
        combo = self.liste[mid]["combo"]
        return combo.currentData()

    def _adauga(self, mid):
        lot_id = self._lot_ales(mid)
        if not lot_id:
            return
        lista = self.liste[mid]["lista"]
        existente = [lista.item(i).data(Qt.UserRole) for i in range(lista.count())]
        if lot_id in existente:
            QMessageBox.information(self, "Lot deja in lista",
                                    "Lotul ales e deja in lista acestui material.")
            return
        lot = next((l for l in self.lots if l["id"] == lot_id), None)
        if lot is None:
            return
        lista.addItem(self._eticheta_lot(lot))
        lista.item(lista.count() - 1).setData(Qt.UserRole, lot_id)

    def _muta(self, mid, directie):
        lista = self.liste[mid]["lista"]
        i = lista.currentRow()
        j = i + directie
        if i < 0 or j < 0 or j >= lista.count():
            return
        item = lista.takeItem(i)
        lista.insertItem(j, item)
        lista.setCurrentRow(j)

    def _sterge(self, mid):
        lista = self.liste[mid]["lista"]
        i = lista.currentRow()
        if i >= 0:
            lista.takeItem(i)

    def _accepta(self):
        nr_bare = int(to_float(self.camp_bare.text()))
        if nr_bare <= 0:
            QMessageBox.warning(self, "Numar de bare",
                                "Scrie cate bare are comanda (numar intreg, mai mare ca 0).")
            return
        if to_float(self.camp_portie.text()) <= 0:
            QMessageBox.warning(self, "Portie", "Scrie portia in kg (mai mare ca 0).")
            return

        alocate = {}
        for mid, widgets in self.liste.items():
            lista = widgets["lista"]
            ids = [lista.item(i).data(Qt.UserRole) for i in range(lista.count())]
            if ids:
                alocate[mid] = ids
        if not alocate:
            QMessageBox.warning(
                self, "Fara loturi",
                "Adauga cel putin lotul de burete si loturile materialelor de aliere, "
                "in ordinea in care intra in comanda."
            )
            return

        target = {k: to_float(c.text()) for k, c in self.campuri_tinta.items()}
        incalcari = standarde.verifica_limite(self.standard, target)
        if incalcari:
            raspuns = QMessageBox.question(
                self, "Tinta iese din standard",
                "Compozitia tinta nu se incadreaza in standardul comenzii:\n\n"
                + "\n".join("\u2022 " + i["text"] for i in incalcari)
                + "\n\nSalvezi oricum?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if raspuns != QMessageBox.Yes:
                return

        self.rezultat = {
            "portie": self.camp_portie.text().strip(),
            "numarPresari": self.camp_presari.text().strip() or "1",
            "nrBare": str(nr_bare),
            "target": target,
            "loturiAlocate": alocate,
        }
        self.accept()


class DialogPreviewGenerare(QDialog):
    """Previzualizarea planului generat automat din loturile comenzii,
    inainte sa fie scris efectiv pe comanda.

    Utilizatorul VALIDEAZA planul (\u201eConfirma si salveaza\u201d) sau renunta;
    nimic nu se schimba pe comanda pana la confirmare. Daca in timpul
    generarii s-a terminat lista de loturi a vreunui material, planul e
    aratat oricum (partial \u2014 pana unde a ajuns) impreuna cu avertismentul,
    si operatorul poate deschide direct configurarea ca sa adauge lotul
    care lipseste.
    """

    def __init__(self, retete_noi, raport, pastrate, parent=None, poate_edita=True):
        super().__init__(parent)
        self.setWindowTitle("Verifica planul generat")
        self.resize(620, 520)
        self.actiune = None   # "confirma" | "configureaza" | None (renunta)

        radacina = QVBoxLayout(self)

        titlu = QLabel(
            f"S-au generat <b>{len(retete_noi)}</b> retete noi, "
            f"<b>{raport['barePlasate']}</b> bare plasate in total"
            + (f", pastrand primele <b>{pastrate}</b> retete (consum deja aplicat)"
               if pastrate else "") + "."
        )
        titlu.setWordWrap(True)
        radacina.addWidget(titlu)

        tabel = QTableWidget()
        tabel.setColumnCount(4)
        tabel.setHorizontalHeaderLabels(["Reteta", "Bare", "Tip bare", "Loturi folosite"])
        tabel.setEditTriggers(QTableWidget.NoEditTriggers)
        tabel.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        tabel.setRowCount(len(retete_noi))
        for i, r in enumerate(retete_noi):
            tipuri = [b.get("tipDozare") for b in r["bare"]]
            nr_report = tipuri.count("report")
            tip_text = (
                (f"{nr_report} report + {len(r['bare']) - nr_report} noi" if nr_report
                 else f"{len(r['bare'])} noi")
            )
            loturi_text = ", ".join(
                f"{nume_material(mid)}: {lid}" for mid, lid in r["lotSel"].items()
            )
            tabel.setItem(i, 0, QTableWidgetItem(r["nume"]))
            tabel.setItem(i, 1, QTableWidgetItem(str(len(r["bare"]))))
            tabel.setItem(i, 2, QTableWidgetItem(tip_text))
            tabel.setItem(i, 3, QTableWidgetItem(loturi_text))
        radacina.addWidget(tabel, 1)

        if raport["bareRamase"]:
            avert = QLabel(
                f"\u26a0 Mai raman <b>{raport['bareRamase']}</b> bare neplasate:"
            )
            avert.setStyleSheet(f"color: {CULOARE_EROARE};")
            radacina.addWidget(avert)
        for text in raport["avertismente"]:
            et = QLabel("\u2022 " + text)
            et.setWordWrap(True)
            et.setStyleSheet(f"color: {CULOARE_EROARE}; font-size: 11px;")
            radacina.addWidget(et)

        rand_btn = QHBoxLayout()
        btn_renunta = QPushButton("Renunta")
        btn_renunta.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_renunta.clicked.connect(self.reject)
        rand_btn.addWidget(btn_renunta)
        rand_btn.addStretch(1)
        if raport["avertismente"] and poate_edita:
            btn_config = QPushButton("Deschide configurarea si adauga lotul")
            btn_config.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_config.clicked.connect(self._configureaza)
            rand_btn.addWidget(btn_config)
        if poate_edita:
            btn_ok = QPushButton("\u2713 Confirma si salveaza")
            btn_ok.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn_ok.setEnabled(len(retete_noi) > 0)
            btn_ok.clicked.connect(self._confirma)
            rand_btn.addWidget(btn_ok)
        radacina.addLayout(rand_btn)

    def _confirma(self):
        self.actiune = "confirma"
        self.accept()

    def _configureaza(self):
        self.actiune = "configureaza"
        self.reject()
