"""Ferestre de dialog: autentificare admin, adaugare/editare lot, istoric lot.

Daca un aliaj nou are elemente in plus in compozitia unui lot (ex. Si, Zr,
Mo pentru Ti-VT9), locul de adaugat campurile noi e in DialogLotNou,
variabila `etichete_doza` din __init__.
"""

from datetime import date

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QGridLayout, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from .config import MATERIALE
from . import conturi as conturi_mod
from .stiluri import STIL_BUTON_PRINCIPAL, STIL_BUTON_SECUNDAR, STIL_BUTON_PERICOL, STIL_CAMP, CULOARE_EROARE, CULOARE_SUCCES, CULOARE_GRI_TEXT, CULOARE_BORDURA, CULOARE_FUNDAL_SECTIUNE
from .utils import fmt, to_float
from .calcule import lot_ti, lot_rest


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
