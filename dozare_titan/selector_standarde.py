"""Control de selectie MULTIPLA a standardelor unei comenzi.

Un QComboBox nu poate avea mai multe optiuni bifate, asa ca standardele se
aleg dintr-un buton cu meniu de casute bifabile. Meniul RAMANE DESCHIS cat
timp bifezi (poti alege AMS 4975 + AMS 4976 dintr-o singura deschidere) si se
inchide la click in afara lui.

ATENTIE la motivul real pentru care meniul se putea inchide dupa o singura
bifare: nu era o problema de QMenu/QCheckBox, ci faptul ca semnalul
``schimbat`` se emitea la FIECARE bifare in parte. In fereastra principala,
acel semnal duce direct la ``_rebuild_retete()``, care sterge si reconstruieste
TOT panoul de comenzi (deci si acest buton, cu meniul lui) — ceea ce inchide
vizual meniul imediat dupa primul click, indiferent cat de bine se comporta
QMenu la nivel de widget. Solutia: semnalul se emite o SINGURA DATA, cand
meniul se inchide (nu la fiecare bifare individuala) — abia atunci se declanseaza
reconstructia din fereastra principala, dupa ce utilizatorul a terminat de ales.

Modulul depinde doar de PySide6 (nu de restul aplicatiei), ca sa poata fi
verificat separat.
"""

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QCheckBox, QMenu, QToolButton, QWidgetAction


class _MeniuFaraInchidere(QMenu):
    """QMenu care NU se inchide la click pe o casuta bifabila din el.

    Strat suplimentar de siguranta: desi fiecare casuta e widget-ul unui
    QWidgetAction (nu o actiune bifabila obisnuita), pe unele platforme/stiluri
    QMenu tot poate inchide meniul la eliberarea click-ului deasupra unei
    actiuni, chiar si cand acea actiune are un widget custom. Interceptam:
    daca actiunea activa e un QWidgetAction cu o QCheckBox, bifam/debifam
    manual caseta si oprim evenimentul (fara sa mai apelam implementarea de
    baza, care ar inchide meniul). Click-ul in afara meniului nu trece pe
    aici, deci inchide meniul normal, ca pana acum.
    """

    def mouseReleaseEvent(self, event):
        actiune = self.activeAction()
        if isinstance(actiune, QWidgetAction):
            widget = actiune.defaultWidget()
            if isinstance(widget, QCheckBox) and widget.isEnabled():
                widget.toggle()
                event.accept()
                return
        super().mouseReleaseEvent(event)


class SelectorStandarde(QToolButton):
    """Buton cu meniu de casute bifabile.

    ``schimbat(list[str])`` se emite cu numele bifate (in ordinea din lista
    de standarde, nu in ordinea bifarii) — dar DOAR cand meniul se inchide
    (click in afara lui), NU la fiecare bifare individuala. Asta permite sa
    bifezi mai multe standarde dintr-o singura deschidere, fara ca vreun
    ascultator extern (care ar reconstrui interfata la fiecare schimbare,
    cum face fereastra principala) sa inchida meniul dupa primul click.
    Textul butonului insa se actualizeaza LIVE, la fiecare bifare, ca sa vezi
    imediat ce ai selectat.
    """

    schimbat = Signal(list)

    def __init__(self, nume_standarde, alese=(), text_gol="\u2014 alege standardele \u2014",
                 parent=None):
        super().__init__(parent)
        self._nume = list(nume_standarde)
        self._text_gol = text_gol
        self._casute = {}
        self._menu = _MeniuFaraInchidere(self)
        for nume in self._nume:
            casuta = QCheckBox(nume)
            casuta.setChecked(nume in set(alese))
            casuta.setStyleSheet("QCheckBox { padding: 4px 10px; }")
            casuta.toggled.connect(self._la_bifare)
            actiune = QWidgetAction(self._menu)
            actiune.setDefaultWidget(casuta)        # meniul nu se inchide la bifare
            self._menu.addAction(actiune)
            self._casute[nume] = casuta
        self.setMenu(self._menu)
        self.setPopupMode(QToolButton.InstantPopup)
        self._alese_la_deschidere = self.alese()
        self._menu.aboutToShow.connect(self._la_deschidere)
        self._menu.aboutToHide.connect(self._la_inchidere)
        self._actualizeaza_text()

    def alese(self):
        """Numele bifate, in ordinea listei de standarde."""
        return [n for n in self._nume if self._casute[n].isChecked()]

    def seteaza_alese(self, nume):
        """Bifeaza exact standardele date, FARA sa emita semnalul."""
        nume = set(nume)
        for n, casuta in self._casute.items():
            casuta.blockSignals(True)
            casuta.setChecked(n in nume)
            casuta.blockSignals(False)
        self._alese_la_deschidere = self.alese()
        self._actualizeaza_text()

    def _la_bifare(self, _stare):
        """O casuta s-a bifat/debifat: doar textul butonului se actualizeaza
        acum (feedback vizual imediat). Semnalul catre restul aplicatiei NU
        pleaca de aici — vezi ``_la_inchidere`` si explicatia din docstring-ul
        clasei."""
        self._actualizeaza_text()

    def _la_deschidere(self):
        """Meniul tocmai s-a deschis: retinem selectia de plecare, ca sa
        stim la inchidere daca s-a schimbat ceva cu adevarat."""
        self._alese_la_deschidere = self.alese()

    def _la_inchidere(self):
        """Meniul tocmai s-a inchis (click in afara lui): emitem semnalul o
        SINGURA DATA, si doar daca selectia chiar s-a schimbat fata de cand
        s-a deschis meniul — ca sa nu declansam o reconstructie a interfetei
        degeaba, doar pentru ca utilizatorul a deschis si inchis meniul fara
        sa bifeze nimic."""
        alese = self.alese()
        if alese != self._alese_la_deschidere:
            self.schimbat.emit(alese)

    def _actualizeaza_text(self):
        alese = self.alese()
        if not alese:
            text = self._text_gol
        elif len(alese) == 1:
            text = alese[0]
        else:
            text = f"{alese[0]} (+{len(alese) - 1})"
        self.setText(text + "  \u25be")
