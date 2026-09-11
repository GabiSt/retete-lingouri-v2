"""Helpere mici, fara dependinte de UI sau de logica de business."""

import re
import uuid


def uid():
    """Genereaza un identificator scurt si unic (10 caractere hex)."""
    return uuid.uuid4().hex[:10]


def fmt(n, d=6):
    """Formateaza un numar cu d zecimale; returneaza '\u2013' pentru None/invalid."""
    try:
        if n is None:
            return "\u2013"
        return f"{float(n):.{d}f}"
    except (TypeError, ValueError):
        return "\u2013"


def numar_comanda(order):
    """Extrage doar numarul comenzii (ex. "26851") din numele complet al
    comenzii (ex. "Cda 26851-Ti5-2VAR-d600-L1-L3"), asa cum apare pe
    formularele tiparite (celula "Comanda", titlul "FISA LIMITA cda ...").
    Daca numele nu contine nicio secventa de cifre, se foloseste numele
    intreg ca fallback."""
    nume = (order.get("nume") or "").strip()
    m = re.search(r"\d+", nume)
    return m.group(0) if m else nume


def pct(v, zecimale=2):
    """Formateaza o fractie (0.002) ca text procentual ("0.20%"), cu numarul
    de zecimale dat. Returneaza sir gol pentru None/valoare invalida (nu
    "-", ca sa ramana celule goale pe formular, nu pline de liniute)."""
    try:
        if v is None or v == "":
            return ""
        return f"{float(v) * 100:.{zecimale}f}%"
    except (TypeError, ValueError):
        return ""


def to_float(v, implicit=0.0):
    """Converteste sigur la float; returneaza 'implicit' daca nu se poate."""
    try:
        if v is None or v == "":
            return implicit
        return float(v)
    except (TypeError, ValueError):
        return implicit
