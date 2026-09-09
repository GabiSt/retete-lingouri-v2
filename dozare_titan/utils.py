"""Helpere mici, fara dependinte de UI sau de logica de business."""

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


def to_float(v, implicit=0.0):
    """Converteste sigur la float; returneaza 'implicit' daca nu se poate."""
    try:
        if v is None or v == "":
            return implicit
        return float(v)
    except (TypeError, ValueError):
        return implicit
