"""Incarcare/salvare a starii aplicatiei (loturi + comenzi) in date.json,
plus migratii pentru date salvate cu o structura mai veche."""

import json
import os

from .config import DATA_DIR, DATA_FILE, ALIAJ_IMPLICIT


def _numar_bare_anterioare(order, recipe_id):
    """Cate bare exista, in total, in retetele comenzii INAINTE de reteta
    recipe_id (in ordinea lor din order['retete']). Folosit ca sa numerotam
    barele continuu pe toata comanda (1,2,3 pe prima reteta, 4,5,6,7 pe
    urmatoarea etc.), nu separat pe fiecare reteta."""
    total = 0
    for r in order.get("retete", []):
        if r["id"] == recipe_id:
            break
        total += len(r.get("bare", []))
    return total


def _migreaza_reteta(r):
    """Completeaza cu valori implicite cheile care lipsesc din retetele salvate
    inainte de introducerea campurilor 'portie'/'numarPresari' la nivel de reteta,
    ca sa nu mai apara KeyError la incarcarea unor date mai vechi."""
    r.setdefault("portie", "")
    r.setdefault("numarPresari", "1")
    r.setdefault("bare", [])
    for bar in r["bare"]:
        bar.pop("portie", None)
        bar.pop("nrPresari", None)
        bar.setdefault("consumApplied", False)
        bar.setdefault("calcSnapshot", None)
    return r


def _migreaza_comanda(order):
    """Completeaza tipul de aliaj pentru comenzile salvate inainte de
    introducerea acestui camp — toate au fost facute pentru Ti6Al4V."""
    order.setdefault("tipAliaj", ALIAJ_IMPLICIT)
    return order


def incarca_date():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                orders = d.get("orders", [])
                for order in orders:
                    _migreaza_comanda(order)
                    for r in order.get("retete", []):
                        _migreaza_reteta(r)
                return {"lots": d.get("lots", []), "orders": orders}
        except Exception:
            pass
    return {"lots": [], "orders": []}


def salveaza_date(state):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
