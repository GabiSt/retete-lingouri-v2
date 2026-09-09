#!/usr/bin/env python3
"""Dozare Lingouri Titan — aplicatie desktop (PySide6).

Gestionare stoc de loturi pe materiale si calcul de retete de sarja
(compozitie chimica tinta -> cantitati per material).

Ruleaza cu:
    python3 dozare_titan.py
"""

import json
import os
import sys
import uuid
from datetime import date

import numpy as np
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFrame, QGridLayout, QHBoxLayout,
    QInputDialog, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea,
    QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
    QFileDialog, QPlainTextEdit,
)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    OPENPYXL_DISPONIBIL = True
except ImportError:
    OPENPYXL_DISPONIBIL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    REPORTLAB_DISPONIBIL = True
except ImportError:
    REPORTLAB_DISPONIBIL = False

# ---------------------------------------------------------------------------
# Date fixe
# ---------------------------------------------------------------------------
MATERIALE = [
    {"id": "burete", "nume": "Burete de titan"},
    {"id": "aliajAlV", "nume": "Aliaj Al-V"},
    {"id": "alMetal", "nume": "Al metal"},
    {"id": "feMetal", "nume": "Fe metal"},
    {"id": "tio2", "nume": "TiO\u2082"},
]
ORDINE_MAT = [m["id"] for m in MATERIALE]
ADMIN_PASS = "titan2026"

DATA_DIR = os.path.join(os.path.expanduser("~"), ".dozare_titan")
DATA_FILE = os.path.join(DATA_DIR, "date.json")

# Etichetele materialelor asa cum apar pe formularul tiparit "Fisa limita -cda"
# (identice cu formularul pe hartie, inclusiv denumirea repetata "Aluminiu").
FISA_LIMITA_ETICHETE = {
    "burete": "Burete Ti",
    "aliajAlV": "Prealiaj AlV-GfE",
    "alMetal": "Aluminiu",
    "feMetal": "Aluminiu",
    "tio2": "TiO2",
}

# ---------------------------------------------------------------------------
# Tipuri de aliaj — calculul de dozare/RetDozare e valabil per tip de aliaj.
# Comenzile de pana acum au fost toate pentru Ti6Al4V; cand apar alte tipuri,
# se adauga o noua intrare aici (cu limitele ei chimice) si un generator
# RetDozare dedicat daca formatul difera.
# ---------------------------------------------------------------------------
ALIAJ_IMPLICIT = "Ti6Al4V"
ALIAJE_SPEC = {
    "Ti6Al4V": {
        "nume": "Ti-6Al-4V ELI",
        "o_max": 0.002,
        "fe_max": 0.003,
        "n_max": 0.0005,
        "c_max": 0.0008,
        "h_max": 0.000125,
        "al_min": 0.055,
        "al_max": 0.0675,
        "v_min": 0.035,
        "v_max": 0.045,
    },
    "Ti -VT9":{
        "nume": "Ti -VT9",
        "o_max": 0.015,
        "si_min": 0.20,
        "si_max": 0.35,
        "fe_max":0.10,
        "zr_min": 1.0,
        "zr_max": 2.0,
        "mo_min": 2.8,
        "mo_max": 3.8,
        "al_min": 5.8,
        "al_max":7.0,   
    },
}
ALIAJE_DISPONIBILE = list(ALIAJE_SPEC.keys())

# ---------------------------------------------------------------------------
# Stil vizual
# ---------------------------------------------------------------------------
CULOARE_BLEUMARIN = "#1f3b73"
CULOARE_EROARE = "#c0392b"
CULOARE_SUCCES = "#1f8a55"
CULOARE_AVERTISMENT = "#c98a1d"
CULOARE_GRI_TEXT = "#6b7280"
CULOARE_FUNDAL_SECTIUNE = "#f4f6f8"
CULOARE_BORDURA = "#d8dee3"

STIL_BUTON_PRINCIPAL = f"""
    QPushButton {{
        background-color: {CULOARE_BLEUMARIN};
        color: white;
        font-weight: 600;
        padding: 7px 14px;
        border-radius: 6px;
        border: none;
    }}
    QPushButton:hover {{ background-color: #16294f; }}
    QPushButton:disabled {{ background-color: #9aa7b8; }}
"""

STIL_BUTON_SECUNDAR = f"""
    QPushButton {{
        background-color: white;
        color: {CULOARE_BLEUMARIN};
        font-weight: 600;
        padding: 7px 14px;
        border-radius: 6px;
        border: 1px solid {CULOARE_BLEUMARIN};
    }}
    QPushButton:hover {{ background-color: #eef2f8; }}
"""

STIL_BUTON_PERICOL = f"""
    QPushButton {{
        background: transparent;
        color: {CULOARE_EROARE};
        border: none;
        font-weight: 700;
        padding: 2px 8px;
    }}
    QPushButton:hover {{ background-color: #fbeceb; border-radius: 4px; }}
"""

STIL_CAMP = f"""
    QLineEdit, QComboBox {{
        border: 1px solid {CULOARE_BORDURA};
        border-radius: 5px;
        padding: 5px 7px;
        background: white;
    }}
    QLineEdit:focus, QComboBox:focus {{ border: 1px solid {CULOARE_BLEUMARIN}; }}
"""

TAG_STYLES = {
    "good": f"color:{CULOARE_SUCCES}; background-color:#e5f5ec; border-radius:5px; padding:2px 8px; font-weight:600;",
    "warn": f"color:{CULOARE_AVERTISMENT}; background-color:#faf1de; border-radius:5px; padding:2px 8px; font-weight:600;",
    "bad": f"color:{CULOARE_EROARE}; background-color:#fbeceb; border-radius:5px; padding:2px 8px; font-weight:600;",
}


def _clear_layout(layout):
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()


# ---------------------------------------------------------------------------
# Helpers de date / calcul
# ---------------------------------------------------------------------------
def uid():
    return uuid.uuid4().hex[:10]


def fmt(n, d=6):
    try:
        if n is None:
            return "\u2013"
        return f"{float(n):.{d}f}"
    except (TypeError, ValueError):
        return "\u2013"


def to_float(v, implicit=0.0):
    try:
        if v is None or v == "":
            return implicit
        return float(v)
    except (TypeError, ValueError):
        return implicit


def lot_ti(lot):
    """
    Calculeaza procentul de Titan din lot.
    Doar Burete Ti si TiO2 contin Ti.
    """
    material = lot.get("material", "")
    
    # Burete Ti: Ti = 100 - Al - V - O - Fe - N
    if material == "burete":
        return 100 - (
            to_float(lot.get("dozaAl")) + to_float(lot.get("dozaV")) +
            to_float(lot.get("dozaO")) + to_float(lot.get("dozaFe")) +
            to_float(lot.get("dozaN"))
        )
    
    # TiO2: are ~60% Ti (40% O)
    elif material == "tio2":
        return 60.0
    
    # Aliaj AlV: nu contine Ti (doar Al si V)
    elif material == "aliajAlV":
        return 0.0
    
    # Al metal: nu contine Ti
    elif material == "alMetal":
        return 0.0
    
    # Fe metal: nu contine Ti
    elif material == "feMetal":
        return 0.0
    
    # Default
    return 0.0


def lot_rest(lot):
    return to_float(lot.get("stocIntrare")) - to_float(lot.get("consum"))


def target_ti(target):
    return 100 - (
        to_float(target.get("al")) + to_float(target.get("v")) +
        to_float(target.get("o")) + to_float(target.get("fe")) +
        to_float(target.get("n", 0))
    )


# Materialul din stanga se calculeaza din balanta elementului din dreapta.
# Daca tinta acelui element e 0%, materialul nu mai e necesar in amestec —
# nu trebuie selectat lot pentru el, e pur si simplu ignorat.
ELEMENT_PER_MATERIAL = {
    "aliajAlV": "v",
    "alMetal": "al",
    "tio2": "o",
    "feMetal": "fe",
}


def material_necesar(mat_id, target):
    """True daca materialul chiar trebuie sa aiba un lot selectat.
    Burete de titan e mereu necesar (e incarcatura de baza a retetei);
    celelalte materiale sunt necesare doar daca tinta elementului lor
    asociat e diferita de 0%."""
    elem = ELEMENT_PER_MATERIAL.get(mat_id)
    if elem is None:
        return True
    return abs(to_float(target.get(elem))) > 1e-9


def _componente_loturi(lot_sel, target):
    """Extrage compozitia (%) fiecarui lot selectat, indexata dupa material.
    Materialele care nu sunt necesare (tinta elementului asociat = 0%) nu
    au nevoie de lot selectat — se trateaza cu compozitie zero si sunt
    ignorate din calcul."""
    comps = {}
    for k in ORDINE_MAT:
        l = lot_sel.get(k)
        if not l:
            if material_necesar(k, target):
                return None, f"Selecteaza un lot pentru {k}."
            comps[k] = {"Al": 0.0, "V": 0.0, "O": 0.0, "Fe": 0.0, "Ti": 0.0}
            continue
        comps[k] = {
            "Al": to_float(l.get("dozaAl")),
            "V": to_float(l.get("dozaV")),
            "O": to_float(l.get("dozaO")),
            "Fe": to_float(l.get("dozaFe")),
            "Ti": lot_ti(l),
        }
    return comps, None


def calculeaza_bara(target, lot_sel, portie, nr_presari=1):
    """
    Calculeaza dozarea EXACT ca in Excel (foaia "Retete"), pe un lant de
    formule secventiale — nu prin rezolvarea unui sistem 3x3 cu solutie unica.

    Ordinea de calcul (identica cu Excel):
      1. Aliaj Al-V se calculeaza din balanta de V (singurul material cu V).
      2. Al metal se calculeaza din balanta de Al, dupa ce se scade
         contributia de Al din Aliajul Al-V.
      3. Burete Ti = portia de baza - Aliaj Al-V - Al metal (balanta de masa
         a incarcaturii principale).
      4. TiO2 se adauga suplimentar pentru a acoperi balanta de O.
      5. Fe metal se adauga suplimentar pentru a acoperi balanta de Fe.

    TiO2 si Fe metal sunt doze mici, adaugate PESTE incarcatura principala
    (ca in Excel), deci masa finala a barei poate fi usor peste "portie".
    Aceasta metoda are mereu o solutie (atata timp cat loturile selectate au
    %V in Aliaj Al-V, %Al in Al metal, %O in TiO2 si %Fe in Fe metal diferite
    de zero) — nu mai exista cazul de "sistem fara solutie unica".

    "nr_presari" e numarul de presari (ca in Excel, coloana "Nr.Pres."):
    cantitatile de materiale calculate pentru o portie se inmultesc cu acest
    numar pentru a obtine cantitatea TOTALA consumata din stoc pentru bara
    respectiva (o bara poate necesita mai multe presari din aceeasi dozare).
    """
    comps, eroare = _componente_loturi(lot_sel, target)
    if eroare:
        return {"eroare": eroare}

    p = to_float(portie)
    if p <= 0:
        return {"eroare": "Introdu o portie valida (kg)."}

    n = to_float(nr_presari, 1)
    if n <= 0:
        n = 1

    rezultat, eroare = _calculeaza_dozare_kg(target, comps, p)
    if eroare:
        return {"eroare": eroare}

    return _rezumat_calcul(target, comps, lot_sel, rezultat, p, n)


def _calculeaza_dozare_kg(target, comps, p):
    """Aplica formulele Excel pentru o portie de p kg. Returneaza (rezultat, eroare)."""
    v_tinta_kg = to_float(target.get("v")) / 100 * p
    al_tinta_kg = to_float(target.get("al")) / 100 * p
    o_tinta_kg = to_float(target.get("o")) / 100 * p
    fe_tinta_kg = to_float(target.get("fe")) / 100 * p

    # 1. Aliaj Al-V din balanta de V (ignorat daca tinta V = 0%)
    if abs(v_tinta_kg) < 1e-9:
        aliaj_kg = 0.0
    else:
        v_alv = comps["aliajAlV"]["V"]
        if abs(v_alv) < 1e-9:
            return None, "Lotul de Aliaj Al-V nu are %V definit — nu se poate calcula."
        aliaj_kg = v_tinta_kg / (v_alv / 100)

    # 2. Al metal din balanta de Al, dupa ce scadem contributia Aliajului
    #    Al-V (ignorat daca tinta Al = 0%)
    al_din_alv = aliaj_kg * comps["aliajAlV"]["Al"] / 100
    if abs(al_tinta_kg) < 1e-9:
        al_metal_kg = 0.0
    else:
        al_am = comps["alMetal"]["Al"]
        if abs(al_am) < 1e-9:
            return None, "Lotul de Al metal nu are %Al definit — nu se poate calcula."
        al_metal_kg = (al_tinta_kg - al_din_alv) / (al_am / 100)

    # 3. Burete Ti = restul incarcaturii principale
    burete_kg = p - aliaj_kg - al_metal_kg

    # 4. TiO2 adaugat suplimentar, din balanta de O (ignorat daca tinta O = 0%)
    o_din_burete = burete_kg * comps["burete"]["O"] / 100
    o_din_alv = aliaj_kg * comps["aliajAlV"]["O"] / 100
    o_din_al = al_metal_kg * comps["alMetal"]["O"] / 100
    if abs(o_tinta_kg) < 1e-9:
        tio2_kg = 0.0
    else:
        o_tio2 = comps["tio2"]["O"]
        if abs(o_tio2) < 1e-9:
            return None, "Lotul de TiO2 nu are %O definit — nu se poate calcula."
        tio2_kg = (o_tinta_kg - o_din_burete - o_din_alv - o_din_al) / (o_tio2 / 100)

    # 5. Fe metal adaugat suplimentar, din balanta de Fe (ignorat daca tinta Fe = 0%)
    fe_din_burete = burete_kg * comps["burete"]["Fe"] / 100
    fe_din_alv = aliaj_kg * comps["aliajAlV"]["Fe"] / 100
    fe_din_al = al_metal_kg * comps["alMetal"]["Fe"] / 100
    if abs(fe_tinta_kg) < 1e-9:
        fe_metal_kg = 0.0
    else:
        fe_fe = comps["feMetal"]["Fe"]
        if abs(fe_fe) < 1e-9:
            return None, "Lotul de Fe metal nu are %Fe definit — nu se poate calcula."
        fe_metal_kg = (fe_tinta_kg - fe_din_burete - fe_din_alv - fe_din_al) / (fe_fe / 100)

    rezultat = {
        "burete": burete_kg,
        "aliajAlV": aliaj_kg,
        "alMetal": al_metal_kg,
        "tio2": tio2_kg,
        "feMetal": fe_metal_kg,
    }
    return rezultat, None


def _rezumat_calcul(target, comps, lot_sel, rezultat, p, n=1):
    """Construieste dictionarul de rezultat afisat in UI, plecand de la cantitatile pe material.

    "rezultat" e dozarea pentru O SINGURA presare; "rezultatTotal" e dozarea
    inmultita cu numarul de presari (n) — aceasta e cantitatea REALA care se
    scade din stoc pentru bara respectiva.
    """
    rezultat_total = {k: v * n for k, v in rezultat.items()}

    negativ = any(rezultat_total[k] < -0.005 for k in rezultat_total)

    depaseste_stoc = False
    for k in ORDINE_MAT:
        lot = lot_sel.get(k)
        if lot and rezultat_total[k] > lot_rest(lot) + 0.001:
            depaseste_stoc = True

    masa_totala = sum(rezultat.values())

    ti_rezultat = 0
    al_rezultat = 0
    v_rezultat = 0
    o_rezultat = 0
    fe_rezultat = 0
    for k in ORDINE_MAT:
        ti_rezultat += rezultat[k] * comps[k]["Ti"] / 100
        al_rezultat += rezultat[k] * comps[k]["Al"] / 100
        v_rezultat += rezultat[k] * comps[k]["V"] / 100
        o_rezultat += rezultat[k] * comps[k]["O"] / 100
        fe_rezultat += rezultat[k] * comps[k]["Fe"] / 100

    divizor = masa_totala if masa_totala > 1e-9 else 1
    return {
        "rezultat": rezultat,
        "rezultatTotal": rezultat_total,
        "nrPresari": n,
        "negativ": negativ or depaseste_stoc,
        "tiRezultat": ti_rezultat / divizor * 100,
        "tiTarget": target_ti(target),
        "portie": p,
        "portieEfectiva": masa_totala,
        "portieTotala": masa_totala * n,
        "compozitie_rezultata": {
            "Al": al_rezultat / divizor * 100,
            "V": v_rezultat / divizor * 100,
            "O": o_rezultat / divizor * 100,
            "Fe": fe_rezultat / divizor * 100,
            "Ti": ti_rezultat / divizor * 100,
        }
    }




def _numar_bare_anterioare(order, recipe_id):
    """Cate bare exista, in total, in retetele comenzii AINTE de reteta
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


# ---------------------------------------------------------------------------
# Generare "Fisa limita -cda" (.xlsx + raport .pdf), pe baza consumului REAL
# ---------------------------------------------------------------------------
def gaseste_istoric_lot(lot_id, orders):
    """Gaseste toate comenzile/retetele/barele in care lotul lot_id a fost
    folosit REAL (consum aplicat pe stoc, prin butonul \"Aplica consumul pe
    stoc\"), cu cantitatea consumata din acest lot in fiecare bara.

    Foloseste exact aceeasi sursa de adevar ca si calculeaza_consum_comanda
    (Fisa limita): consumSnapshot-ul inregistrat la aplicare, nicio
    recalculare — deci reflecta EXACT ce s-a intamplat, indiferent daca
    reteta/lotul selectat s-au schimbat ulterior.

    Returneaza o lista de dict-uri (comanda, data, reteta, bara, material,
    kg), sortata pe nume de comanda si numarul barei.
    """
    istoric = []
    for order in orders:
        for r in order.get("retete", []):
            offset_bare = _numar_bare_anterioare(order, r["id"])
            for i, bar in enumerate(r.get("bare", [])):
                if not bar.get("consumApplied"):
                    continue
                snapshot = bar.get("consumSnapshot") or {}
                for mat_id, parti in snapshot.items():
                    intrari = parti if isinstance(parti, list) else [parti]
                    for intrare in intrari:
                        if intrare.get("lotId") != lot_id:
                            continue
                        kg = to_float(intrare.get("kg"))
                        if abs(kg) < 1e-9:
                            continue
                        istoric.append({
                            "comanda": order.get("nume", ""),
                            "data": order.get("data", ""),
                            "reteta": r.get("nume", ""),
                            "bara": offset_bare + i + 1,
                            "material": mat_id,
                            "kg": kg,
                        })
    istoric.sort(key=lambda x: (x["comanda"], x["bara"]))
    return istoric


def calculeaza_consum_comanda(order, lots):
    """Aduna consumul efectiv (deja aplicat pe stoc, prin butonul \"Aplica
    consumul pe stoc\") al comenzii, pe material si pe lot.

    Materialele fixe din MATERIALE sunt doar orientative pentru formularul
    tiparit; cantitatile din fisa limita trebuie sa reflecte STRICT ce s-a
    consumat efectiv din loturi, nu dozarea teoretica sau targetul retetei.

    Returneaza dict material_id -> {"loturi": [(nrLot, kg), ...], "total": kg}.
    """
    lots_by_id = {l["id"]: l for l in lots}
    agregat = {m["id"]: {} for m in MATERIALE}  # material_id -> {lotId: kg}

    for r in order.get("retete", []):
        for bar in r.get("bare", []):
            if not bar.get("consumApplied"):
                continue
            snapshot = bar.get("consumSnapshot") or {}
            for mat_id, parti in snapshot.items():
                if mat_id not in agregat:
                    continue
                intrari = parti if isinstance(parti, list) else [parti]
                for intrare in intrari:
                    lot_id = intrare.get("lotId")
                    kg = to_float(intrare.get("kg"))
                    if lot_id is None or abs(kg) < 1e-9:
                        continue
                    agregat[mat_id][lot_id] = agregat[mat_id].get(lot_id, 0.0) + kg

    rezultat = {}
    for mat_id, per_lot in agregat.items():
        loturi = []
        total = 0.0
        for lot_id, kg in per_lot.items():
            lot = lots_by_id.get(lot_id)
            nr_lot = lot.get("nrLot", "?") if lot else "?"
            loturi.append((nr_lot, kg))
            total += kg
        loturi.sort(key=lambda item: item[0])
        rezultat[mat_id] = {"loturi": loturi, "total": total}
    return rezultat


def genereaza_fisa_limita_xlsx(order, consum, cale_iesire):
    """Genereaza fisierul .xlsx \"Fisa limita -cda\", cu structura identica
    formularului tiparit: cate un rand per material, cu LOT-ul (loturile)
    efectiv consumate si cantitatea reala scazuta din stoc."""
    if not OPENPYXL_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'openpyxl' nu este instalata. Ruleaza: pip install openpyxl"
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fisa limita"

    subtire = Side(style="thin", color="000000")
    chenar = Border(left=subtire, right=subtire, top=subtire, bottom=subtire)

    ws.merge_cells("A1:E1")
    ws["A1"] = "FISA  LIMITA  -cda"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center")

    if order.get("nume"):
        ws.merge_cells("A2:E2")
        ws["A2"] = order["nume"]
        ws["A2"].alignment = Alignment(horizontal="center")

    rand_antet = 4
    for col in range(1, 6):
        ws.cell(row=rand_antet, column=col).border = chenar
    ws.cell(row=rand_antet, column=2, value="LOT").font = Font(italic=True, bold=True)
    ws.cell(row=rand_antet, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=rand_antet, column=3, value="Cantitate,Kg").font = Font(italic=True, bold=True)
    ws.cell(row=rand_antet, column=3).alignment = Alignment(horizontal="center")

    rand = rand_antet + 1
    for mat in MATERIALE:
        eticheta = FISA_LIMITA_ETICHETE.get(mat["id"], mat["nume"])
        date_mat = consum.get(mat["id"], {"loturi": [], "total": 0.0})
        loturi_text = ", ".join(nr for nr, _ in date_mat["loturi"])
        total = date_mat["total"]

        ws.cell(row=rand, column=1, value=eticheta)
        ws.cell(row=rand, column=2, value=loturi_text)
        ws.cell(row=rand, column=3, value=round(total, 3) if total else None)
        ws.cell(row=rand, column=3).alignment = Alignment(horizontal="right")
        for col in range(1, 6):
            ws.cell(row=rand, column=col).border = chenar
        rand += 1

    rand += 4
    ws.cell(row=rand, column=1, value="Nume/Semnatura")

    rand += 6
    ws.cell(row=rand, column=1, value="Sef Sectie Lingouri")
    ws.cell(row=rand, column=5, value="Intocmit,")

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    wb.save(cale_iesire)


def genereaza_fisa_limita_pdf(order, consum, cale_iesire):
    """Genereaza un raport .pdf cu acelasi continut ca fisa limita .xlsx,
    pentru arhivare/tiparire."""
    if not REPORTLAB_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'reportlab' nu este instalata. Ruleaza: pip install reportlab"
        )

    doc = SimpleDocTemplate(
        cale_iesire, pagesize=A4,
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )
    stiluri = getSampleStyleSheet()
    stil_titlu = ParagraphStyle("TitluFisaLimita", parent=stiluri["Heading2"], alignment=TA_CENTER)

    elemente = [Paragraph("FISA LIMITA -cda", stil_titlu)]
    if order.get("nume"):
        elemente.append(Paragraph(order["nume"], ParagraphStyle(
            "SubtitluFisaLimita", parent=stiluri["Normal"], alignment=TA_CENTER
        )))
    elemente.append(Spacer(1, 0.6 * cm))

    date_tabel = [["", "LOT", "Cantitate,Kg", "", ""]]
    for mat in MATERIALE:
        eticheta = FISA_LIMITA_ETICHETE.get(mat["id"], mat["nume"])
        date_mat = consum.get(mat["id"], {"loturi": [], "total": 0.0})
        loturi_text = ", ".join(nr for nr, _ in date_mat["loturi"])
        total = date_mat["total"]
        date_tabel.append([eticheta, loturi_text, fmt(total, 3) if total else "", "", ""])

    tabel = Table(date_tabel, colWidths=[4.2 * cm, 4.4 * cm, 3.2 * cm, 2.5 * cm, 2.5 * cm])
    tabel.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
        ("FONTNAME", (1, 0), (2, 0), "Helvetica-BoldOblique"),
        ("ALIGN", (1, 0), (2, 0), "CENTER"),
        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elemente.append(tabel)

    elemente.append(Spacer(1, 2.2 * cm))
    elemente.append(Paragraph("Nume/Semnatura", stiluri["Normal"]))
    elemente.append(Spacer(1, 2.2 * cm))

    tabel_semnaturi = Table([["Sef Sectie Lingouri", "", "Intocmit,"]],
                             colWidths=[6 * cm, 6 * cm, 5 * cm])
    tabel_semnaturi.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),
    ]))
    elemente.append(tabel_semnaturi)

    doc.build(elemente)


# ---------------------------------------------------------------------------
# Generare "RetDozare" (.xlsx + .pdf cu o pagina per reteta) — reproduce
# tabul "RetDozare" folosit pana acum manual in Excel, pentru comenzile de
# tip Ti6Al4V. Calculul e specific acestui tip de aliaj; cand apar alte
# aliaje, se adauga generatoare proprii (structura de mai jos e izolata
# tocmai pentru asta).
# ---------------------------------------------------------------------------
def construieste_date_retdozare_comanda(order, lots):
    """Calculeaza datele RetDozare pentru TOATE retetele comenzii, intr-o
    singura trecere secventiala (in ordinea lor din order['retete']).

    Important: bilantul de stoc se tine per LOT (nu per reteta), pornind de
    la stocul de INTRARE al lotului (cantitatea completa primita), nu de la
    ce a mai ramas AZI din el (stocIntrare - consum). Un lot poate fi fost
    consumat intre timp si de alte comenzi/retete facute ulterior — asta nu
    schimba faptul ca, la momentul in care s-a facut ACEASTA comanda, lotul
    era intreg. RetDozare simuleaza deci planul acestei comenzi ca si cum
    ar fi executat imediat, pornind de la stocul intreg al lotului, apoi
    scazand DOAR consumul barelor din comanda curenta, in ordine.

    Daca doua retete succesive din aceeasi comanda folosesc acelasi lot
    fizic (ex. acelasi lot de Al metal / Fe metal / TiO2 refolosit pe mai
    multe retete), stocul de pornire al celei de-a doua retete este stocul
    RAMAS dupa consumul din reteta anterioara — nu se reseteaza la stocul
    de intrare de doua ori pentru acelasi lot.

    Returneaza o lista de dict-uri, unul per reteta, in ordinea lor.
    """
    stoc_per_lot = {l["id"]: to_float(l.get("stocIntrare")) for l in lots}

    def _stoc_pt_materiale(lot_sel, materiale_active):
        return {
            mat["id"]: stoc_per_lot.get(lot_sel[mat["id"]]["id"], 0.0) if lot_sel.get(mat["id"]) else 0.0
            for mat in materiale_active
        }

    rezultate = []
    for r in order.get("retete", []):
        materiale_active = [
            m for m in MATERIALE
            if m["id"] == "burete" or material_necesar(m["id"], r["target"])
        ]
        lot_sel = {k: next((l for l in lots if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}

        stoc_initial = _stoc_pt_materiale(lot_sel, materiale_active)

        bilant_randuri = []  # [(eticheta, {material_id: stoc_kg}), ...]
        offset_bare = _numar_bare_anterioare(order, r["id"])
        for i, bar in enumerate(r.get("bare", [])):
            # RetDozare e o fisa de PLANIFICARE (ca tab-ul original din Excel):
            # foloseste mereu dozarea calculata cu parametrii ACTUALI ai
            # retetei (target/loturi/portie/nr.presari), nu un calcSnapshot
            # inghetat de cand o bara a fost aplicata pe stoc cu alti
            # parametri — altfel barele din aceeasi reteta ar iesi cu
            # cantitati diferite intre ele, desi ar trebui sa fie identice.
            calc = calculeaza_bara(r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
            bilant_randuri.append((f"B{offset_bare + i + 1}", _stoc_pt_materiale(lot_sel, materiale_active)))
            if not calc.get("eroare"):
                for mat in materiale_active:
                    lot = lot_sel.get(mat["id"])
                    if lot:
                        stoc_per_lot[lot["id"]] -= calc["rezultatTotal"].get(mat["id"], 0.0)
                bilant_randuri.append(("", _stoc_pt_materiale(lot_sel, materiale_active)))

        portie = to_float(r.get("portie"))
        nr_presari = to_float(r.get("numarPresari"), 1) or 1
        randuri_portie = []
        for p_ref in (portie, portie / 2 if portie else 0.0, portie * nr_presari if portie else 0.0):
            if p_ref > 0:
                calc_ref = calculeaza_bara(r["target"], lot_sel, p_ref, 1)
            else:
                calc_ref = {"eroare": "-"}
            valori = None if calc_ref.get("eroare") else {
                mat["id"]: calc_ref["rezultat"].get(mat["id"], 0.0) for mat in materiale_active
            }
            randuri_portie.append((p_ref, valori))

        rezultate.append({
            "reteta": r,
            "materiale_active": materiale_active,
            "lot_sel": lot_sel,
            "stoc_initial": stoc_initial,
            "bilant_randuri": bilant_randuri,
            "randuri_portie": randuri_portie,
            "nr_presari": nr_presari,
            "portie": portie,
        })

    return rezultate


def genereaza_retdozare_xlsx(order, lots, cale_iesire):
    """Genereaza fisierul .xlsx \"RetDozare\", cu cate un bloc per reteta,
    in formatul folosit pana acum manual in Excel."""
    if not OPENPYXL_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'openpyxl' nu este instalata. Ruleaza: pip install openpyxl"
        )

    spec = ALIAJE_SPEC.get(order.get("tipAliaj"), {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RetDozare"

    bold = Font(bold=True)
    italic_bold = Font(bold=True, italic=True)

    rand = 1
    for date_ret in construieste_date_retdozare_comanda(order, lots):
        r = date_ret["reteta"]
        rand = _scrie_bloc_retdozare_xlsx(ws, rand, order, r, date_ret, spec, bold, italic_bold)
        rand += 2

    for col, latime in zip("ABCDEFGHIJ", [20, 14, 14, 12, 12, 12, 12, 12, 12, 12]):
        ws.column_dimensions[col].width = latime

    wb.save(cale_iesire)


def _scrie_bloc_retdozare_xlsx(ws, start_row, order, r, date_ret, spec, bold, italic_bold):
    """Scrie blocul unei singure retete incepand de la randul start_row si
    returneaza randul imediat urmator liber."""
    rand = start_row
    ws.cell(row=rand, column=7, value="Comanda").font = bold
    ws.cell(row=rand, column=8, value=order.get("nume", ""))
    rand += 1

    etichete_spec = ["O max", "Fe max", "N max", "C max", "H max", "Al min", "Al max", "V min", "V max"]
    valori_spec = [spec.get("o_max"), spec.get("fe_max"), spec.get("n_max"), spec.get("c_max"),
                   spec.get("h_max"), spec.get("al_min"), spec.get("al_max"),
                   spec.get("v_min"), spec.get("v_max")]
    for c, titlu in enumerate(etichete_spec, start=2):
        ws.cell(row=rand, column=c, value=titlu).font = italic_bold
    rand += 1
    for c, val in enumerate(valori_spec, start=2):
        ws.cell(row=rand, column=c, value=val)
    rand += 2

    ws.cell(row=rand, column=1, value="Dozare").font = bold
    ws.cell(row=rand, column=2, value=round(to_float(r["target"].get("o")) / 100, 6))
    ws.cell(row=rand, column=7, value=round(to_float(r["target"].get("al")) / 100, 6))
    ws.cell(row=rand, column=9, value=round(to_float(r["target"].get("v")) / 100, 6))
    rand += 2

    ws.cell(row=rand, column=1, value=f"Loturi si compozitie initiala \u2014 {r.get('nume', '')}").font = bold
    rand += 1
    coloane_lot = ["Material", "LOT", "Stoc disponibil [Kg]", "% Ti", "% Al", "% V", "% O", "% Fe"]
    for c, titlu in enumerate(coloane_lot, start=1):
        ws.cell(row=rand, column=c, value=titlu).font = italic_bold
    rand += 1
    for mat in date_ret["materiale_active"]:
        lot = date_ret["lot_sel"].get(mat["id"])
        ws.cell(row=rand, column=1, value=mat["nume"])
        ws.cell(row=rand, column=2, value=lot.get("nrLot", "") if lot else "")
        ws.cell(row=rand, column=3, value=round(date_ret["stoc_initial"].get(mat["id"], 0.0), 3))
        if lot:
            ti_val = lot_ti(lot)
            ws.cell(row=rand, column=4, value=round(ti_val, 4) if ti_val else None)
            ws.cell(row=rand, column=5, value=round(to_float(lot.get("dozaAl")), 4) or None)
            ws.cell(row=rand, column=6, value=round(to_float(lot.get("dozaV")), 4) or None)
            ws.cell(row=rand, column=7, value=round(to_float(lot.get("dozaO")), 4) or None)
            ws.cell(row=rand, column=8, value=round(to_float(lot.get("dozaFe")), 4) or None)
        rand += 1
    rand += 1

    ws.cell(row=rand, column=1, value=f"Bilant presare [Kg] \u2014 {r.get('nume', '')}").font = bold
    rand += 1
    for c, mat in enumerate(date_ret["materiale_active"], start=2):
        ws.cell(row=rand, column=c, value=mat["nume"]).font = italic_bold
    rand += 1
    if not date_ret["bilant_randuri"]:
        ws.cell(row=rand, column=1, value="(nicio bara adaugata la aceasta reteta)")
        rand += 1
    else:
        for eticheta, valori in date_ret["bilant_randuri"]:
            ws.cell(row=rand, column=1, value=eticheta)
            for c, mat in enumerate(date_ret["materiale_active"], start=2):
                ws.cell(row=rand, column=c, value=round(valori.get(mat["id"], 0.0), 3))
            rand += 1
    rand += 1

    ws.cell(row=rand, column=1, value="Concentratii tinta in bara [%]").font = bold
    rand += 1
    ws.cell(row=rand, column=1, value="Ti (bal.)").font = italic_bold
    ws.cell(row=rand, column=2, value="Al").font = italic_bold
    ws.cell(row=rand, column=3, value="V").font = italic_bold
    ws.cell(row=rand, column=4, value="O").font = italic_bold
    ws.cell(row=rand, column=5, value="Fe").font = italic_bold
    rand += 1
    ws.cell(row=rand, column=1, value=round(target_ti(r["target"]), 4))
    ws.cell(row=rand, column=2, value=round(to_float(r["target"].get("al")), 4))
    ws.cell(row=rand, column=3, value=round(to_float(r["target"].get("v")), 4))
    ws.cell(row=rand, column=4, value=round(to_float(r["target"].get("o")), 4))
    ws.cell(row=rand, column=5, value=round(to_float(r["target"].get("fe")), 4))
    rand += 2

    ws.cell(row=rand, column=1, value="Dozare portie [Kg]").font = bold
    rand += 1
    ws.cell(row=rand, column=1, value="Portie").font = italic_bold
    for c, mat in enumerate(date_ret["materiale_active"], start=2):
        ws.cell(row=rand, column=c, value=mat["nume"]).font = italic_bold
    ws.cell(row=rand, column=len(date_ret["materiale_active"]) + 3, value="Nr.Pres.").font = italic_bold
    rand += 1
    for idx, (p_ref, valori) in enumerate(date_ret["randuri_portie"]):
        ws.cell(row=rand, column=1, value=round(p_ref, 3) if p_ref else None)
        if valori:
            for c, mat in enumerate(date_ret["materiale_active"], start=2):
                ws.cell(row=rand, column=c, value=round(valori.get(mat["id"], 0.0), 6))
        if idx == len(date_ret["randuri_portie"]) - 1:
            ws.cell(row=rand, column=len(date_ret["materiale_active"]) + 3, value=date_ret["nr_presari"])
        rand += 1

    return rand


def genereaza_retdozare_pdf(order, lots, cale_iesire):
    """Genereaza raportul .pdf \"RetDozare\", cu o pagina separata per
    reteta a comenzii."""
    if not REPORTLAB_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'reportlab' nu este instalata. Ruleaza: pip install reportlab"
        )

    spec = ALIAJE_SPEC.get(order.get("tipAliaj"), {})
    doc = SimpleDocTemplate(
        cale_iesire, pagesize=A4,
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )
    stiluri = getSampleStyleSheet()
    stil_titlu = ParagraphStyle("TitluRetDozare", parent=stiluri["Heading2"], alignment=TA_CENTER)
    stil_sectiune = ParagraphStyle("SectiuneRetDozare", parent=stiluri["Heading4"], spaceBefore=10)

    stil_tabel_standard = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-BoldOblique"),
    ])

    toate_datele = construieste_date_retdozare_comanda(order, lots)
    elemente = []
    for idx, date_ret in enumerate(toate_datele):
        r = date_ret["reteta"]
        materiale = date_ret["materiale_active"]
        nume_mat = [m["nume"] for m in materiale]

        elemente.append(Paragraph(f"RetDozare \u2014 Comanda {order.get('nume', '')}", stil_titlu))
        elemente.append(Paragraph(r.get("nume", ""), ParagraphStyle(
            "SubtitluRetDozare", parent=stiluri["Normal"], alignment=TA_CENTER
        )))
        elemente.append(Spacer(1, 0.4 * cm))

        elemente.append(Paragraph("Limite chimice aliaj / Dozare tinta", stil_sectiune))
        tabel_spec = Table([
            ["O max", "Fe max", "N max", "C max", "H max", "Al min", "Al max", "V min", "V max"],
            [fmt(spec.get("o_max"), 5), fmt(spec.get("fe_max"), 5), fmt(spec.get("n_max"), 5),
             fmt(spec.get("c_max"), 5), fmt(spec.get("h_max"), 6), fmt(spec.get("al_min"), 4),
             fmt(spec.get("al_max"), 4), fmt(spec.get("v_min"), 4), fmt(spec.get("v_max"), 4)],
            ["Dozare O", "", "", "", "", "Dozare Al", "", "Dozare V", ""],
            [fmt(to_float(r["target"].get("o")) / 100, 5), "", "", "", "",
             fmt(to_float(r["target"].get("al")) / 100, 5), "", fmt(to_float(r["target"].get("v")) / 100, 5), ""],
        ])
        tabel_spec.setStyle(stil_tabel_standard)
        elemente.append(tabel_spec)

        elemente.append(Paragraph("Loturi si compozitie initiala", stil_sectiune))
        date_lot = [["Material", "LOT", "Stoc disponibil [Kg]", "% Ti", "% Al", "% V", "% O", "% Fe"]]
        for mat in materiale:
            lot = date_ret["lot_sel"].get(mat["id"])
            ti_val = lot_ti(lot) if lot else 0
            date_lot.append([
                mat["nume"],
                lot.get("nrLot", "") if lot else "",
                fmt(date_ret["stoc_initial"].get(mat["id"], 0.0), 3),
                fmt(ti_val, 4) if ti_val else "",
                fmt(to_float(lot.get("dozaAl")), 4) if lot else "",
                fmt(to_float(lot.get("dozaV")), 4) if lot else "",
                fmt(to_float(lot.get("dozaO")), 4) if lot else "",
                fmt(to_float(lot.get("dozaFe")), 4) if lot else "",
            ])
        tabel_lot = Table(date_lot, repeatRows=1)
        tabel_lot.setStyle(stil_tabel_standard)
        elemente.append(tabel_lot)

        elemente.append(Paragraph("Bilant presare [Kg]", stil_sectiune))
        date_bilant = [[""] + nume_mat]
        if not date_ret["bilant_randuri"]:
            date_bilant.append(["(nicio bara adaugata la aceasta reteta)"] + [""] * len(nume_mat))
        else:
            for eticheta, valori in date_ret["bilant_randuri"]:
                date_bilant.append([eticheta] + [fmt(valori.get(m["id"], 0.0), 3) for m in materiale])
        tabel_bilant = Table(date_bilant, repeatRows=1)
        tabel_bilant.setStyle(stil_tabel_standard)
        elemente.append(tabel_bilant)

        elemente.append(Paragraph("Concentratii tinta in bara [%]", stil_sectiune))
        tabel_conc = Table([
            ["Ti (bal.)", "Al", "V", "O", "Fe"],
            [fmt(target_ti(r["target"]), 4), fmt(to_float(r["target"].get("al")), 4),
             fmt(to_float(r["target"].get("v")), 4), fmt(to_float(r["target"].get("o")), 4),
             fmt(to_float(r["target"].get("fe")), 4)],
        ])
        tabel_conc.setStyle(stil_tabel_standard)
        elemente.append(tabel_conc)

        elemente.append(Paragraph("Dozare portie [Kg]", stil_sectiune))
        date_portie = [["Portie"] + nume_mat + ["Nr.Pres."]]
        for i_p, (p_ref, valori) in enumerate(date_ret["randuri_portie"]):
            rand_p = [fmt(p_ref, 3) if p_ref else "-"]
            if valori:
                rand_p += [fmt(valori.get(m["id"], 0.0), 4) for m in materiale]
            else:
                rand_p += [""] * len(materiale)
            rand_p.append(str(date_ret["nr_presari"]) if i_p == len(date_ret["randuri_portie"]) - 1 else "")
            date_portie.append(rand_p)
        tabel_portie = Table(date_portie, repeatRows=1)
        tabel_portie.setStyle(stil_tabel_standard)
        elemente.append(tabel_portie)

        if idx < len(toate_datele) - 1:
            elemente.append(PageBreak())

    doc.build(elemente)


def salveaza_date(state):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Dialoguri
# ---------------------------------------------------------------------------
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
        etichete_doza = [("dozaO", "Doza O (%)"), ("dozaFe", "Doza Fe (%)"),
                          ("dozaN", "Doza N (%)"), ("dozaV", "Doza V (%)"),
                          ("dozaAl", "Doza Al (%)")]
        for i, (cheie, text) in enumerate(etichete_doza):
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
        if "TiO2" in self.windowTitle():
            self.eticheta_ti.setText("Ti: 60% (fix)")
        elif "Burete" in self.windowTitle():
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



# ---------------------------------------------------------------------------
# Fereastra principala
# ---------------------------------------------------------------------------
class FereastraDozareTitan(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dozare lingouri titan")
        self.resize(1180, 780)
        self.setStyleSheet("background-color: white;")

        self.state = incarca_date()
        self.is_admin = False
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
        if self.is_admin:
            eticheta = QLabel("Admin activ")
            eticheta.setStyleSheet(
                f"color: {CULOARE_SUCCES}; font-weight: 600; font-size: 11px; "
                f"padding: 4px 10px; border: 1px solid {CULOARE_SUCCES}; border-radius: 10px;"
            )
            self.zona_admin.addWidget(eticheta)
            btn = QPushButton("Iesire admin")
            btn.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn.clicked.connect(self._logout)
            self.zona_admin.addWidget(btn)
        else:
            btn = QPushButton("Autentificare admin")
            btn.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn.clicked.connect(self._deschide_login)
            self.zona_admin.addWidget(btn)

    def _deschide_login(self):
        dialog = DialogLoginAdmin(self)
        if dialog.exec() == QDialog.Accepted:
            self.is_admin = True
            self.refresh()

    def _logout(self):
        self.is_admin = False
        self.refresh()

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

        if self.is_admin:
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
            mesaj += " Foloseste \u201e+ Lot nou\u201d." if self.is_admin else " Autentifica-te ca admin pentru a adauga loturi."
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
                ti_display, fmt(l.get("stocIntrare")), fmt(l.get("consum")),
                f"{fmt(rest)} kg",
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
            if self.is_admin:
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
                 "tipAliaj": ALIAJ_IMPLICIT, "retete": []}
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
            genereaza_fisa_limita_xlsx(order, consum, cale)
        except Exception as e:
            QMessageBox.critical(self, "Eroare la generare", str(e))
            return

        cale_pdf = os.path.splitext(cale)[0] + ".pdf"
        pdf_ok, pdf_eroare = True, ""
        if REPORTLAB_DISPONIBIL:
            try:
                genereaza_fisa_limita_pdf(order, consum, cale_pdf)
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
            genereaza_retdozare_xlsx(order, self.state["lots"], cale)
        except Exception as e:
            QMessageBox.critical(self, "Eroare la generare", str(e))
            return

        cale_pdf = os.path.splitext(cale)[0] + ".pdf"
        pdf_ok, pdf_eroare = True, ""
        if REPORTLAB_DISPONIBIL:
            try:
                genereaza_retdozare_pdf(order, self.state["lots"], cale_pdf)
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
        antet_layout.addWidget(combo_aliaj)

        n = len(order["retete"])
        meta = QLabel(f"{n} reteta{'e' if n != 1 else ''} \u00b7 creata {order['data']}")
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
        rand_titlu.addWidget(camp_nume)
        btn_sterge = QPushButton("\u2715")
        btn_sterge.setStyleSheet(STIL_BUTON_PERICOL)
        btn_sterge.clicked.connect(lambda _, oid=order["id"], rid=r["id"]: self._sterge_reteta(oid, rid))
        rand_titlu.addWidget(btn_sterge)
        layout.addLayout(rand_titlu)

        layout.addWidget(self._eticheta_eyebrow("Compozitie chimica tinta in bara (%)"))
        layout.addWidget(QLabel("Modifica valorile pentru a stabili compozitia dorita. Ti% se calculeaza automat ca rest."))
        
        rand_tinta = QHBoxLayout()
        for camp, eticheta in [("al", "Al %"), ("v", "V %"), ("o", "O %"), ("fe", "Fe %")]:
            bloc = QVBoxLayout()
            bloc.addWidget(self._eticheta_mica(eticheta))
            camp_edit = QLineEdit(str(r["target"].get(camp, 0)))
            camp_edit.setStyleSheet(STIL_CAMP)
            camp_edit.editingFinished.connect(
                lambda oid=order["id"], rid=r["id"], f=camp, c=camp_edit:
                    self._seteaza_tinta(oid, rid, f, c.text())
            )
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
            combo = QComboBox()
            combo.setStyleSheet(STIL_CAMP)
            combo.setEnabled(necesar)
            combo.addItem("\u2014 alege lot \u2014" if necesar else "\u2014 neutilizat \u2014", "")
            lots_mat = [l for l in self.state["lots"] if l["material"] == mat["id"]]
            sel_actual = r["lotSel"].get(mat["id"], "")
            index_sel = 0
            for i, l in enumerate(lots_mat, start=1):
                text = f"{l['nrLot']} (rest {fmt(lot_rest(l), 0)} kg)"
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
        bloc_portie.addWidget(camp_portie)
        rand_multiplicatori.addLayout(bloc_portie)

        bloc_nr_bare = QVBoxLayout()
        bloc_nr_bare.addWidget(self._eticheta_mica("Numar de bare"))
        camp_nr_bare = QLineEdit(str(len(r["bare"])))
        camp_nr_bare.setStyleSheet(STIL_CAMP)
        camp_nr_bare.setFixedWidth(90)
        camp_nr_bare.editingFinished.connect(
            lambda oid=order["id"], rid=r["id"], c=camp_nr_bare: self._seteaza_numar_bare(oid, rid, c.text())
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
        bloc_nr_presari.addWidget(camp_nr_presari)
        rand_multiplicatori.addLayout(bloc_nr_presari)
        rand_multiplicatori.addStretch(1)
        layout.addLayout(rand_multiplicatori)

        layout.addWidget(self._eticheta_eyebrow("Bare - urmarire consum pe stoc (toate folosesc aceeasi portie si numar de presari)"))
        offset_bare = _numar_bare_anterioare(order, r["id"])
        for i, bar in enumerate(r["bare"]):
            layout.addWidget(self._construieste_rand_bara(order, r, bar, offset_bare + i))
        btn_bara = QPushButton("+ Bara noua")
        btn_bara.setStyleSheet(STIL_BUTON_SECUNDAR)
        btn_bara.clicked.connect(lambda _, oid=order["id"], rid=r["id"]: self._adauga_bara(oid, rid))
        layout.addWidget(btn_bara)

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

        lot_sel = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        totaluri = {k: 0.0 for k in ORDINE_MAT}
        erori = []
        for bar in r["bare"]:
            if bar.get("calcSnapshot"):
                calc = bar["calcSnapshot"]
            else:
                calc = calculeaza_bara(r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
            if calc.get("eroare"):
                erori.append(calc["eroare"])
                continue
            for k in ORDINE_MAT:
                totaluri[k] += calc["rezultatTotal"][k]

        if erori:
            layout.addWidget(self._banner(erori[0], "err"))

        grid = QHBoxLayout()
        depaseste_stoc = False
        for mat in MATERIALE:
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

    def _seteaza_numar_bare(self, order_id, recipe_id, text):
        r = self._gaseste_reteta(order_id, recipe_id)
        n = int(to_float(text, len(r["bare"])))
        if n < 0:
            n = 0
        curente = r["bare"]
        if n > len(curente):
            for _ in range(n - len(curente)):
                curente.append({"id": uid(), "consumApplied": False})
        elif n < len(curente):
            eliminate = curente[n:]
            if any(b.get("consumApplied") for b in eliminate):
                raspuns = QMessageBox.question(
                    self, "Confirmare",
                    "Unele bare care ar fi sterse au deja consum aplicat pe stoc. "
                    "Daca continui, acele bare sunt sterse dar consumul deja aplicat pe loturi NU este anulat automat. Continui?"
                )
                if raspuns != QMessageBox.Yes:
                    self._rebuild_retete()
                    return
            r["bare"] = curente[:n]
        salveaza_date(self.state)
        self._rebuild_retete()

    def _adauga_bara(self, order_id, recipe_id):
        r = self._gaseste_reteta(order_id, recipe_id)
        r["bare"].append({"id": uid(), "consumApplied": False})
        salveaza_date(self.state)
        self._rebuild_retete()

    def _sterge_bara(self, order_id, recipe_id, bar_id):
        r = self._gaseste_reteta(order_id, recipe_id)
        r["bare"] = [b for b in r["bare"] if b["id"] != bar_id]
        salveaza_date(self.state)
        self._rebuild_retete()

    def _construieste_rand_bara(self, order, r, bar, index):
        all_selected = all(r["lotSel"].get(k) for k in ORDINE_MAT if material_necesar(k, r["target"]))

        cadru = QFrame()
        cadru.setStyleSheet(f"QFrame {{ border: 1px solid {CULOARE_BORDURA}; border-radius: 7px; }}")
        layout = QVBoxLayout(cadru)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        rand_sus = QHBoxLayout()
        eticheta_bara = QLabel(f"Bara #{index + 1}")
        eticheta_bara.setStyleSheet("font-weight: 600;")
        rand_sus.addWidget(eticheta_bara)
        rand_sus.addStretch(1)

        if bar.get("consumApplied"):
            tag = QLabel("Consum aplicat")
            tag.setStyleSheet(TAG_STYLES["good"])
            rand_sus.addWidget(tag)
            btn_revert = QPushButton("Anuleaza consumul")
            btn_revert.setStyleSheet(STIL_BUTON_SECUNDAR)
            btn_revert.clicked.connect(
                lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._revert_consum(oid, rid, bid)
            )
            rand_sus.addWidget(btn_revert)

        btn_sterge = QPushButton("\u2715")
        btn_sterge.setStyleSheet(STIL_BUTON_PERICOL)
        btn_sterge.clicked.connect(
            lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._sterge_bara(oid, rid, bid)
        )
        rand_sus.addWidget(btn_sterge)
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
                calc = calculeaza_bara(r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
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
            if not bar.get("consumApplied"):
                btn_muta = QPushButton("\u2192 Muta bara pe reteta urmatoare")
                btn_muta.setStyleSheet(STIL_BUTON_SECUNDAR)
                btn_muta.clicked.connect(
                    lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._muta_bara_pe_urmatoarea_reteta(oid, rid, bid)
                )
                layout.addWidget(btn_muta)

        if not bar.get("consumApplied"):
            btn = QPushButton("✅ Aplica consumul pe stoc")
            btn.setStyleSheet(STIL_BUTON_PRINCIPAL)
            btn.clicked.connect(
                lambda _, oid=order["id"], rid=r["id"], bid=bar["id"]: self._aplica_consum(oid, rid, bid)
            )
            layout.addWidget(btn)
        else:
            layout.addWidget(self._banner("✅ Consumul a fost aplicat pe stoc.", "ok"))

        return container

    def _aplica_consum(self, order_id, recipe_id, bar_id):
        r = self._gaseste_reteta(order_id, recipe_id)
        bar = next(b for b in r["bare"] if b["id"] == bar_id)
        lot_sel = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        calc = calculeaza_bara(r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
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
            lot["consum"] = to_float(lot.get("consum")) + calc["rezultatTotal"][k]
            snapshot[k] = [{"lotId": lot["id"], "kg": calc["rezultatTotal"][k]}]
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

        # Dozarea (portia) barei se calculeaza o singura data, pe baza retetei
        # curente — aceasta ramane fixa indiferent unde ajunge bara.
        lot_sel_curent = {k: next((l for l in self.state["lots"] if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}
        calc = calculeaza_bara(r["target"], lot_sel_curent, r.get("portie"), r.get("numarPresari") or 1)
        if calc.get("eroare"):
            QMessageBox.warning(self, "Eroare", calc["eroare"])
            return

        lot_sel_urmator = {k: next((l for l in self.state["lots"] if l["id"] == reteta_urmatoare["lotSel"].get(k)), None) for k in ORDINE_MAT}

        # Verificam dinainte daca vreun material are nevoie de un lot "de
        # rezerva" (stocul din lotul curent nu ajunge) si daca acel lot e
        # deja selectat in reteta urmatoare — altfel oprim mutarea aici.
        lipsuri = []
        for k in ORDINE_MAT:
            necesar = calc["rezultatTotal"][k]
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
            necesar = calc["rezultatTotal"][k]
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


def main():
    app = QApplication(sys.argv)
    fereastra = FereastraDozareTitan()
    fereastra.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
