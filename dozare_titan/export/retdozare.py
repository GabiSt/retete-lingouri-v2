"""Generare "RetDozare" (.xlsx + .pdf cu o pagina per reteta) — reproduce
tabul "RetDozare" folosit pana acum manual in Excel.

Foloseste calcule.calculeaza_bara(), care alege automat formula corecta
in functie de order["tipAliaj"] — deci acest fisier NU trebuie modificat
cand adaugi un aliaj nou, doar coloanele afisate (Al/V/O/Fe) sunt inca
specifice compozitiei Ti6Al4V. Daca Ti-VT9 are elemente diferite in
compozitia rezultata (Si/Zr/Mo), tabelele de mai jos (coloanele "% Al",
"% V" etc. si sectiunea "Concentratii tinta") vor trebui adaptate.
"""

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYXL_DISPONIBIL = True
except ImportError:
    OPENPYXL_DISPONIBIL = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    REPORTLAB_DISPONIBIL = True
except ImportError:
    REPORTLAB_DISPONIBIL = False

from ..config import (
    MATERIALE, ORDINE_MAT, ALIAJE_SPEC, ALIAJ_IMPLICIT, ORDINE_AFISARE_MATERIALE,
    MATERIAL_NUME_RETDOZARE, MATERIAL_ABREVIERE_BILANT, MATERIAL_ELEMENTE_RELEVANTE,
    SEF_SECTIE_LINGOURI, INTOCMIT_NUME, COD_FORMULAR_RETDOZARE,
)
from ..utils import to_float, fmt, numar_comanda, pct
from ..calcule import calculeaza_bara, target_ti, material_necesar
from ..calcule.comun import ELEMENT_PER_MATERIAL, lot_ti
from ..persistenta import _numar_bare_anterioare
from .logo import adauga_logo_xlsx, logo_flowable_pdf

_MATERIALE_PRIN_ID = {m["id"]: m for m in MATERIALE}


def _element_principal(mat_id):
    """Elementul chimic "principal" al unui material, folosit ca a patra
    coloana (procent) in tabelul compact "Loturi si compozitie initiala"
    (ex. burete -> Ti, Aliaj Al-V -> V, Al metal -> Al, TiO2 -> O)."""
    cheie = ELEMENT_PER_MATERIAL.get(mat_id)
    return cheie.capitalize() if cheie else "Ti"


def _materiale_active_ordonate(target):
    """Materialele active ale unei retete, in ordinea de afisare a
    formularelor tiparite (ORDINE_AFISARE_MATERIALE, nu ordinea interna
    ORDINE_MAT)."""
    return [
        _MATERIALE_PRIN_ID[mid] for mid in ORDINE_AFISARE_MATERIALE
        if mid in _MATERIALE_PRIN_ID and (mid == "burete" or material_necesar(mid, target))
    ]


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

    Bare deja PRODUSE (bar["consumApplied"] cu calcSnapshot/consumSnapshot)
    urmaresc STRICT istoricul real inregistrat la momentul aplicarii —
    cantitatile (calcSnapshot) si loturile efectiv consumate
    (consumSnapshot, care poate contine mai multe loturi pentru acelasi
    material daca stocul s-a schimbat in timpul acelei bare: intai ce mai
    ramanea din lotul vechi, apoi restul din lotul nou). NU se mai
    recalculeaza aceste bare cu selectia curenta de lot a retetei — altfel,
    daca lotul a fost schimbat intre timp (cel vechi ramas la 0 stoc),
    simularea ar scadea din GRESEALA toata cantitatea din lotul nou, ca si
    cum bara ar fi fost produsa integral din el, ignorand ce s-a consumat
    real din lotul vechi.

    Doar barele INCA NEPRODUSE (fara consumApplied) raman un PLAN: acestea
    se calculeaza cu parametrii ACTUALI ai retetei (target/loturi/portie/
    nr.presari), pentru ca reflecta ce se va intampla daca sunt produse
    acum, cu loturile curent selectate.

    Returneaza o lista de dict-uri, unul per reteta, in ordinea lor.
    """
    tip_aliaj = order.get("tipAliaj", ALIAJ_IMPLICIT)
    stoc_per_lot = {l["id"]: to_float(l.get("stocIntrare")) for l in lots}
    lots_by_id = {l["id"]: l for l in lots}

    def _stoc_pt_materiale(lot_map, materiale_active):
        return {
            mat["id"]: stoc_per_lot.get(lot_map[mat["id"]]["id"], 0.0) if lot_map.get(mat["id"]) else 0.0
            for mat in materiale_active
        }

    rezultate = []
    for r in order.get("retete", []):
        materiale_active = _materiale_active_ordonate(r["target"])
        lot_sel = {k: next((l for l in lots if l["id"] == r["lotSel"].get(k)), None) for k in ORDINE_MAT}

        # Lotul "in uz" pentru afisare (Loturi si compozitie initiala +
        # Bilant presare al barelor deja produse): porneste de la selectia
        # curenta, dar e corectat mai jos pe baza istoricului REAL al
        # primei bare produse, daca acela arata alt lot decat cel selectat
        # acum (de ex. lotul initial a fost intre timp epuizat/inlocuit).
        lot_initial = dict(lot_sel)
        for bar in r.get("bare", []):
            if not (bar.get("consumApplied") and bar.get("consumSnapshot")):
                continue
            for k in ORDINE_MAT:
                intrari = bar["consumSnapshot"].get(k) or []
                if intrari:
                    lot_prim = lots_by_id.get(intrari[0].get("lotId"))
                    if lot_prim:
                        lot_initial[k] = lot_prim
            break  # doar prima bara deja produsa conteaza pentru starea initiala

        stoc_initial = _stoc_pt_materiale(lot_initial, materiale_active)

        # Compozitia (%) fiecarui lot folosit initial, pentru tabelul
        # "Initial" (transpus: randuri = elemente, coloane = materiale).
        comps_initial = {}
        for mat in materiale_active:
            lot = lot_initial.get(mat["id"])
            if not lot:
                comps_initial[mat["id"]] = {}
                continue
            comps_initial[mat["id"]] = {
                "Ti": lot_ti(lot),
                "Al": to_float(lot.get("dozaAl")),
                "V": to_float(lot.get("dozaV")),
                "O": to_float(lot.get("dozaO")),
                "Fe": to_float(lot.get("dozaFe")),
                "Mo": to_float(lot.get("dozaMo")),
                "Si": to_float(lot.get("dozaSi")),
                "Zr": to_float(lot.get("dozaZr")),
            }

        # Lotul "curent" pe masura ce parcurgem istoricul barelor produse —
        # se actualizeaza de fiecare data cand o bara arata ca s-a trecut,
        # intre timp, pe alt lot pentru un material.
        lot_in_uz = dict(lot_initial)

        # NOTA: sablonul tiparit "RetDozare" (vezi Cda 26851-Ti5-2VAR-...)
        # arata O SINGURA linie per bara in "Bilant presare [Kg]" — stocul
        # DISPONIBIL INAINTE de a produce acea bara — plus o singura linie
        # finala (fara eticheta) cu stocul RAMAS dupa ULTIMA bara din
        # reteta. NU arata o pereche inainte/dupa pentru fiecare bara in
        # parte (liniile "dupa" intermediare ar fi identice cu linia
        # "inainte" a barei urmatoare, deci redundante pe formular).
        bilant_randuri = []  # [(eticheta, {material_id: stoc_kg}), ...]
        offset_bare = _numar_bare_anterioare(order, r["id"])
        for i, bar in enumerate(r.get("bare", [])):
            produsa = bool(bar.get("consumApplied") and bar.get("consumSnapshot") and bar.get("calcSnapshot"))
            if produsa:
                # Bara deja produsa: urmarim EXACT ce s-a consumat, din
                # consumSnapshot (posibil impartit intre lotul vechi si cel
                # nou) — nu recalculam cu lotul selectat ACUM in reteta.
                snapshot = bar["consumSnapshot"]
                bilant_randuri.append((f"B{offset_bare + i + 1}", _stoc_pt_materiale(lot_in_uz, materiale_active)))
                for mat in materiale_active:
                    intrari = snapshot.get(mat["id"]) or []
                    for intrare in intrari:
                        lot_id = intrare.get("lotId")
                        stoc_per_lot[lot_id] = stoc_per_lot.get(lot_id, 0.0) - to_float(intrare.get("kg"))
                    if intrari:
                        lot_final = lots_by_id.get(intrari[-1].get("lotId"))
                        if lot_final:
                            lot_in_uz[mat["id"]] = lot_final
            else:
                # Bara inca neprodusa: ramane un PLAN, calculat cu
                # parametrii ACTUALI ai retetei (target/loturi/portie/
                # nr.presari) — la fel ca pana acum.
                calc = calculeaza_bara(tip_aliaj, r["target"], lot_sel, r.get("portie"), r.get("numarPresari") or 1)
                bilant_randuri.append((f"B{offset_bare + i + 1}", _stoc_pt_materiale(lot_sel, materiale_active)))
                if not calc.get("eroare"):
                    for mat in materiale_active:
                        lot = lot_sel.get(mat["id"])
                        if lot:
                            stoc_per_lot[lot["id"]] -= calc["rezultatTotal"].get(mat["id"], 0.0)

        if r.get("bare"):
            ultima_produsa = bool(
                r["bare"][-1].get("consumApplied") and r["bare"][-1].get("consumSnapshot")
                and r["bare"][-1].get("calcSnapshot")
            )
            lot_dupa_ultima = lot_in_uz if ultima_produsa else lot_sel
            bilant_randuri.append(("", _stoc_pt_materiale(lot_dupa_ultima, materiale_active)))

        # "Dozare portie [Kg]": doar doua linii de referinta, ca in sablon —
        # o singura portie, si portia INMULTITA cu numarul de presari
        # (consumul total planificat pentru reteta).
        portie = to_float(r.get("portie"))
        nr_presari = to_float(r.get("numarPresari"), 1) or 1
        randuri_portie = []
        for p_ref in (portie, portie * nr_presari if portie else 0.0):
            if p_ref > 0:
                calc_ref = calculeaza_bara(tip_aliaj, r["target"], lot_sel, p_ref, 1)
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
            "lot_initial": lot_initial,
            "stoc_initial": stoc_initial,
            "comps_initial": comps_initial,
            "bilant_randuri": bilant_randuri,
            "randuri_portie": randuri_portie,
            "nr_presari": nr_presari,
            "portie": portie,
        })

    return rezultate


def genereaza_retdozare_xlsx(order, lots, cale_iesire, intocmit_nume=None):
    """Genereaza fisierul .xlsx \"RetDozare\", cu cate un bloc per reteta,
    reproducand sablonul tiparit folosit pana acum manual in Excel (vezi
    "Cda ...-Ti5-...pdf\": antet cu limitele chimice + \"Comanda\", tabelul
    compact de loturi, tabelul \"Initial\" (compozitie pe material),
    \"Bilant presare [Kg]\", \"Concentratii tinta in bara\" si \"Dozare
    portie [Kg]\").

    intocmit_nume: numele utilizatorului logat, afisat la rubrica
    "Intocmit" — daca lipseste (None), se foloseste config.INTOCMIT_NUME."""
    if not OPENPYXL_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'openpyxl' nu este instalata. Ruleaza: pip install openpyxl"
        )
    intocmit_nume = intocmit_nume or INTOCMIT_NUME

    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    spec = ALIAJE_SPEC.get(order.get("tipAliaj"), {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RetDozare"

    stiluri = {
        "bold": Font(bold=True),
        "italic": Font(italic=True),
        "italic_bold": Font(bold=True, italic=True),
        "titlu_logo": Font(bold=True, size=14, color="1F3864"),
        "titlu_comanda": Font(bold=True, italic=True, size=11),
        "portie_mare": Font(bold=True, size=18),
        "centru": Alignment(horizontal="center", vertical="center"),
        "dreapta": Alignment(horizontal="right", vertical="center"),
        "galben": PatternFill("solid", fgColor="FFFF99"),
        "subtire": Border(*(Side(style="thin"),) * 4),
    }

    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.pagebreak import Break

    rand = 1
    for i, date_ret in enumerate(construieste_date_retdozare_comanda(order, lots)):
        r = date_ret["reteta"]
        rand_start = rand
        rand = _scrie_bloc_retdozare_xlsx(ws, rand, order, r, date_ret, spec, stiluri, intocmit_nume)
        if i > 0:
            ws.row_breaks.append(Break(id=rand_start - 1))
        rand += 2

    latimi = {"A": 15, "B": 15, "C": 18, "D": 13, "E": 13, "F": 12, "G": 10, "H": 10}
    for col in "IJKLMNOPQRSTUVWXYZ":
        latimi[col] = 9
    for col, latime in latimi.items():
        ws.column_dimensions[col].width = latime

    # Fiecare bloc de reteta e gandit ca o pagina tiparita separata (ca in
    # sablonul original, un fisier .pdf per reteta) — orientare vedere,
    # incadrata automat pe latimea unei pagini.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    wb.save(cale_iesire)


def _scrie_bloc_retdozare_xlsx(ws, start_row, order, r, date_ret, spec, st, intocmit_nume):
    """Scrie blocul unei singure retete incepand de la randul start_row si
    returneaza randul imediat urmator liber. Reproduce, coloana cu
    coloana si rand cu rand, pozitionarile din sablonul original (foaia
    "Retete" din Excel-urile de dozare folosite pana acum manual):
    "Bilant presare [Kg]" sta ALATURI de tabelul "Loturi" (nu dedesubt),
    incepand chiar de pe primul rand al acestuia — eticheta randului
    (numele retetei / "B1", "B2"...) e in coloana H, iar valorile pe
    material incep din coloana I."""
    materiale = date_ret["materiale_active"]
    label_col_bilant = 8   # coloana H — eticheta randului ("Reteta 1", "B1", "B2"...)
    nr_col_bilant = 9      # coloana I — inceputul valorilor tabelului "Bilant presare"

    def mg(r1, c1, r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    rand = start_row

    # --- Antet pagina: logo / numele comenzii / data -------------------
    if not adauga_logo_xlsx(ws, f"A{rand}", latime_px=110):
        ws.cell(row=rand, column=1, value="ZIROM TITANIUM").font = st["titlu_logo"]
    # Randuri mai inalte decat implicit (15pt), ca logo-ul (~32pt inaltime
    # la latime_px=110) sa aiba loc fara sa se suprapuna peste blocul
    # urmator ("AMS 4928 X" / "Comanda") — vezi si nota similara din
    # fisa_limita.py.
    ws.row_dimensions[rand].height = 20
    ws.row_dimensions[rand + 1].height = 20
    mg(rand, 4, rand, 9)
    c = ws.cell(row=rand, column=4, value=order.get("nume", ""))
    c.font = st["titlu_comanda"]
    c.alignment = st["centru"]
    c = ws.cell(row=rand, column=nr_col_bilant + len(materiale), value=f"Data: {order.get('data', '')}")
    c.alignment = st["dreapta"]
    rand += 2

    # --- Bloc "AMS 4928 X" / "Comanda" / limite chimice / Dozare ------
    mg(rand, 1, rand, 7)
    c = ws.cell(row=rand, column=1, value=spec.get("titlu_formular", spec.get("nume", "")))
    c.font = st["italic_bold"]
    c.alignment = st["centru"]
    c.fill = st["galben"]
    c = ws.cell(row=rand, column=8, value="Comanda")
    c.font = st["bold"]
    c = ws.cell(row=rand, column=9, value=numar_comanda(order))
    c.font = st["bold"]
    c.alignment = st["centru"]
    c.fill = st["galben"]
    c.border = st["subtire"]
    rand_ams = rand
    rand += 1

    etichete_spec = ["O min", "O max", "Fe max", "N max", "C max", "H max", "Al min", "Al max", "V min", "V max"]
    for c_idx, titlu in enumerate(etichete_spec, start=2):
        ws.cell(row=rand, column=c_idx, value=titlu).font = st["italic"]
    rand += 1

    valori_spec = [
        (3, pct(spec.get("o_max"), 2)), (4, pct(spec.get("fe_max"), 2)), (5, pct(spec.get("n_max"), 2)),
        (6, pct(spec.get("c_max"), 2)), (7, pct(spec.get("h_max"), 4)), (8, pct(spec.get("al_min"), 2)),
        (9, pct(spec.get("al_max"), 2)), (10, pct(spec.get("v_min"), 2)), (11, pct(spec.get("v_max"), 2)),
    ]
    for c_idx, val in valori_spec:
        c = ws.cell(row=rand, column=c_idx, value=val or None)
        c.font = st["italic"]
    rand += 1

    c = ws.cell(row=rand, column=1, value="Dozare")
    c.font = st["italic_bold"]
    mg(rand, 2, rand, 3)
    c = ws.cell(row=rand, column=2, value=pct(to_float(r["target"].get("o")) / 100, 4) or None)
    c.font = st["bold"]
    c.alignment = st["centru"]
    mg(rand, 4, rand, 5)
    c = ws.cell(row=rand, column=4, value=pct(to_float(r["target"].get("fe")) / 100, 4) or None)
    c.font = st["bold"]
    c.alignment = st["centru"]
    c = ws.cell(row=rand, column=9, value=pct(to_float(r["target"].get("al")) / 100, 3) or None)
    c.font = st["bold"]
    c.alignment = st["centru"]
    mg(rand, 10, rand, 11)
    c = ws.cell(row=rand, column=10, value=pct(to_float(r["target"].get("v")) / 100, 3) or None)
    c.font = st["bold"]
    c.alignment = st["centru"]
    rand_dozare = rand
    rand += 2

    # --- "Bilant presare [Kg]" (dreapta, incepand din acelasi rand ca
    # primul rand al tabelului "Loturi" de mai jos) ---------------------
    rand += 1
    rand_loturi_start = rand
    rand_bilant = rand_loturi_start
    if materiale:
        mg(rand_bilant, nr_col_bilant, rand_bilant, nr_col_bilant + len(materiale) - 1)
    c = ws.cell(row=rand_bilant, column=nr_col_bilant, value="Bilant presare [Kg]")
    c.font = st["bold"]
    c.alignment = st["centru"]
    rand_bilant += 1
    c = ws.cell(row=rand_bilant, column=label_col_bilant, value=r.get("nume") or "Reteta")
    c.font = st["bold"]
    for c_idx, mat in enumerate(materiale, start=nr_col_bilant):
        cc = ws.cell(row=rand_bilant, column=c_idx, value=MATERIAL_ABREVIERE_BILANT.get(mat["id"], mat["nume"]))
        cc.font = st["italic_bold"]
    rand_bilant += 1
    if not date_ret["bilant_randuri"]:
        ws.cell(row=rand_bilant, column=label_col_bilant, value="(nicio bara adaugata)")
        rand_bilant += 1
    else:
        for eticheta, valori in date_ret["bilant_randuri"]:
            ws.cell(row=rand_bilant, column=label_col_bilant, value=eticheta)
            for c_idx, mat in enumerate(materiale, start=nr_col_bilant):
                ws.cell(row=rand_bilant, column=c_idx, value=round(valori.get(mat["id"], 0.0), 3))
            rand_bilant += 1

    # --- Tabel compact "Loturi si compozitie initiala" (coloanele A, C,
    # D, E — coloana B ramane libera, ca in sablonul original) ----------
    rand = rand_loturi_start
    for mat in materiale:
        lot = date_ret["lot_initial"].get(mat["id"])
        elem = _element_principal(mat["id"])
        val_pct = date_ret["comps_initial"].get(mat["id"], {}).get(elem)
        ws.cell(row=rand, column=1, value=MATERIAL_NUME_RETDOZARE.get(mat["id"], mat["nume"]))
        ws.cell(row=rand, column=3, value=lot.get("nrLot", "") if lot else "")
        ws.cell(row=rand, column=4, value=round(date_ret["stoc_initial"].get(mat["id"], 0.0), 3))
        ws.cell(row=rand, column=5, value=pct(val_pct / 100, 2) if val_pct else None)
        for col_idx in (1, 3, 4, 5):
            ws.cell(row=rand, column=col_idx).border = st["subtire"]
        rand += 1
    rand_loturi_end = rand
    rand += 1

    # --- Tabel "Initial" (compozitie procentuala pe material, transpus) -
    c = ws.cell(row=rand, column=1, value="Initial")
    c.font = st["italic_bold"]
    c.border = st["subtire"]
    for c_idx, mat in enumerate(materiale, start=2):
        cc = ws.cell(row=rand, column=c_idx, value=MATERIAL_NUME_RETDOZARE.get(mat["id"], mat["nume"]))
        cc.font = st["italic_bold"]
        cc.border = st["subtire"]
    rand += 1
    elemente_ordine = ["Ti", "Al", "V", "O", "Fe", "Mo", "Si", "Zr"]
    elemente_relevante = {
        e for mat in materiale for e in MATERIAL_ELEMENTE_RELEVANTE.get(mat["id"], [])
    }
    for elem in elemente_ordine:
        if elem not in elemente_relevante:
            continue
        c = ws.cell(row=rand, column=1, value=f"[%] {elem}")
        c.border = st["subtire"]
        for c_idx, mat in enumerate(materiale, start=2):
            relevante_mat = MATERIAL_ELEMENTE_RELEVANTE.get(mat["id"], [])
            cc = ws.cell(row=rand, column=c_idx)
            cc.border = st["subtire"]
            if elem in relevante_mat:
                val = date_ret["comps_initial"].get(mat["id"], {}).get(elem)
                cc.value = pct((val or 0.0) / 100, 4)
        rand += 1
    rand += 1

    # --- "Concentratii tinta in bara [%]" -------------------------------
    ordine_elem_conc = [("v", "V"), ("al", "Al"), ("o", "O"), ("fe", "Fe"), ("mo", "Mo"), ("si", "Si"), ("zr", "Zr")]
    elem_conc = [(cheie, eticheta) for cheie, eticheta in ordine_elem_conc if abs(to_float(r["target"].get(cheie))) > 1e-9]
    c = ws.cell(row=rand, column=1, value="Concentratii")
    c.font = st["italic_bold"]
    c.border = st["subtire"]
    cc = ws.cell(row=rand, column=2, value="Ti")
    cc.font = st["italic_bold"]
    cc.alignment = st["centru"]
    cc.border = st["subtire"]
    for c_idx, (_, eticheta) in enumerate(elem_conc, start=3):
        cc = ws.cell(row=rand, column=c_idx, value=eticheta)
        cc.font = st["italic_bold"]
        cc.alignment = st["centru"]
        cc.border = st["subtire"]
    rand += 1
    c = ws.cell(row=rand, column=1, value="In bara")
    c.font = st["italic_bold"]
    c.border = st["subtire"]
    cc = ws.cell(row=rand, column=2, value="bal.")
    cc.alignment = st["centru"]
    cc.border = st["subtire"]
    for c_idx, (cheie, _) in enumerate(elem_conc, start=3):
        cc = ws.cell(row=rand, column=c_idx, value=pct(to_float(r["target"].get(cheie)) / 100, 2))
        cc.font = st["bold"]
        cc.alignment = st["centru"]
        cc.border = st["subtire"]
    rand += 2

    # --- "DOZARE PORTIE [Kg]" -------------------------------------------
    ultima_col_portie = 1 + len(materiale)
    if materiale:
        mg(rand, 1, rand, ultima_col_portie)
    c = ws.cell(row=rand, column=1, value="DOZARE PORTIE [Kg]")
    c.font = st["bold"]
    c.alignment = st["centru"]
    rand += 1
    c = ws.cell(row=rand, column=1, value="PORTIE")
    c.font = st["italic"]
    for c_idx, mat in enumerate(materiale, start=2):
        cc = ws.cell(row=rand, column=c_idx, value=MATERIAL_NUME_RETDOZARE.get(mat["id"], mat["nume"]))
        cc.font = st["italic"]
    rand += 1
    for idx, (p_ref, valori) in enumerate(date_ret["randuri_portie"]):
        e_mare = idx == 0
        font_rand = st["portie_mare"] if e_mare else Font()
        c = ws.cell(row=rand, column=1, value=round(p_ref, 3) if p_ref else None)
        c.font = font_rand
        if valori:
            for c_idx, mat in enumerate(materiale, start=2):
                cc = ws.cell(row=rand, column=c_idx, value=round(valori.get(mat["id"], 0.0), 3))
                cc.font = font_rand
        if idx == len(date_ret["randuri_portie"]) - 1:
            ws.cell(row=rand, column=ultima_col_portie + 1, value="Nr.Pres.").font = st["italic"]
            c = ws.cell(row=rand, column=ultima_col_portie + 2, value=date_ret["nr_presari"])
            c.font = st["bold"]
            c.border = st["subtire"]
        rand += 1

    rand = max(rand, rand_bilant, rand_loturi_end)

    # --- Semnaturi -------------------------------------------------------
    rand += 2
    ws.cell(row=rand, column=1, value="Sef Sectie Lingouri")
    ws.cell(row=rand, column=6, value="Intocmit,")
    rand += 1
    ws.cell(row=rand, column=1, value=SEF_SECTIE_LINGOURI)
    ws.cell(row=rand, column=6, value=intocmit_nume)
    rand += 1
    ws.cell(row=rand, column=nr_col_bilant + max(len(materiale) - 2, 0), value=f"Formular Cod:{COD_FORMULAR_RETDOZARE}")

    return rand + 1


def genereaza_retdozare_pdf(order, lots, cale_iesire, intocmit_nume=None):
    """Genereaza raportul .pdf \"RetDozare\", cu o pagina separata per
    reteta a comenzii, reproducand sablonul tiparit folosit pana acum
    manual (vezi \"Cda 26804-Ti5-2VAR500-L1-L12.pdf\"): pagina PORTRET (nu
    vedere), cu antetul de limite chimice + \"Comanda\" sus, tabelul
    compact de loturi langa \"Bilant presare [Kg]\", tabelul \"Initial\" si
    \"Concentratii tinta\" mai jos, si tabelul \"DOZARE PORTIE [Kg]\" MARE
    (cifre in font mare, bold) in partea de jos — asta e sectiunea cea
    mai importanta de pe formular, deci trebuie sa sara imediat in ochi.

    intocmit_nume: numele utilizatorului logat, afisat la rubrica
    "Intocmit" — daca lipseste (None), se foloseste config.INTOCMIT_NUME."""
    if not REPORTLAB_DISPONIBIL:
        raise RuntimeError(
            "Biblioteca 'reportlab' nu este instalata. Ruleaza: pip install reportlab"
        )
    intocmit_nume = intocmit_nume or INTOCMIT_NUME

    spec = ALIAJE_SPEC.get(order.get("tipAliaj"), {})
    doc = SimpleDocTemplate(
        cale_iesire, pagesize=A4,  # PORTRET, ca sablonul original tiparit
        leftMargin=1.3 * cm, rightMargin=1.3 * cm,
        topMargin=0.6 * cm, bottomMargin=0.5 * cm,
    )
    stiluri = getSampleStyleSheet()
    stil_titlu = ParagraphStyle("TitluRetDozare", parent=stiluri["Normal"], alignment=TA_CENTER, fontName="Helvetica-BoldOblique", fontSize=9)
    stil_data = ParagraphStyle("DataRetDozare", parent=stiluri["Normal"], alignment=2, fontSize=9)  # TA_RIGHT
    stil_sectiune = ParagraphStyle("SectiuneRetDozare", parent=stiluri["Heading4"], spaceBefore=4, spaceAfter=2, fontSize=9)

    stil_tabel_standard = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-BoldOblique"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ])
    stil_tabel_fara_antet = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ])

    USABIL_CM = 18.4  # latimea utila a paginii A4 portret, cu marginile de mai sus

    def _latimi_coloane(n, latime_eticheta, buget_total, min_col=1.0, max_col=2.6):
        """Latimea coloanelor de materiale a unui tabel, impartind spatiul
        ramas (dupa coloana de eticheta) in mod egal intre ele — cu limite
        rezonabile, ca sa functioneze corect si cu putine materiale (Ti5,
        4 coloane) si cu multe (VT9, pana la 7), fara sa iasa de pe
        pagina."""
        if n <= 0:
            return []
        disponibil = max(buget_total - latime_eticheta, n * min_col)
        lat = disponibil / n
        lat = max(min_col, min(max_col, lat))
        return [lat] * n

    def _text_o(spec):
        """Coloana "O" a tabelului de limite chimice: un singur interval
        ("0.17-0.20%") daca specificatia are si minim, altfel doar maximul
        ("0.20%") — ca pe sablonul original (o singura coloana, nu doua)."""
        o_max_txt = pct(spec.get("o_max"), 2)
        o_min_txt = pct(spec.get("o_min"), 2)
        if o_min_txt and o_max_txt:
            return f"{o_min_txt}-{o_max_txt}"
        return o_max_txt

    toate_datele = construieste_date_retdozare_comanda(order, lots)
    elemente = []
    for idx, date_ret in enumerate(toate_datele):
        r = date_ret["reteta"]
        materiale = date_ret["materiale_active"]
        nume_mat_lung = [MATERIAL_NUME_RETDOZARE.get(m["id"], m["nume"]) for m in materiale]
        nume_mat_scurt = [MATERIAL_ABREVIERE_BILANT.get(m["id"], m["nume"]) for m in materiale]

        # --- Antet pagina: logo / numele comenzii / data -------------------
        logo_cell = logo_flowable_pdf(latime_cm=3.4) or Paragraph("ZIROM TITANIUM", stiluri["Heading2"])
        antet = Table([[
            logo_cell,
            Paragraph(order.get("nume", ""), stil_titlu),
            Paragraph(f"Data: {order.get('data', '')}", stil_data),
        ]], colWidths=[4.5 * cm, 9.5 * cm, 4.5 * cm])
        antet.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        elemente.append(antet)
        elemente.append(Spacer(1, 0.1 * cm))

        # --- Bloc "AMS 4928 X" / "Comanda" / limite chimice / Dozare -----
        tabel_spec = Table([
            [spec.get("titlu_formular", spec.get("nume", "")), "", "", "", "", "", "Comanda", numar_comanda(order), ""],
            ["", "O", "Fe max", "N max", "C max", "H max", "Al min", "Al max", "V min", "V max"],
            ["", _text_o(spec), pct(spec.get("fe_max"), 2), pct(spec.get("n_max"), 2),
             pct(spec.get("c_max"), 2), pct(spec.get("h_max"), 4), pct(spec.get("al_min"), 2),
             pct(spec.get("al_max"), 2), pct(spec.get("v_min"), 2), pct(spec.get("v_max"), 2)],
            ["Dozare", pct(to_float(r["target"].get("o")) / 100, 4), "", "", "", "", "",
             pct(to_float(r["target"].get("al")) / 100, 3), pct(to_float(r["target"].get("v")) / 100, 3), ""],
        ], colWidths=[1.7 * cm] + [1.75 * cm] * 9)
        tabel_spec.setStyle(TableStyle([
            ("GRID", (0, 1), (-1, -1), 0.6, colors.black),
            ("SPAN", (0, 0), (5, 0)),
            ("BACKGROUND", (0, 0), (5, 0), colors.HexColor("#FFFF99")),
            ("BACKGROUND", (7, 0), (7, 0), colors.HexColor("#FFFF99")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-BoldOblique"),
            ("FONTNAME", (0, 1), (-1, 2), "Helvetica-Oblique"),
            ("FONTNAME", (0, 3), (0, 3), "Helvetica-BoldOblique"),
            ("FONTNAME", (1, 3), (-1, 3), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        elemente.append(tabel_spec)
        elemente.append(Spacer(1, 0.2 * cm))

        # --- "Loturi" (fara antet de coloane, ca pe sablon) + "Bilant
        # presare [Kg]" ALATURI, pe acelasi rand -----------------------
        date_lot = []
        for mat in materiale:
            lot = date_ret["lot_initial"].get(mat["id"])
            elem = _element_principal(mat["id"])
            val_pct = date_ret["comps_initial"].get(mat["id"], {}).get(elem)
            date_lot.append([
                MATERIAL_NUME_RETDOZARE.get(mat["id"], mat["nume"]),
                lot.get("nrLot", "") if lot else "",
                fmt(date_ret["stoc_initial"].get(mat["id"], 0.0), 3),
                pct(val_pct / 100, 2) if val_pct else "",
            ])
        tabel_lot = Table(date_lot, colWidths=[2.8 * cm, 2.3 * cm, 1.9 * cm, 1.5 * cm])
        tabel_lot.setStyle(stil_tabel_fara_antet)
        LAT_LOTURI = 2.8 + 2.3 + 1.9 + 1.5  # cm

        date_bilant = [[r.get("nume") or "Reteta"] + nume_mat_scurt]
        if not date_ret["bilant_randuri"]:
            date_bilant.append(["(nicio bara adaugata)"] + [""] * len(materiale))
        else:
            for eticheta, valori in date_ret["bilant_randuri"]:
                date_bilant.append([eticheta] + [fmt(valori.get(m["id"], 0.0), 3) for m in materiale])
        lat_eticheta_bilant = 1.3
        # Ce spatiu mai ramane, din latimea utila a paginii, dupa tabelul
        # "Loturi" din stanga (+ spatierea dintre cele doua tabele).
        buget_bilant = USABIL_CM - LAT_LOTURI - 0.5
        lat_bilant_mat = _latimi_coloane(len(materiale), lat_eticheta_bilant, buget_bilant, min_col=0.95, max_col=1.7)
        lat_bilant = [lat_eticheta_bilant * cm] + [w * cm for w in lat_bilant_mat]
        tabel_bilant = Table(date_bilant, colWidths=lat_bilant, repeatRows=1)
        tabel_bilant.setStyle(stil_tabel_standard)
        # Font mai mic automat cand sunt multe materiale (coloane inguste),
        # ca textul sa nu se suprapuna peste liniile tabelului.
        if len(materiale) >= 6:
            tabel_bilant.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 6.5)]))

        bloc_lot_bilant = Table([[
            [Paragraph("Loturi", stil_sectiune), tabel_lot],
            [Paragraph("Bilant presare [Kg]", stil_sectiune), tabel_bilant],
        ]], colWidths=[(LAT_LOTURI + 0.5) * cm, None])
        bloc_lot_bilant.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (1, 0), (1, 0), 14),
        ]))
        elemente.append(bloc_lot_bilant)

        # --- "Initial" (compozitie pe material) -----------------------
        elemente.append(Paragraph("Initial \u2014 compozitie pe material", stil_sectiune))
        elemente_ordine = ["Ti", "Al", "V", "O", "Fe", "Mo", "Si", "Zr"]
        elemente_relevante = {e for mat in materiale for e in MATERIAL_ELEMENTE_RELEVANTE.get(mat["id"], [])}
        date_initial = [["Initial"] + nume_mat_lung]
        for elem in elemente_ordine:
            if elem not in elemente_relevante:
                continue
            rand_e = [f"[%] {elem}"]
            for mat in materiale:
                relevante_mat = MATERIAL_ELEMENTE_RELEVANTE.get(mat["id"], [])
                if elem in relevante_mat:
                    val = date_ret["comps_initial"].get(mat["id"], {}).get(elem)
                    rand_e.append(pct((val or 0.0) / 100, 4))
                else:
                    rand_e.append("")
            date_initial.append(rand_e)
        tabel_initial = Table(
            date_initial,
            colWidths=[1.9 * cm] + [w * cm for w in _latimi_coloane(len(materiale), 1.9, USABIL_CM, min_col=1.6, max_col=2.4)],
            repeatRows=1,
        )
        tabel_initial.setStyle(stil_tabel_standard)
        elemente.append(tabel_initial)

        # --- "Concentratii tinta in bara" -------------------------------
        elemente.append(Paragraph("Concentratii tinta in bara [%]", stil_sectiune))
        ordine_elem_conc = [("v", "V"), ("al", "Al"), ("o", "O"), ("fe", "Fe"), ("mo", "Mo"), ("si", "Si"), ("zr", "Zr")]
        elem_conc = [(cheie, eticheta) for cheie, eticheta in ordine_elem_conc if abs(to_float(r["target"].get(cheie))) > 1e-9]
        tabel_conc = Table([
            ["Concentratii", "Ti"] + [eticheta for _, eticheta in elem_conc],
            ["In bara", "bal."] + [pct(to_float(r["target"].get(cheie)) / 100, 2) for cheie, _ in elem_conc],
        ], colWidths=[2.6 * cm] + [1.8 * cm] * (1 + len(elem_conc)))
        tabel_conc.setStyle(stil_tabel_standard)
        elemente.append(tabel_conc)
        elemente.append(Spacer(1, 0.25 * cm))

        # --- "DOZARE PORTIE [Kg]" — sectiunea CEA MAI IMPORTANTA de pe
        # formular: cifrele portiei trebuie sa fie mari, ca sa sara in
        # ochi de la distanta la cantarul de dozare. Randul cu o singura
        # portie foloseste font mare bold; randul cu totalul pe toata
        # reteta (portie x nr.presari) ramane mic, sub el, ca pe sablon.
        # "Nr.Pres" sta intr-o cutie separata, langa tabel — nu ca o
        # coloana in plus, care ar ingusta coloanele cu cifre mari.
        elemente.append(Paragraph("DOZARE PORTIE [Kg]", ParagraphStyle(
            "TitluPortie", parent=stiluri["Heading3"], alignment=TA_CENTER, spaceBefore=0, spaceAfter=3,
        )))
        date_portie = [["PORTIE"] + nume_mat_lung]
        for p_ref, valori in date_ret["randuri_portie"]:
            rand_p = [fmt(p_ref, 3) if p_ref else "-"]
            if valori:
                rand_p += [fmt(valori.get(m["id"], 0.0), 3) for m in materiale]
            else:
                rand_p += [""] * len(materiale)
            date_portie.append(rand_p)
        lat_portie_prima = 2.3 * cm
        n_col_portie = 1 + len(materiale)
        # Latimea disponibila pt tabelul mare, lasand loc casetei
        # "Nr.Pres" din dreapta (~3.0cm cu tot cu spatiere) — impartita
        # egal intre coloane, cu limite rezonabile ca cifrele mari sa nu
        # se inghesuie (min 1.7cm) nici sa nu depaseasca latimea paginii
        # cand sunt multe materiale (max 2.8cm).
        buget_portie = USABIL_CM - 2.3 - 3.0
        lat_portie_mat_cm = _latimi_coloane(len(materiale), 0, buget_portie, min_col=1.55, max_col=2.8)
        lat_portie_mat = lat_portie_mat_cm[0] * cm if lat_portie_mat_cm else 2.3 * cm
        # Font mare, dar mai mic automat cand sunt multe materiale (coloane
        # inguste) — altfel cifrele s-ar suprapune peste liniile tabelului.
        font_mare = 22 if len(materiale) <= 5 else (18 if len(materiale) <= 6 else 15)
        leading_mare = font_mare + 4
        tabel_portie = Table(
            date_portie,
            colWidths=[lat_portie_prima] + [lat_portie_mat] * len(materiale),
        )
        tabel_portie.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            # Antet (PORTIE / nume materiale)
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Oblique"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, 0), 3),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
            # Randul cu O SINGURA portie — mare si bold, ca sa domine
            # vizual tabelul (cerinta explicita: "sa se vada mare").
            # LEADING trebuie setat explicit si el (nu doar FONTSIZE) —
            # altfel randul mosteneste un leading mic (de la stilul
            # implicit), mai mic decat fontul de 22pt, iar cifrele mari
            # raman taiate sus de marginea randului.
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, 1), font_mare),
            ("LEADING", (0, 1), (-1, 1), leading_mare),
            ("TOPPADDING", (0, 1), (-1, 1), 10),
            ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
            ("LINEABOVE", (0, 1), (-1, 1), 1.4, colors.black),
            ("LINEBELOW", (0, 1), (-1, 1), 1.4, colors.black),
            # Randul cu totalul pe reteta (portie x nr.presari) — mic,
            # informativ, sub cel mare.
            ("FONTNAME", (0, 2), (-1, 2), "Helvetica"),
            ("FONTSIZE", (0, 2), (-1, 2), 8),
            ("TOPPADDING", (0, 2), (-1, 2), 2),
            ("BOTTOMPADDING", (0, 2), (-1, 2), 2),
        ]))

        tabel_nrpres = Table(
            [["Nr.Pres."], [str(date_ret["nr_presari"])]],
            colWidths=[2.6 * cm], rowHeights=[0.55 * cm, 0.8 * cm],
        )
        tabel_nrpres.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Oblique"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, 1), 13),
        ]))

        bloc_portie = Table(
            [[tabel_portie, tabel_nrpres]],
            colWidths=[None, 2.6 * cm],
        )
        bloc_portie.setStyle(TableStyle([
            ("VALIGN", (0, 0), (0, 0), "TOP"),
            ("VALIGN", (1, 0), (1, 0), "BOTTOM"),
            ("LEFTPADDING", (1, 0), (1, 0), 12),
        ]))
        elemente.append(bloc_portie)
        elemente.append(Spacer(1, 0.5 * cm))

        # --- Semnaturi -------------------------------------------------------
        semnaturi = Table([
            ["Sef Sectie Lingouri", "", "Intocmit,"],
            [SEF_SECTIE_LINGOURI, "", intocmit_nume],
        ], colWidths=[6 * cm, 6 * cm, 6 * cm], splitByRow=0)
        semnaturi.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        elemente.append(semnaturi)
        elemente.append(Paragraph(f"Formular Cod:{COD_FORMULAR_RETDOZARE}", ParagraphStyle(
            "CodFormularRetDozare", parent=stiluri["Normal"], alignment=2, fontSize=8,
        )))

        if idx < len(toate_datele) - 1:
            elemente.append(PageBreak())

    doc.build(elemente)
