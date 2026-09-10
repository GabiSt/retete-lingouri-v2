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

from ..config import MATERIALE, ORDINE_MAT, ALIAJE_SPEC, ALIAJ_IMPLICIT
from ..utils import to_float, fmt
from ..calcule import calculeaza_bara, target_ti, material_necesar
from ..persistenta import _numar_bare_anterioare


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
        materiale_active = [
            m for m in MATERIALE
            if m["id"] == "burete" or material_necesar(m["id"], r["target"])
        ]
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

        # Lotul "curent" pe masura ce parcurgem istoricul barelor produse —
        # se actualizeaza de fiecare data cand o bara arata ca s-a trecut,
        # intre timp, pe alt lot pentru un material.
        lot_in_uz = dict(lot_initial)

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
                bilant_randuri.append(("", _stoc_pt_materiale(lot_in_uz, materiale_active)))
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
                    bilant_randuri.append(("", _stoc_pt_materiale(lot_sel, materiale_active)))

        portie = to_float(r.get("portie"))
        nr_presari = to_float(r.get("numarPresari"), 1) or 1
        randuri_portie = []
        for p_ref in (portie, portie / 2 if portie else 0.0, portie * nr_presari if portie else 0.0):
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

    from openpyxl.styles import Font as _Font  # import local, doar aici e nevoie

    spec = ALIAJE_SPEC.get(order.get("tipAliaj"), {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RetDozare"

    bold = _Font(bold=True)
    italic_bold = _Font(bold=True, italic=True)

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
    returneaza randul imediat urmator liber.

    NOTA: coloanele de mai jos (O/Fe/N/C/H/Al/V) reflecta limitele chimice
    ale Ti6Al4V. Pentru Ti-VT9, spec va contine si "si_min"/"zr_min"/
    "mo_min" (vezi config.ALIAJE_SPEC) — daca vrei ca ele sa apara si in
    RetDozare, extinde listele etichete_spec/valori_spec de mai jos.
    """
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
        lot = date_ret["lot_initial"].get(mat["id"])
        ws.cell(row=rand, column=1, value=mat["nume"])
        ws.cell(row=rand, column=2, value=lot.get("nrLot", "") if lot else "")
        ws.cell(row=rand, column=3, value=round(date_ret["stoc_initial"].get(mat["id"], 0.0), 3))
        if lot:
            from ..calcule import lot_ti as _lot_ti
            ti_val = _lot_ti(lot)
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

    from ..calcule import lot_ti as _lot_ti

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
            lot = date_ret["lot_initial"].get(mat["id"])
            ti_val = _lot_ti(lot) if lot else 0
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
