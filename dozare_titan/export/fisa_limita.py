"""Generare "Fisa limita -cda" (.xlsx + raport .pdf), pe baza consumului
REAL de materiale (ce s-a aplicat efectiv pe stoc), independent de tipul
de aliaj al comenzii."""

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
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    REPORTLAB_DISPONIBIL = True
except ImportError:
    REPORTLAB_DISPONIBIL = False

from ..config import MATERIALE, FISA_LIMITA_ETICHETE
from ..utils import to_float, fmt


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
    # importat local pentru a evita dependinta circulara (export -> persistenta)
    from ..persistenta import _numar_bare_anterioare

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
