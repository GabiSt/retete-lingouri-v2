"""Planificarea barelor pe retete si aplicarea/retragerea consumului.

Aici sta TOATA logica de business a distributiei barelor — fara nicio
dependinta de interfata grafica, ca sa poata fi verificata automat pe
Excel-urile reale de productie (vezi teste/test_cda_26858.py).
``fereastra_principala.py`` doar cheama functiile de aici si afiseaza
mesajele.

REGULA, PE SCURT (asa cum se lucreaza in sectie):

  - numarul total de bare e cunoscut de la inceput, pe COMANDA (ex. 12);
  - reteta 1 primeste atatea bare cate incap in loturile ei;
  - PRIMA bara care nu mai incape NU se toarna partial: trece pe reteta
    urmatoare, dar PASTREAZA DOZAREA VECHE (calculata pe loturile
    retetei din care pleaca). Din loturile vechi mai ia doar restul
    ramas din materialul care s-a terminat (ca sa nu ramana stoc mort),
    iar diferenta — plus doza INTREAGA la toate celelalte materiale — se
    consuma din loturile retetei urmatoare;
  - TOATE celelalte bare ale retetei urmatoare se dozeaza de la zero, pe
    loturile noi;
  - mecanismul se repeta pana se plaseaza toate barele comenzii.

Exemplu real (Cda 26858-Ti5, 12 bare, 900 kg/bara):
  Reteta 1: barele 1-4 (se termina prealiajul Al-V la bara 5)
  Reteta 2: bara 5 de report (dozare de pe reteta 1) + barele 6-10
            (se termina buretele la bara 11)
  Reteta 3: bara 11 de report (dozare de pe reteta 2) + bara 12
"""

from ..config import ORDINE_MAT
from ..utils import to_float, r2, uid
from .capacitate import desfasoara_reteta, nume_material, TOLERANTA_KG
from .comun import lot_rest, material_necesar, material_lot_lipsa


# ---------------------------------------------------------------------------
# Helpere
# ---------------------------------------------------------------------------

def loturi_selectate(lots, r):
    """Loturile efectiv selectate intr-o reteta, indexate pe material."""
    return {
        k: next((l for l in lots if l["id"] == r["lotSel"].get(k)), None)
        for k in ORDINE_MAT
    }


def _lot_dupa_id(lots, lot_id):
    return next((l for l in lots if l["id"] == lot_id), None)


def reteta_continuare(order, model):
    """Reteta noua, goala, in care utilizatorul introduce loturile noi."""
    return {
        "id": uid(),
        "nume": f"Reteta {len(order['retete']) + 1}",
        "target": dict(model["target"]),
        "lotSel": {},
        "bare": [],
        "numarPresari": model.get("numarPresari") or "1",
        "portie": model.get("portie"),
    }


def doza_veche_bara(bar):
    """Dozarea (kg pe material, cu presarile incluse) inghetata pe o bara
    de report — cea calculata in reteta din care a plecat."""
    calc = bar.get("calcSnapshot") or {}
    return calc.get("rezultatTotal") or {}


# ---------------------------------------------------------------------------
# Planificarea barelor pe retete
# ---------------------------------------------------------------------------

def planifica(order, lots, calculeaza_bara, tip_aliaj, creeaza_retete=True):
    """Distribuie barele comenzii pe retete, in functie de loturi.

    Nu modifica barele si nu atinge stocul — doar calculeaza planul.
    Singurul efect lateral (daca creeaza_retete=True) e adaugarea
    retetelor de continuare de care e nevoie ca sa incapa toate barele.
    """
    total = max(0, int(to_float(order.get("nrBare"))))
    plan = []
    report = None      # bara de report care coboara in reteta urmatoare
    # Cat "promite" planul, din fiecare lot, retetelor de mai sus. Fara
    # asta, doua retete care folosesc ACELASI lot (tipic: buretele, care
    # tine mai multe retete) si-l numara fiecare intreg, iar a doua iese
    # cu bare in plus. Retetele cu consumul deja aplicat nu intra aici:
    # scaderea lor e deja in stoc.
    consum_planificat = {}
    asignate = 0
    idx = 0
    limita_retete = len(order["retete"]) + 50   # plasa de siguranta

    while idx < len(order["retete"]) and idx < limita_retete:
        r = order["retete"][idx]
        lot_sel = loturi_selectate(lots, r)

        # Barele cu consum deja aplicat sunt batute in cuie: raman unde
        # sunt si se scad din totalul comenzii.
        aplicate = [b for b in r["bare"] if b.get("consumApplied")]
        asignate += len(aplicate)
        ramase = total - asignate

        # O bara de report poate exista DEJA in reteta, materializata de o
        # distribuire anterioara (cazul obisnuit cand reteta care a cedat-o
        # are consumul aplicat si nu se mai replanifica). Atunci ea e sursa
        # de adevar, nu planul curent.
        bara_report_existenta = next(
            (b for b in r["bare"] if b.get("reportRamas") and not b.get("consumApplied")),
            None,
        )
        if report is None and bara_report_existenta is not None:
            report = {
                "ramas": dict(bara_report_existenta["reportRamas"]),
                "dinLotAnterior": dict(bara_report_existenta.get("reportDinLotAnterior") or {}),
                "calc": bara_report_existenta.get("calcSnapshot"),
                "retetaSursa": bara_report_existenta.get("retetaSursaDozare"),
            }

        intrare = {
            "reteta": r, "report": report, "bareNoi": 0, "bareAsteptare": 0,
            "reportOut": None, "desf": None, "calc": None, "aplicate": len(aplicate),
            "stare": "ok",
        }

        # Reteta cu consumul deja aplicat e INCHISA: nu mai primeste bare
        # noi (loturile ei au fost deja scazute, iar bara pe care a cedat-o
        # exista deja, ca bara de report, in reteta urmatoare).
        if aplicate and bara_report_existenta is None:
            intrare["stare"] = "consum aplicat"
            plan.append(intrare)
            report = None
            idx += 1
            if idx >= len(order["retete"]) and (total - asignate) > 0 and creeaza_retete:
                order["retete"].append(reteta_continuare(order, r))
            continue

        if ramase <= 0:
            intrare["stare"] = "fara bare ramase"
            plan.append(intrare)
            report = None
            idx += 1
            continue

        calc = calculeaza_bara(
            tip_aliaj, r["target"], lot_sel,
            r.get("portie"), r.get("numarPresari") or 1,
        )
        if calc.get("eroare"):
            # Reteta inca nu are loturile complete: barele stau aici in
            # asteptare, inclusiv bara de report venita de mai sus.
            intrare["stare"] = "loturi incomplete"
            intrare["eroare"] = calc["eroare"]
            intrare["bareAsteptare"] = max(0, ramase - (1 if report else 0))
            plan.append(intrare)
            report = None
            break

        disponibile = max(0, ramase - (1 if report else 0))
        resturi = {}
        for k in ORDINE_MAT:
            lot = lot_sel[k]
            if lot is None:
                resturi[k] = 0.0
                continue
            resturi[k] = r2(lot_rest(lot) - consum_planificat.get(lot["id"], 0.0))
        desf = desfasoara_reteta(
            calc["rezultatTotal"], resturi, disponibile,
            report["ramas"] if report else None,
        )

        for k in ORDINE_MAT:
            lot = lot_sel[k]
            if lot is None:
                continue
            consum_planificat[lot["id"]] = r2(
                consum_planificat.get(lot["id"], 0.0)
                + desf["consumTotalReteta"].get(k, 0.0)
            )

        intrare["calc"] = calc
        intrare["desf"] = desf
        intrare["bareNoi"] = desf["bareIntregi"]
        asignate += (1 if report else 0) + desf["bareIntregi"]

        if desf["baraReport"]:
            intrare["reportOut"] = {
                # Tot ce mai are de consumat bara asta din loturile
                # retetei URMATOARE: la materialul limitant doar
                # diferenta, la celelalte doza veche intreaga.
                "ramas": {k: v for k, v in desf["baraReport"]["dinLotNou"].items() if v > 1e-9},
                # ...si ce a mai apucat sa ia din loturile de AICI.
                "dinLotAnterior": {k: v for k, v in desf["baraReport"]["dinLotCurent"].items() if v > 1e-9},
                "calc": calc,
                "retetaSursa": r["nume"],
            }

        plan.append(intrare)
        report = intrare["reportOut"]
        idx += 1

        # Mai sunt bare de plasat dar nu mai exista reteta urmatoare?
        if idx >= len(order["retete"]) and (total - asignate) > 0 and creeaza_retete:
            order["retete"].append(reteta_continuare(order, r))

    return plan


def aplica_plan(order, plan):
    """Rescrie listele de bare ale retetelor conform planului.

    Barele existente sunt REFOLOSITE (se pastreaza id-ul si eventualele
    date deja inregistrate), ca sa nu se piarda istoricul la fiecare
    redistribuire.
    """
    for intrare in plan:
        r = intrare["reteta"]
        aplicate = [b for b in r["bare"] if b.get("consumApplied")]
        libere = [b for b in r["bare"] if not b.get("consumApplied")]

        def ia_bara():
            return libere.pop(0) if libere else {"id": uid(), "consumApplied": False}

        noi = []

        # 1. Bara de report: pastreaza dozarea retetei ANTERIOARE.
        if intrare["report"]:
            bar = ia_bara()
            bar["tipDozare"] = "report"
            bar["reportRamas"] = intrare["report"]["ramas"]
            bar["reportDinLotAnterior"] = intrare["report"]["dinLotAnterior"]
            bar["calcSnapshot"] = intrare["report"]["calc"]
            bar["retetaSursaDozare"] = intrare["report"]["retetaSursa"]
            noi.append(bar)

        # 2. Barele normale: se dozeaza pe loturile ACESTEI retete.
        for _ in range(intrare["bareNoi"]):
            bar = ia_bara()
            bar["tipDozare"] = "normal"
            for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare"):
                bar.pop(cheie, None)
            noi.append(bar)

        # 3. Bare in asteptare (reteta inca fara loturile selectate).
        for _ in range(intrare["bareAsteptare"]):
            bar = ia_bara()
            bar["tipDozare"] = "asteptare"
            for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare"):
                bar.pop(cheie, None)
            noi.append(bar)

        r["bare"] = aplicate + noi


# ---------------------------------------------------------------------------
# Desfasurarea unei singure retete
# ---------------------------------------------------------------------------

def desfasurare_reteta(order, r, lots, calculeaza_bara, tip_aliaj):
    """(calc, desf, lot_sel, eroare) pentru o singura reteta.

    Bara de report mostenita nu se numara printre barele dozate aici: ea
    are dozarea ei veche, care se scade separat, inaintea celorlalte.
    """
    lot_sel = loturi_selectate(lots, r)
    calc = calculeaza_bara(
        tip_aliaj, r["target"], lot_sel,
        r.get("portie"), r.get("numarPresari") or 1,
    )
    if calc.get("eroare"):
        return None, None, None, calc["eroare"]

    report = {k: 0.0 for k in ORDINE_MAT}
    for bar in r["bare"]:
        if bar.get("consumApplied"):
            continue
        for k, kg in (bar.get("reportRamas") or {}).items():
            report[k] = report.get(k, 0.0) + to_float(kg)

    normale = [
        b for b in r["bare"]
        if not b.get("consumApplied") and not b.get("reportRamas")
    ]

    # Daca reteta a cedat deja o bara retetei urmatoare, acea bara nu mai
    # apare in lista de aici — dar loturile de AICI ii datoreaza restul
    # ramas din materialul care s-a terminat. O cerem inapoi in numarul de
    # bare, ca desfasurarea sa goleasca lotul epuizat, exact ca in Excel.
    _, bara_cedata = bara_report_trimisa(order, r)
    cedeaza = 1 if (bara_cedata is not None and not bara_cedata.get("consumApplied")) else 0

    resturi = {k: (lot_rest(lot_sel[k]) if lot_sel[k] else 0.0) for k in ORDINE_MAT}
    desf = desfasoara_reteta(calc["rezultatTotal"], resturi, len(normale) + cedeaza, report)
    return calc, desf, lot_sel, None


# ---------------------------------------------------------------------------
# Aplicarea consumului pe o reteta
# ---------------------------------------------------------------------------

def aplica_consum_reteta(order, r, lots, calculeaza_bara, tip_aliaj):
    """Aplica O SINGURA DATA consumul intregii retete pe loturi.

    Bara de report (mostenita din reteta anterioara) e cazul special:
    materialul care s-a terminat in reteta anterioara isi completeaza
    acum diferenta din loturile de aici, iar celelalte materiale ale ei
    — neatinse pana acum — se consuma tot acum, integral, tot din
    loturile de aici, dar cu DOZAREA VECHE (calcSnapshot al barei).
    Ambele sunt deja cuprinse in desf["reportAcoperit"].

    Returneaza (ok, mesaj).
    """
    calc, desf, lot_sel, eroare = desfasurare_reteta(order, r, lots, calculeaza_bara, tip_aliaj)
    if eroare:
        return False, eroare

    neaplicate = [b for b in r["bare"] if not b.get("consumApplied")]
    # O reteta poate fi FARA nicio bara a ei (tot ce a putut da s-a dus pe
    # bara de report cedata retetei urmatoare — lotul era prea mic pentru
    # macar o bara intreaga, ex. un rest mic de burete). O astfel de reteta
    # tot are consum real de aplicat (restul lotului mic, creditat barei de
    # report de la reteta urmatoare), deci NU e cazul "nimic de aplicat"
    # decat daca nici cedarea aia nu mai e in asteptare.
    _, bara_cedata_verif = bara_report_trimisa(order, r)
    cedare_in_asteptare = (
        bara_cedata_verif is not None and not bara_cedata_verif.get("consumApplied")
    )
    if not neaplicate and not cedare_in_asteptare:
        return False, "Toate barele acestei retete au deja consum aplicat."

    # --- 1. Scaderea efectiva din loturile curente ---------------------
    # Se retine EXACT cat s-a scazut din fiecare lot, pe reteta: retragerea
    # foloseste aceleasi cifre, deci stocul se intoarce fix de unde a
    # plecat, fara resturi de rotunjire de cate 10 grame.
    aplicat_pe_lot = {}
    for k in ORDINE_MAT:
        kg = r2(desf["consumTotalReteta"].get(k, 0.0))
        lot = lot_sel[k]
        if lot is None:
            if kg > TOLERANTA_KG and material_necesar(k, r["target"], tip_aliaj):
                return False, f"Lotul pentru {k} nu a fost selectat."
            continue
        if kg <= 1e-9:
            continue
        lot["consum"] = r2(to_float(lot.get("consum")) + kg)
        aplicat_pe_lot[lot["id"]] = r2(aplicat_pe_lot.get(lot["id"], 0.0) + kg)
    r["consumAplicat"] = aplicat_pe_lot

    # --- 2. Bara CEDATA retetei urmatoare ------------------------------
    # Restul pe care l-a luat din loturile de AICI s-a scazut adineauri din
    # stoc, dar bara nu e inca inregistrata nicaieri (ea sta in reteta
    # urmatoare si asteapta sa fie rezolvata acolo). Ii scriem acum partea
    # asta in snapshot, altfel cantitatea nu apare pe Fisa limita.
    _, bara_cedata = bara_report_trimisa(order, r)
    if bara_cedata is not None and not bara_cedata.get("consumApplied") and desf["baraReport"]:
        din_lot_anterior = desf["baraReport"]["dinLotCurent"]
        snap_cedata = {}
        retinut = {}
        for k, kg in din_lot_anterior.items():
            kg = r2(kg)
            if kg <= 1e-9 or lot_sel[k] is None:
                continue
            snap_cedata[k] = [{"lotId": lot_sel[k]["id"], "kg": kg}]
            retinut[k] = kg
        bara_cedata["consumSnapshot"] = snap_cedata
        bara_cedata["reportDinLotAnterior"] = retinut

    # --- 3. Inregistrarea pe barele retetei (istoric, Fisa limita) -----
    doza = calc["rezultatTotal"]
    index_pas = 0
    for bar in r["bare"]:
        if bar.get("consumApplied"):
            continue

        # Bara de report isi stinge datoria fata de loturile de aici.
        if bar.get("reportRamas"):
            snapshot = bar.get("consumSnapshot") or {}
            for k in ORDINE_MAT:
                kg = desf["reportAcoperit"].get(k, 0.0)
                if kg <= 1e-9 or lot_sel[k] is None:
                    continue
                snapshot.setdefault(k, [])
                snapshot[k].append({"lotId": lot_sel[k]["id"], "kg": r2(kg)})

            ramas_dupa = {k: v for k, v in desf["reportLipsa"].items() if v > 1e-9}
            if ramas_dupa:
                # Nu ajunge nici din loturile de aici -> ramane in
                # asteptare si coboara mai departe (caz rar).
                bar["consumSnapshot"] = snapshot
                bar["reportRamas"] = ramas_dupa
                continue

            bar["consumSnapshot"] = snapshot
            bar.pop("reportRamas", None)
            bar["consumApplied"] = True
            continue

        pas = desf["pasi"][index_pas] if index_pas < len(desf["pasi"]) else None
        index_pas += 1
        if pas is None or not pas["incape"]:
            continue

        # Cantitatile se retin nerotunjite: rotunjirea se face la afisare
        # si pe total, ca suma barelor sa dea exact consumul scazut din lot.
        bar["consumSnapshot"] = {
            k: ([{"lotId": lot_sel[k]["id"], "kg": doza.get(k, 0.0)}]
                if lot_sel[k] and doza.get(k, 0.0) > 1e-9 else [])
            for k in ORDINE_MAT
        }
        bar["consumApplied"] = True
        bar["calcSnapshot"] = calc

    return True, "Consumul a fost aplicat pe toata reteta."


# ---------------------------------------------------------------------------
# Retragerea consumului de pe o reteta
# ---------------------------------------------------------------------------

def bara_report_trimisa(order, r):
    """(reteta_urmatoare, bara) daca ACEASTA reteta a trimis mai departe o
    bara de report (o bara de-a ei nu a incaput toata aici si a plecat, cu
    dozarea veche, pe reteta urmatoare). (None, None) daca nu e cazul."""
    try:
        idx = order["retete"].index(r)
    except ValueError:
        return None, None
    if idx + 1 >= len(order["retete"]):
        return None, None
    r_urm = order["retete"][idx + 1]
    bar = next(
        (b for b in r_urm["bare"] if b.get("retetaSursaDozare") == r["nume"]),
        None,
    )
    return r_urm, bar


def stare_retragere(order, r):
    """Ce se poate retrage de pe o reteta, ca sa stie interfata ce buton
    sa arate si ce sa scrie in confirmare.

    Returneaza un dict cu:
      bareAplicate   -- cate bare de-ale retetei au consum aplicat aici
      reportTrimis   -- True daca reteta a cedat o bara mai departe si
                        acea bara a apucat sa ia deja material de aici
      retetaUrmatoare / baraTrimisa -- referintele respective
      blocat         -- True daca bara cedata a fost DEJA rezolvata pe
                        reteta urmatoare (se retrage intai de acolo)
      areCeva        -- True daca exista ceva de retras
    """
    aplicate = [b for b in r["bare"] if b.get("consumApplied")]
    r_urm, bar_out = bara_report_trimisa(order, r)
    report_trimis = bool(bar_out and bar_out.get("reportDinLotAnterior"))
    return {
        "consumAplicat": r.get("consumAplicat") or {},
        "bareAplicate": len(aplicate),
        "listaAplicate": aplicate,
        "reportTrimis": report_trimis,
        "retetaUrmatoare": r_urm,
        "baraTrimisa": bar_out,
        "blocat": bool(bar_out and bar_out.get("consumApplied")),
        "areCeva": bool(aplicate) or report_trimis or bool(r.get("consumAplicat")),
    }


def retrage_consum_reteta(order, r, lots):
    """Retrage tot consumul aplicat pe stoc de o reteta — simetric cu
    ``aplica_consum_reteta``: repune in loturi tot ce s-a scazut de acolo.

    Cazul special e bara de report: daca reteta a "cedat" o bara pe reteta
    urmatoare, acea bara a consumat deja, direct din loturile ACESTEI
    retete, restul care s-a mai gasit (``reportDinLotAnterior``) — fara sa
    fie inregistrata pe nicio bara de-a acestei retete. La retragere, acea
    cantitate se repune aici, iar bara redevine "in asteptare" pe reteta
    urmatoare.

    Daca bara cedata a fost deja REZOLVATA pe reteta urmatoare, retragerea
    se blocheaza — se retrage intai consumul de acolo (ordine inversa).

    Returneaza (ok, mesaj).
    """
    stare = stare_retragere(order, r)
    if not stare["areCeva"]:
        return False, "Aceasta reteta nu are niciun consum aplicat pe stoc."
    if stare["blocat"]:
        return False, (
            "Bara de report trimisa de aceasta reteta a fost deja consumata pe "
            f"reteta \u201e{stare['retetaUrmatoare']['nume']}\u201d. Retrage mai intai "
            "consumul de acolo, apoi revino la aceasta reteta."
        )

    # --- 0. Repunerea stocului ----------------------------------------
    # Varianta exacta: cifrele memorate la aplicare, pe reteta. (Retetele
    # consumate inainte de introducerea acestui camp nu le au — pentru ele
    # se cade inapoi pe snapshot-urile de pe fiecare bara, mai jos.)
    consum_exact = stare["consumAplicat"]
    for lot_id, kg in consum_exact.items():
        lot = _lot_dupa_id(lots, lot_id)
        if lot:
            lot["consum"] = max(0.0, r2(to_float(lot.get("consum")) - to_float(kg)))
    r.pop("consumAplicat", None)

    # --- 1. Bare cu consum aplicat AICI (normale, mutate manual, sau
    #        bare de report venite din reteta anterioara si rezolvate
    #        chiar in aceasta reteta). --------------------------------
    for bar in stare["listaAplicate"]:
        era_report_aici = bar.get("tipDozare") == "report" and bar.get("reportDinLotAnterior")
        if not consum_exact:
            for parti in (bar.get("consumSnapshot") or {}).values():
                intrari = parti if isinstance(parti, list) else [parti]
                for snap in intrari:
                    lot = _lot_dupa_id(lots, snap["lotId"])
                    if lot:
                        lot["consum"] = max(0.0, r2(to_float(lot.get("consum")) - to_float(snap.get("kg"))))
        bar["consumApplied"] = False
        if era_report_aici:
            # Partea luata din loturile retetei ANTERIOARE ramane
            # inregistrata pe bara: ea nu se retrage de aici, ci abia cand
            # se retrage consumul retetei anterioare.
            lot_ids_aici = {l["id"] for l in loturi_selectate(lots, r).values() if l}
            pastrat = {}
            for k, parti in (bar.get("consumSnapshot") or {}).items():
                intrari = parti if isinstance(parti, list) else [parti]
                ramase = [x for x in intrari if x.get("lotId") not in lot_ids_aici]
                if ramase:
                    pastrat[k] = ramase
            bar["consumSnapshot"] = pastrat or None
        else:
            bar["consumSnapshot"] = None
        if era_report_aici:
            # Redevine bara de report NEREZOLVATA: pastreaza dozarea
            # veche (calcSnapshot) si reconstruieste ce mai are de luat
            # din loturile de aici — la materialul limitant diferenta,
            # la celelalte doza veche intreaga.
            doza_veche = doza_veche_bara(bar)
            din_anterior = bar.get("reportDinLotAnterior") or {}
            ramas = {}
            for k in ORDINE_MAT:
                kg = r2(to_float(doza_veche.get(k, 0.0)) - to_float(din_anterior.get(k, 0.0)))
                if kg > 1e-9:
                    ramas[k] = kg
            bar["reportRamas"] = ramas
        else:
            bar.pop("calcSnapshot", None)

    # --- 2. Bara de report TRIMISA mai departe de aceasta reteta:
    #        repune partea deja luata din loturile de aici si o
    #        redeschide (revine "in asteptare"). ---------------------
    if stare["reportTrimis"]:
        bar_out = stare["baraTrimisa"]
        if not consum_exact:
            # Partea pe care bara cedata a luat-o din loturile de AICI nu e
            # inregistrata pe nicio bara a acestei retete — se repune acum.
            lot_sel = loturi_selectate(lots, r)
            for k, kg in bar_out["reportDinLotAnterior"].items():
                kg = to_float(kg)
                lot = lot_sel.get(k)
                if lot and kg > 1e-9:
                    lot["consum"] = max(0.0, r2(to_float(lot.get("consum")) - kg))
        for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare", "calcSnapshot"):
            bar_out.pop(cheie, None)
        bar_out["tipDozare"] = "asteptare"
        bar_out["consumApplied"] = False
        bar_out["consumSnapshot"] = None

    return True, "Consumul retetei a fost retras de pe stoc."


# ---------------------------------------------------------------------------
# Generarea AUTOMATA a retetelor din loturile alocate comenzii
# ---------------------------------------------------------------------------
# Aici nu mai exista pasul manual "adauga reteta -> alege loturile". Comanda
# primeste, pe fiecare material, o LISTA ORDONATA de loturi (o coada), iar
# programul coboara singur prin ea: se toarna bare din loturile curente pana
# cand unul se termina, bara care nu mai incape trece pe reteta urmatoare cu
# dozarea veche, materialul epuizat se inlocuieste cu urmatorul lot din lista
# si tot asa, pana se plaseaza toate barele comenzii.

TARGET_IMPLICIT = {"al": 6.3, "v": 4.05, "o": 0.175, "fe": 0.17}


def config_comanda(order):
    """Portia, numarul de presari si compozitia tinta ale comenzii.

    Daca nu au fost setate inca pe comanda, se iau de pe prima reteta
    existenta (comenzi facute inainte de generarea automata), iar in
    ultima instanta se folosesc valorile implicite pentru Ti5.
    """
    prima = (order.get("retete") or [{}])[0]
    target = order.get("target") or prima.get("target") or dict(TARGET_IMPLICIT)
    portie = order.get("portie") or prima.get("portie") or ""
    presari = order.get("numarPresari") or prima.get("numarPresari") or "1"
    return dict(target), portie, presari


def loturi_alocate(order):
    """Cozile de loturi ale comenzii: {material_id: [lot_id, ...]}."""
    alocate = order.get("loturiAlocate") or {}
    return {k: list(alocate.get(k) or []) for k in ORDINE_MAT}


def genereaza_retete(order, lots, calculeaza_bara, tip_aliaj):
    """Construieste TOATE retetele comenzii, singur, din cozile de loturi.

    Regula e cea din capacitate.py: se umple reteta curenta pana cand un
    lot se termina; bara care nu mai incape trece pe reteta urmatoare cu
    dozarea veche (luand restul lotului epuizat), iar pentru materialul
    terminat se intra automat in urmatorul lot din lista comenzii.

    Retetele de la inceput care au TOT consumul aplicat pe stoc raman
    neatinse — se regenereaza doar ce nu s-a consumat inca.

    Returneaza (retete, raport), unde raport are cheile:
      pastrate    -- cate retete au ramas neatinse (consum deja aplicat)
      generate    -- cate retete noi s-au construit
      barePlasate -- cate bare au fost plasate in total
      bareRamase  -- cate bare NU au putut fi plasate
      avertismente-- lista de mesaje (ex. lista de loturi s-a terminat)
    """
    total = max(0, int(to_float(order.get("nrBare"))))
    target, portie, presari = config_comanda(order)
    cozi = loturi_alocate(order)
    lots_by_id = {l["id"]: l for l in lots}
    avertismente = []

    # --- Retetele deja consumate raman asa cum sunt --------------------
    pastrate = []
    for r in order.get("retete", []):
        if r.get("bare") and all(b.get("consumApplied") for b in r["bare"]):
            pastrate.append(r)
        else:
            break
    restante = order.get("retete", [])[len(pastrate):]

    # Bara pe care ultima reteta pastrata a cedat-o mai departe: se
    # REFOLOSESTE ca obiect, ca sa nu se piarda ce are deja inregistrat
    # (partea luata din lotul golit acolo).
    report = None
    bara_report_existenta = None
    if pastrate:
        nume_ultima = pastrate[-1]["nume"]
        for r in restante:
            for b in r.get("bare", []):
                if (b.get("reportRamas") and not b.get("consumApplied")
                        and b.get("retetaSursaDozare") == nume_ultima):
                    bara_report_existenta = b
                    report = {
                        "ramas": dict(b["reportRamas"]),
                        "dinLotAnterior": dict(b.get("reportDinLotAnterior") or {}),
                        "calc": b.get("calcSnapshot"),
                        "retetaSursa": nume_ultima,
                    }
                    break
            if report:
                break

    # Barele libere din retetele care se regenereaza se refolosesc (isi
    # pastreaza id-ul), ca sa nu se piarda istoricul lor.
    rezerva = [
        b for r in restante for b in r.get("bare", [])
        if not b.get("consumApplied") and b is not bara_report_existenta
    ]

    def ia_bara():
        return rezerva.pop(0) if rezerva else {"id": uid(), "consumApplied": False}

    # --- Pozitia curenta in fiecare coada de loturi --------------------
    idx = {k: 0 for k in ORDINE_MAT}
    if pastrate:
        for k in ORDINE_MAT:
            ultim = pastrate[-1]["lotSel"].get(k)
            if ultim in cozi[k]:
                idx[k] = cozi[k].index(ultim)

    consum_planificat = {}

    def disponibil(lot):
        return r2(lot_rest(lot) - consum_planificat.get(lot["id"], 0.0))

    retete_noi = []
    bare_ramase = total - sum(len(r.get("bare", [])) for r in pastrate)
    nr_reteta = len(pastrate)

    while bare_ramase > 0 and len(retete_noi) < 60:
        # Lotul curent al fiecarui material: primul din coada care mai are
        # ceva in el (cele golite se sar automat — asta e "intrarea in
        # urmatorul lot" cand unul se termina).
        lot_sel = {}
        for k in ORDINE_MAT:
            while idx[k] < len(cozi[k]):
                lot = lots_by_id.get(cozi[k][idx[k]])
                if lot is not None and disponibil(lot) > TOLERANTA_KG:
                    break
                idx[k] += 1
            lot_sel[k] = lots_by_id.get(cozi[k][idx[k]]) if idx[k] < len(cozi[k]) else None

        lipsa = [k for k in ORDINE_MAT if material_lot_lipsa(k, target, tip_aliaj, lot_sel)]
        if lipsa:
            avertismente.append(
                "S-a terminat lista de loturi alocate pentru: "
                + ", ".join(nume_material(k) for k in lipsa)
                + f". Mai raman {bare_ramase} bare de plasat \u2014 adauga loturi "
                  "in configurarea comenzii si genereaza din nou."
            )
            break

        calc = calculeaza_bara(tip_aliaj, target, lot_sel, portie, presari)
        if calc.get("eroare"):
            avertismente.append(f"Reteta {nr_reteta + 1}: {calc['eroare']}")
            break

        disponibile = max(0, bare_ramase - (1 if report else 0))
        resturi = {k: (disponibil(lot_sel[k]) if lot_sel[k] else 0.0) for k in ORDINE_MAT}
        desf = desfasoara_reteta(
            calc["rezultatTotal"], resturi, disponibile,
            report["ramas"] if report else None,
        )

        if desf["bareIntregi"] == 0 and report is None and desf["baraReport"] is None:
            avertismente.append(
                f"In loturile disponibile nu mai incape nicio bara intreaga. "
                f"Mai raman {bare_ramase} bare de plasat."
            )
            break

        # --- Reteta propriu-zisa --------------------------------------
        nr_reteta += 1
        r_noua = {
            "id": uid(),
            "nume": f"Reteta {nr_reteta}",
            "target": dict(target),
            "portie": portie,
            "numarPresari": presari,
            "lotSel": {k: lot_sel[k]["id"] for k in ORDINE_MAT if lot_sel[k]},
            "bare": [],
        }

        bare = []
        if report:
            bar = bara_report_existenta or ia_bara()
            bara_report_existenta = None
            bar["tipDozare"] = "report"
            bar["reportRamas"] = report["ramas"]
            bar["reportDinLotAnterior"] = report["dinLotAnterior"]
            bar["calcSnapshot"] = report["calc"]
            bar["retetaSursaDozare"] = report["retetaSursa"]
            bare.append(bar)
        for _ in range(desf["bareIntregi"]):
            bar = ia_bara()
            bar["tipDozare"] = "normal"
            for cheie in ("reportRamas", "reportDinLotAnterior", "retetaSursaDozare"):
                bar.pop(cheie, None)
            bare.append(bar)
        r_noua["bare"] = bare
        retete_noi.append(r_noua)
        bare_ramase -= len(bare)

        # --- Ce promite reteta asta din loturile ei -------------------
        for k in ORDINE_MAT:
            if lot_sel[k] is None:
                continue
            consum_planificat[lot_sel[k]["id"]] = r2(
                consum_planificat.get(lot_sel[k]["id"], 0.0)
                + desf["consumTotalReteta"].get(k, 0.0)
            )

        # --- Bara care nu mai incape trece mai departe ----------------
        if desf["baraReport"]:
            report = {
                "ramas": {k: v for k, v in desf["baraReport"]["dinLotNou"].items() if v > 1e-9},
                "dinLotAnterior": {k: v for k, v in desf["baraReport"]["dinLotCurent"].items() if v > 1e-9},
                "calc": calc,
                "retetaSursa": r_noua["nume"],
            }
        else:
            report = None
            if bare_ramase > 0:
                # Nu s-a terminat niciun lot, dar nici barele nu s-au
                # plasat toate: loturile curente nu mai au ce da.
                avertismente.append(
                    f"Mai raman {bare_ramase} bare neplasate dupa {r_noua['nume']}."
                )
                break

    retete = pastrate + retete_noi
    return retete, {
        "pastrate": len(pastrate),
        "generate": len(retete_noi),
        "barePlasate": sum(len(r.get("bare", [])) for r in retete),
        "bareRamase": max(0, bare_ramase),
        "avertismente": avertismente,
    }
