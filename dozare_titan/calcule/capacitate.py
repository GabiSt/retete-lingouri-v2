"""Capacitatea unei retete: cate bare incap in loturile selectate.

Numarul de bare al retetei e CUNOSCUT DINAINTE (campul "Numar de bare"),
deci aici NU se valideaza fiecare bara separat cu consum aplicat pe stoc.
Se face doar DESFASURAREA: bara cu bara, cu consumul pe care il aduce
fiecare la pasul ei si cu restul ramas din fiecare lot dupa acel pas.
Consumul se aplica apoi O SINGURA DATA, pentru toata reteta.

Regula pentru bara care nu mai incape (validata pe date reale de
productie, tab "Retete" al unui RetDozare existent): bara respectiva NU
se considera partial executata. Niciun material nu se scade din loturile
curente pentru ea, CU O SINGURA EXCEPTIE: materialul (sau materialele)
care chiar s-au terminat isi cedeaza restul ramas, ca sa nu se piarda —
diferenta se completeaza din lotul nou. Toate celelalte materiale ale
acelei bare raman neatinse si se dozeaza normal, din loturile selectate,
cand bara respectiva ruleaza efectiv in reteta urmatoare (deci NU sunt
tratate ca "report" — continua pur si simplu cu lotul lor, eventual
acelasi lot, daca nu s-a schimbat).

Modulul e PUR: doar calculeaza, nu modifica niciun lot si nicio reteta.
"""

from ..config import ORDINE_MAT, MATERIALE
from ..utils import to_float, r2

# Toleranta in kg sub care consideram ca "incape". Trebuie sa fie peste
# pragul de rotunjire la 2 zecimale (0.005), altfel o bara care iese
# fix pe zero ar fi raportata gresit ca "nu incape".
TOLERANTA_KG = 0.005


def nume_material(mat_id):
    """Numele afisabil al unui material (ex. "burete" -> "Burete de titan")."""
    return next((m["nume"] for m in MATERIALE if m["id"] == mat_id), mat_id)


def desfasoara_reteta(dozare_bara, resturi, nr_bare, consum_report=None):
    """Desfasoara, bara cu bara, consumul adus de fiecare bara.

    Parametri:
      dozare_bara   -- dict {material_id: kg} consumati de O bara, cu
                       presarile ei incluse (adica exact calc["rezultatTotal"])
      resturi       -- dict {material_id: kg} disponibili ACUM in loturile
                       selectate in reteta (lot_rest pentru fiecare)
      nr_bare       -- numarul de bare cerut pe reteta (cunoscut dinainte)
      consum_report -- optional, dict {material_id: kg} pe care o bara de
                       report MOSTENITA din reteta anterioara il mai are de
                       consumat din loturile ACESTEI retete (doar pentru
                       materialul/materialele care s-au terminat acolo).
                       Se scade inaintea barei 1, pentru ca reportul are
                       prioritate fata de barele noi.

    Returneaza dictionarul de desfasurare (vezi cheile la final).
    """
    nr_bare = max(0, int(to_float(nr_bare)))
    doza = {k: to_float(dozare_bara.get(k)) for k in ORDINE_MAT}
    disponibil = {k: to_float(resturi.get(k)) for k in ORDINE_MAT}
    report = {k: to_float((consum_report or {}).get(k)) for k in ORDINE_MAT}

    # --- Pasul 0: bara de report mostenita isi ia partea prima ---------
    report_acoperit = {}
    report_lipsa = {}
    for k in ORDINE_MAT:
        luat = min(report[k], max(0.0, disponibil[k]))
        report_acoperit[k] = r2(luat)
        report_lipsa[k] = r2(report[k] - luat)
        disponibil[k] = r2(disponibil[k] - luat)

    # --- Cate bare INTREGI incap dupa report ---------------------------
    # Limita e data de materialul care se termina primul. Materialele cu
    # doza 0 (element tinta 0%) nu limiteaza nimic si sunt ignorate.
    capacitate_pe_material = {}
    for k in ORDINE_MAT:
        if doza[k] <= TOLERANTA_KG:
            continue
        capacitate_pe_material[k] = int((disponibil[k] + TOLERANTA_KG) // doza[k])

    capacitate_max = min(capacitate_pe_material.values()) if capacitate_pe_material else nr_bare
    capacitate_max = max(0, capacitate_max)
    bare_intregi = min(nr_bare, capacitate_max)

    # Materialele care CHIAR limiteaza. Pot fi mai multe deodata (doua
    # stocuri care se termina la aceeasi bara) — de-aia e lista, nu unul.
    limitante = sorted(
        [k for k, c in capacitate_pe_material.items() if c == capacitate_max],
        key=ORDINE_MAT.index,
    )

    # --- Desfasurarea barelor care CHIAR se executa (bare_intregi) -----
    # Fiecare bara de aici incape integral, la toate materialele — de-aia
    # se scade normal, din toate stocurile deodata.
    ramas = dict(disponibil)
    pasi = []
    for i in range(1, bare_intregi + 1):
        pas = {"bara": i, "materiale": {}, "incape": True, "materialeLipsa": []}
        for k in ORDINE_MAT:
            if doza[k] <= TOLERANTA_KG:
                continue
            inainte = ramas[k]
            ramas[k] = r2(inainte - doza[k])
            pas["materiale"][k] = {
                "necesar": r2(doza[k]),
                "restInainte": r2(inainte),
                "dinLotCurent": r2(doza[k]),
                "dinLotNou": 0.0,
                "restDupa": ramas[k],
                "incape": True,
            }
        pasi.append(pas)

    # --- Bara de report NOUA (prima care nu mai incape) ---------------
    # NU se considera partial executata: singurul (singurele) material(e)
    # care se scad acum din loturile curente sunt cele care CHIAR se
    # termina (limitante) — restul asteapta neatins, sa fie dozat normal
    # cand bara asta ruleaza efectiv in reteta urmatoare.
    bara_report = None
    if bare_intregi < nr_bare:
        materiale_pas = {}
        for k in ORDINE_MAT:
            if doza[k] <= TOLERANTA_KG:
                continue
            if k in limitante:
                din_curent = r2(max(0.0, ramas[k]))
                din_nou = r2(doza[k] - din_curent)
                incape_k = False
            else:
                din_curent = 0.0
                din_nou = 0.0
                incape_k = True
            materiale_pas[k] = {
                "necesar": r2(doza[k]),
                "restInainte": r2(ramas[k]),
                "dinLotCurent": din_curent,
                "dinLotNou": din_nou,
                # restDupa ramane neschimbat pentru materialele care nu
                # limiteaza — nu s-a consumat nimic din ele inca.
                "restDupa": 0.0 if k in limitante else r2(ramas[k]),
                "incape": incape_k,
            }
        pasi.append({
            "bara": bare_intregi + 1,
            "materiale": materiale_pas,
            "incape": False,
            "materialeLipsa": list(limitante),
        })
        bara_report = {
            "index": bare_intregi + 1,
            "materiale": materiale_pas,
            # Ce mai ia din loturile ACESTEI retete (doar materialul/
            # materialele limitante; restul e 0 — nu s-a atins nimic).
            "dinLotCurent": {k: v["dinLotCurent"] for k, v in materiale_pas.items()},
            # ...si ce ramane de consumat din lotul NOU, in reteta urmatoare
            # (tot doar la materialele limitante).
            "dinLotNou": {k: v["dinLotNou"] for k, v in materiale_pas.items()},
        }
        for k in limitante:
            ramas[k] = 0.0

    # Consumul total al retetei din loturile curente: reportul mostenit +
    # barele intregi + resturile luate de bara de report (doar la
    # materialele care s-au terminat chiar aici).
    consum_total = {}
    for k in ORDINE_MAT:
        total = report_acoperit[k] + doza[k] * bare_intregi
        if bara_report:
            total += bara_report["dinLotCurent"].get(k, 0.0)
        consum_total[k] = r2(total)

    return {
        "nrBareCerut": nr_bare,
        "dozareBara": {k: r2(doza[k]) for k in ORDINE_MAT},
        "resturiInitiale": {k: r2(to_float(resturi.get(k))) for k in ORDINE_MAT},
        "reportMostenit": {k: r2(report[k]) for k in ORDINE_MAT},
        "reportAcoperit": report_acoperit,
        "reportLipsa": report_lipsa,
        "capacitatePeMaterial": capacitate_pe_material,
        "capacitateMaxima": capacitate_max,
        "bareIntregi": bare_intregi,
        "bareRamase": nr_bare - bare_intregi,
        "materialeLimitante": limitante,
        "pasi": pasi,
        "baraReport": bara_report,
        "consumTotalReteta": consum_total,
        # Ce ramane cu adevarat disponibil dupa aceasta reteta, pentru
        # materialele care NU au limitat (limitantele ajung la 0, pentru
        # ca bara de report le-a golit complet pe cele ramase).
        "resturiFinale": {k: ramas.get(k, r2(disponibil[k])) for k in ORDINE_MAT},
    }
