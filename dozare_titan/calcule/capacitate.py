"""Capacitatea unei retete: cate bare incap in loturile selectate.

Numarul total de bare e CUNOSCUT DINAINTE, pe COMANDA (ex. 12 bare), iar
retetele se umplu una dupa alta, in ordine. Aici NU se valideaza fiecare
bara separat cu consum aplicat pe stoc: se face doar DESFASURAREA, bara
cu bara, cu consumul pe care il aduce fiecare la pasul ei si cu restul
ramas din fiecare lot dupa acel pas. Consumul se aplica apoi O SINGURA
DATA, pentru toata reteta.

REGULA BAREI CARE NU MAI INCAPE (validata pe Cda 26858-Ti5, tab "Retete"
al Excel-ului de productie):

  - bara respectiva NU se considera partial executata in reteta curenta;
  - materialul (sau materialele) care CHIAR s-au terminat isi cedeaza
    restul ramas, ca sa nu se piarda nimic pe stoc — lotul vechi ramane
    cu rest 0, iar diferenta pana la doza intreaga se ia din lotul nou
    (adica din reteta urmatoare);
  - TOATE celelalte materiale ale acelei bare raman NEATINSE in reteta
    curenta si se consuma abia in reteta urmatoare, din loturile de
    acolo (care pot fi chiar aceleasi loturi, daca nu s-au schimbat);
  - bara isi PASTREAZA DOZAREA VECHE (cea calculata pe loturile retetei
    din care pleaca), pentru ca reteta de presare era deja stabilita in
    momentul in care s-a constatat ca nu mai ajunge materialul.

De aici rezulta un lucru esential pentru planificare, care nu era tratat
corect inainte: bara de report consuma din loturile retetei URMATOARE nu
doar diferenta de material limitant, ci DOZA INTREAGA la toate celelalte
materiale. Reteta urmatoare trebuie deci sa scada tot acest consum
INAINTE sa calculeze cate bare noi mai incap in loturile ei — altfel iese
cu o bara in plus (pe Cda 26858, reteta 2 raporta 6 bare noi in loc de 5,
adica materialul parea ca se termina la bara 11, nu la bara 10).

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
      dozare_bara   -- dict {material_id: kg} consumati de O bara dozata
                       pe loturile ACESTEI retete, cu presarile ei incluse
                       (adica exact calc["rezultatTotal"])
      resturi       -- dict {material_id: kg} disponibili ACUM in loturile
                       selectate in reteta (lot_rest pentru fiecare)
      nr_bare       -- cate bare NOI (in afara barei de report) se cere sa
                       intre in aceasta reteta
      consum_report -- optional, dict {material_id: kg} pe care bara de
                       report MOSTENITA din reteta anterioara il mai are de
                       consumat din loturile ACESTEI retete. Contine doza
                       VECHE a barei minus cat s-a mai putut lua din
                       loturile vechi: la materialul care s-a terminat
                       acolo doar diferenta, la TOATE celelalte doza
                       intreaga. Se scade inaintea barei 1, pentru ca
                       reportul are prioritate fata de barele noi.

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

    # Daca nici loturile de aici nu acopera bara de report mostenita,
    # reteta nu mai poate primi nicio bara noua si NU are voie sa cedeze
    # inca una mai departe (ar insemna doua bare de report deodata).
    report_neacoperit = any(v > TOLERANTA_KG for v in report_lipsa.values())

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
    bare_intregi = 0 if report_neacoperit else min(nr_bare, capacitate_max)

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
            # Restul se tine NEROTUNJIT pe parcurs si se rotunjeste doar la
            # afisare: altfel, dupa 6-7 bare, jumatatile de gram adunate ar
            # face ca bilantul retetei sa nu mai cada pe cantitatile
            # inregistrate bara cu bara (si invers).
            ramas[k] = inainte - doza[k]
            pas["materiale"][k] = {
                "necesar": r2(doza[k]),
                "restInainte": r2(inainte),
                "dinLotCurent": r2(doza[k]),
                "dinLotNou": 0.0,
                "restDupa": r2(ramas[k]),
                "incape": True,
            }
        pasi.append(pas)

    # --- Bara de report NOUA (prima care nu mai incape) ---------------
    # NU se considera partial executata: singurul (singurele) material(e)
    # care se scad acum din loturile curente sunt cele care CHIAR se
    # termina (limitante), si doar cu restul ramas din ele. Toate
    # celelalte materiale raman neatinse aici si se vor consuma INTEGRAL
    # in reteta urmatoare, cu aceeasi dozare veche.
    bara_report = None
    if bare_intregi < nr_bare and not report_neacoperit:
        materiale_pas = {}
        for k in ORDINE_MAT:
            if doza[k] <= TOLERANTA_KG:
                continue
            if k in limitante:
                din_curent = r2(max(0.0, ramas[k]))
                incape_k = False
            else:
                din_curent = 0.0
                incape_k = True
            # Diferenta pana la doza intreaga se ia din reteta urmatoare:
            # la materialele limitante doar cat lipseste, la celelalte
            # doza intreaga (nu s-a atins nimic din ele aici).
            din_nou = r2(doza[k] - din_curent)
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
            # Ce mai ia din loturile ACESTEI retete: doar restul ramas la
            # materialul/materialele limitante (la celelalte e 0 — nu s-a
            # atins nimic din ele).
            "dinLotCurent": {k: v["dinLotCurent"] for k, v in materiale_pas.items()},
            # ...si ce ramane de consumat in reteta URMATOARE, cu aceeasi
            # dozare veche: la limitante doar diferenta, la restul doza
            # intreaga.
            "dinLotNou": {k: v["dinLotNou"] for k, v in materiale_pas.items()},
        }
        for k in limitante:
            ramas[k] = 0.0

    # Consumul total al retetei din loturile curente, ca BILANT: cat era
    # in loturi minus cat a ramas. Asa iese fix, fara resturi fantoma de
    # cate 10 grame din rotunjirile pas cu pas — iar lotul care s-a
    # terminat ramane exact pe zero, nu pe -0.01 kg.
    # (Ca suma, e acelasi lucru: reportul mostenit + barele intregi +
    # restul luat de bara de report la materialele care s-au terminat.)
    consum_total = {
        k: r2(to_float(resturi.get(k)) - ramas[k]) for k in ORDINE_MAT
    }

    return {
        "nrBareCerut": nr_bare,
        "dozareBara": {k: r2(doza[k]) for k in ORDINE_MAT},
        "resturiInitiale": {k: r2(to_float(resturi.get(k))) for k in ORDINE_MAT},
        "reportMostenit": {k: r2(report[k]) for k in ORDINE_MAT},
        "reportAcoperit": report_acoperit,
        "reportLipsa": report_lipsa,
        "reportNeacoperit": report_neacoperit,
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
