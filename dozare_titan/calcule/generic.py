"""Motor generic de dozare, pe baza unei balante de masa COMPLETE (curata) —
fara particularitatile de precedenta a operatorilor gasite in unele CDA-uri
vechi (unde doar primul termen al unei formule se imparte la %continut, iar
scaderile raman ne-impartite; vezi vechile ti6al4v.py/vt9.py/ti5.py, acum
inlocuite de acest motor + calcule/specificatii.py).

Ideea: o reteta e o lista ordonata de PASI. Fiecare pas e un material care
acopera un anumit element din tinta chimica. Un pas poate scadea, inainte sa
imparta la propriul %continut, contributia acelui element deja adusa de toate
materialele calculate INAINTEA lui (aliaje-mama, alte metale, sau chiar
buretele) — asta e diferenta fata de particularitatile gasite: aici se scade
INTAI, se imparte DUPA, intotdeauna.

Doua categorii de pasi:
  - "principal" (implicit)  -> intra in incarcatura principala; se aduna la
    burete cand se calculeaza restul (portie - toate materialele principale)
  - "supliment" (dupa_burete=True) -> se calculeaza DUPA burete, ca TiO2 sau
    Fe metal: se scade contributia adusa de burete + toate materialele
    principale, apoi se imparte complet; NU se scade din burete (e adaugat
    PESTE portie, la fel ca in toate CDA-urile analizate).

Ordinea materialelor in spec["pasi"] CONTEAZA: un material care trebuie sa
scada contributia altuia (ex. Al metal scade Al-ul adus de Aliajul Al-V)
trebuie listat DUPA acela.

Validat pe date reale (CDA VT9, Ti6242, Ti-834) fata de Excel-urile deja
calculate — vezi test_generic.py din pachetul motor_test livrat separat.
"""

from ..utils import to_float


def calculeaza_generic(target, comps, p, spec):
    """Calculeaza dozarea pentru o portie de p kg, dupa o specificatie
    declarativa de materiale (spec).

    Parametri:
      target -- compozitia chimica tinta (dict), chei = orice element
                folosit in spec (ex. "al", "v", "mo", "zr", "sn", "si",
                "o", "fe", "c" — in procente)
      comps  -- compozitia (%) fiecarui lot selectat, indexata dupa
                material: comps[material_id][Element] (ex. comps["burete"]["O"]).
                Element se scrie cu majuscula initiala (Al, V, Mo, Zr, Sn,
                Si, O, Fe, C), la fel ca target dar capitalizat.
      p      -- portia, in kg
      spec   -- descrierea retetei:
          {
            "burete": "burete",          # id-ul materialului burete
            "pasi": [
                {"id": "aliajAlV", "element": "V"},
                {"id": "aliajAlMo", "element": "Mo"},
                {"id": "alMetal", "element": "Al"},      # dupa aliajele-mama
                {"id": "zrMetal", "element": "Zr"},
                {"id": "aliajSiTi", "element": "Si"},
                {"id": "tio2", "element": "O", "dupa_burete": True},
                {"id": "feMetal", "element": "Fe", "dupa_burete": True},
            ],
          }

    Returneaza (rezultat, eroare):
      rezultat -- dict {material_id: kg, rotunjit la 3 zecimale}, sau None
      eroare   -- mesaj de eroare (string), sau None daca s-a calculat OK
    """
    burete_id = spec["burete"]
    pasi = spec["pasi"]

    rezultat = {}
    calculate = []  # id-urile deja calculate, in ordine (fara burete inca)

    def tinta_kg(element):
        return to_float(target.get(element.lower())) / 100 * p

    def contributie(material_id, element):
        return rezultat[material_id] * comps[material_id].get(element, 0) / 100

    def calculeaza_pas(mid, element):
        t_kg = tinta_kg(element)
        if abs(t_kg) < 1e-9:
            return 0.0, None
        continut = comps[mid].get(element, 0)
        if abs(continut) < 1e-9:
            return None, f"Lotul de {mid} nu are %{element} definit — nu se poate calcula."
        deja_adus = sum(contributie(prev, element) for prev in calculate)
        return (t_kg - deja_adus) / (continut / 100), None

    # 1. Toti pasii "principali" (inainte de burete)
    principale_ids = []
    for pas in pasi:
        if pas.get("dupa_burete"):
            continue
        mid, el = pas["id"], pas["element"]
        val, err = calculeaza_pas(mid, el)
        if err:
            return None, err
        rezultat[mid] = val
        calculate.append(mid)
        principale_ids.append(mid)

    # 2. Burete = portie - toate materialele principale (balanta de masa)
    rezultat[burete_id] = p - sum(rezultat[i] for i in principale_ids)
    calculate.append(burete_id)

    # 3. Suplimente (dupa burete): TiO2, Fe metal etc. — adaugate PESTE portie
    for pas in pasi:
        if not pas.get("dupa_burete"):
            continue
        mid, el = pas["id"], pas["element"]
        val, err = calculeaza_pas(mid, el)
        if err:
            return None, err
        rezultat[mid] = val
        calculate.append(mid)

    return {k: round(v, 3) for k, v in rezultat.items()}, None
