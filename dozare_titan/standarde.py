"""Standardele (AMS, ASTM, ISO, ...) pe aliaj, CITITE DIN EXCEL.

Sursa datelor e fisierul ``standarde.xlsx`` (un rand pe standard, o coloana
pe element chimic). Aplicatia il cauta in doua locuri, in aceasta ordine:

  1. ``standarde.xlsx`` asezat LANGA aplicatie (langa dozare_titan.exe sau
     langa main.py) — acesta e fisierul pe care il editezi tu; daca exista,
     are prioritate;
  2. copia livrata odata cu aplicatia, in ``dozare_titan/assets/``.

Fisierul se reciteste AUTOMAT cand se schimba pe disc (se compara data
modificarii la fiecare reconstruire a ecranului), deci dupa ce salvezi
Excel-ul nu trebuie repornita aplicatia.

Formatul asteptat (ca in fisierul actual):
  - randul 1 = antet: "Aliaj", "Standards Requirements", apoi cate o
    coloana pe element (Al, V, Fe, C, O, N, H, Y, Zr, Sn, Mo, Si, Co ...).
    Elementele se iau din antet, deci o coloana noua se poate adauga fara
    sa se schimbe codul;
  - numele aliajului apare doar pe primul rand al grupului (celulele de
    dedesubt raman goale, ca in fisierul actual) — se completeaza in jos;
  - valorile: ``5.50-6.75`` (interval), ``max. 0.30`` / ``max 0.30``
    (maxim), ``min. x`` (minim), ``-`` (fara limita), ``N/A`` (nu se aplica).
    Toate in procente de masa (%).

Ce ofera modulul restului aplicatiei:
  - aliaje_disponibile()  lista pentru combobox: aliajele cu reteta
                          implementata + PLACEHOLDER-e pentru aliajele care
                          apar in Excel dar nu au (inca) reteta;
  - standarde_pentru()    standardele unui aliaj;
  - standarde_alese()     LISTA standardelor alese pe o comanda (se pot alege
                          mai multe, ex. AMS 4975 + AMS 4976);
  - standard_ales()       standardul comenzii: cel ales, sau — daca sunt mai
                          multe — unul COMBINAT (cea mai stricta limita din
                          fiecare element), folosit de formulare si dialoguri;
  - verifica_limite()     ce iese din limitele standardului;
  - spec_efectiv()        limitele pentru formularele tiparite (RetDozare,
                          Fisa limita), luate din standardul ales.

Modulul NU depinde de interfata grafica.
"""

import os
import re

from .config import (
    ALIAJE_SPEC, ALIAJE_EXCEL, _director_aplicatie, _director_resurse,
)

NUME_FISIER = "standarde.xlsx"
TOLERANTA = 1e-6          # in procente; sub ea nu se considera depasire

_NUM = r"(\d+(?:\.\d+)?)"


# ---------------------------------------------------------------------------
# Helpere de text
# ---------------------------------------------------------------------------

def _curata(text):
    """Text pe o singura linie, fara spatii multiple sau la margini."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\u00a0", " ")).strip()


def _cheie(text):
    """Forma de comparare a unui nume (fara diferente de spatii/majuscule)."""
    return _curata(text).casefold()


def _parseaza_limita(valoare):
    """Interpreteaza o celula de limita.

    Returneaza (stare, limita):
      ("nimic", None)   celula goala sau "-"      -> fara limita
      ("na", None)      "N/A"                     -> nu se aplica
      ("ok", {...})     {"min": x} / {"max": y} / ambele
      ("invalid", None) text neinteles            -> se raporteaza
    """
    if valoare is None:
        return "nimic", None
    if isinstance(valoare, (int, float)) and not isinstance(valoare, bool):
        # Un numar "gol" nu spune daca e minim, maxim sau interval.
        return "invalid", None
    t = _curata(valoare).casefold()
    t = t.replace("\u2013", "-").replace("\u2014", "-").replace(",", ".")
    t = re.sub(r"\s+", " ", t)
    if t in ("", "-"):
        return "nimic", None
    if t in ("n/a", "na"):
        return "na", None
    m = re.fullmatch(rf"max\.?\s*{_NUM}", t)
    if m:
        return "ok", {"max": float(m.group(1))}
    m = re.fullmatch(rf"min\.?\s*{_NUM}", t)
    if m:
        return "ok", {"min": float(m.group(1))}
    m = re.fullmatch(rf"{_NUM}\s*-\s*{_NUM}", t)
    if m:
        lo, hi = float(m.group(1)), float(m.group(2))
        if lo <= hi:
            return "ok", {"min": lo, "max": hi}
    return "invalid", None


def _numar(x):
    """0.3 -> "0.30", 5.5 -> "5.50", 0.0125 -> "0.0125" (minim 2 zecimale)."""
    t = f"{x:.4f}".rstrip("0")
    intreg, _, zec = t.partition(".")
    return f"{intreg}.{zec.ljust(2, '0')}"


def text_limita(limita):
    """"5.50-6.75 %" / "max 0.30 %" / "min 1.00 %" — pentru afisare."""
    if not limita:
        return ""
    lo, hi = limita.get("min"), limita.get("max")
    if lo is not None and hi is not None:
        return f"{_numar(lo)}-{_numar(hi)} %"
    if hi is not None:
        return f"max {_numar(hi)} %"
    return f"min {_numar(lo)} %"


# ---------------------------------------------------------------------------
# Citirea fisierului
# ---------------------------------------------------------------------------

def cai_posibile():
    """Locurile in care se cauta fisierul, in ordinea prioritatii."""
    return [
        os.path.join(_director_aplicatie(), NUME_FISIER),
        os.path.join(_director_resurse(), NUME_FISIER),
    ]


def cale_fisier():
    """Prima cale existenta din cai_posibile(), sau None."""
    for cale in cai_posibile():
        if os.path.isfile(cale):
            return cale
    return None


def _catalog_gol(cale=None, eroare=None):
    return {"cale": cale, "aliaje": {}, "avertismente": [], "eroare": eroare,
            "semnatura": None}


def incarca_din_excel(cale):
    """Citeste fisierul si returneaza catalogul:

      {"cale", "eroare", "avertismente": [str],
       "aliaje": {cheie_normalizata: {"nume": str, "standarde": [
            {"nume": str, "limite": {"Al": {"min":..,"max":..}, ...},
             "fara_limite": bool, "neaplicabil": bool}]}}}

    Nu ridica exceptii: orice problema ajunge in ``eroare`` (fisier
    lipsa/blocat/ilizibil) sau in ``avertismente`` (celule neintelese,
    duplicate contradictorii).
    """
    catalog = _catalog_gol(cale)
    try:
        import openpyxl
    except ImportError:
        catalog["eroare"] = "biblioteca 'openpyxl' nu este instalata (pip install openpyxl)"
        return catalog
    try:
        wb = openpyxl.load_workbook(cale, data_only=True, read_only=True)
    except Exception as e:      # fisier deschis in Excel, corupt etc.
        catalog["eroare"] = f"{type(e).__name__}: {e}"
        return catalog

    ws = wb.worksheets[0]
    randuri = list(ws.iter_rows(values_only=True))
    wb.close()
    if not randuri:
        catalog["eroare"] = "fisierul e gol"
        return catalog

    # --- antetul: coloana aliajului, a standardului, apoi elementele ----
    antet = [_curata(c) for c in randuri[0]]
    col_aliaj = next((i for i, h in enumerate(antet) if h.casefold().startswith("aliaj")), 0)
    col_std = next((i for i, h in enumerate(antet) if "standard" in h.casefold()), 1)
    coloane_elem = [(i, h) for i, h in enumerate(antet)
                    if h and i not in (col_aliaj, col_std)]
    if not coloane_elem:
        catalog["eroare"] = "nu am gasit coloane de elemente chimice in antet"
        return catalog

    aliaj_curent = None
    avert = catalog["avertismente"]
    for nr, rand in enumerate(randuri[1:], start=2):
        rand = list(rand) + [None] * (len(antet) - len(rand))
        nume_aliaj = _curata(rand[col_aliaj])
        nume_std = _curata(rand[col_std])
        if nume_aliaj:
            aliaj_curent = nume_aliaj          # se completeaza in jos
        if not nume_std or aliaj_curent is None:
            continue

        limite, stari = {}, []
        for i, element in coloane_elem:
            stare, lim = _parseaza_limita(rand[i])
            stari.append(stare)
            if stare == "ok":
                limite[element] = lim
            elif stare == "invalid":
                avert.append(
                    f"Randul {nr}, {nume_std}, coloana {element}: nu inteleg valoarea "
                    f"\u201e{_curata(rand[i])}\u201d (folosesc formatul \u201e5.5-6.75\u201d, "
                    f"\u201emax. 0.30\u201d sau \u201e-\u201d). Limita ignorata."
                )
        standard = {
            "nume": nume_std,
            "limite": limite,
            "fara_limite": not limite,
            "neaplicabil": bool(stari) and all(s == "na" for s in stari),
        }

        grup = catalog["aliaje"].setdefault(
            _cheie(aliaj_curent), {"nume": aliaj_curent, "standarde": []})
        existent = next((s for s in grup["standarde"]
                         if _cheie(s["nume"]) == _cheie(nume_std)), None)
        if existent is None:
            grup["standarde"].append(standard)
        elif existent["limite"] != limite:
            avert.append(
                f"Randul {nr}: standardul \u201e{nume_std}\u201d apare de doua ori la "
                f"\u201e{aliaj_curent}\u201d cu limite DIFERITE; il folosesc pe primul."
            )
        # duplicat identic (ex. acelasi standard repetat pe mai multe randuri): ignorat

    # --- legaturile din config catre Excel ------------------------------
    for cheie_app, nume_excel in ALIAJE_EXCEL.items():
        if _cheie(nume_excel) not in catalog["aliaje"]:
            avert.append(
                f"Aliajul \u201e{cheie_app}\u201d e legat de \u201e{nume_excel}\u201d, care nu mai "
                f"exista in Excel; nu are standarde pana repari numele (sau "
                f"config.ALIAJE_EXCEL)."
            )
    return catalog


# ---------------------------------------------------------------------------
# Cache cu reincarcare automata
# ---------------------------------------------------------------------------

_cache = {"semnatura": "nimic", "catalog": _catalog_gol()}


def _semnatura(cale):
    if not cale:
        return None
    try:
        st = os.stat(cale)
        return (cale, st.st_mtime, st.st_size)
    except OSError:
        return None


def catalog():
    """Catalogul curent; se reciteste singur daca fisierul s-a schimbat."""
    cale = cale_fisier()
    semnatura = _semnatura(cale)
    if semnatura != _cache["semnatura"]:
        if cale is None:
            cat = _catalog_gol(None, "nu gasesc " + NUME_FISIER + " (nici langa aplicatie, nici in assets)")
        else:
            cat = incarca_din_excel(cale)
        cat["semnatura"] = semnatura
        _cache["semnatura"] = semnatura
        _cache["catalog"] = cat
    return _cache["catalog"]


def reincarca():
    """Forteaza recitirea (ex. dupa ce fisierul a fost inlocuit)."""
    _cache["semnatura"] = "nimic"
    return catalog()


# ---------------------------------------------------------------------------
# Aliaje si standarde pentru interfata
# ---------------------------------------------------------------------------

def _nume_excel_pentru(cheie_aliaj):
    """Numele din Excel al aliajului cu cheia data, sau None.

    Aliajele cu reteta se leaga prin config.ALIAJE_EXCEL; un placeholder are
    chiar numele din Excel drept cheie.
    """
    nume = ALIAJE_EXCEL.get(cheie_aliaj)
    if nume and _cheie(nume) in catalog()["aliaje"]:
        return _cheie(nume)
    if _cheie(cheie_aliaj) in catalog()["aliaje"]:
        return _cheie(cheie_aliaj)
    return None


def este_placeholder(cheie_aliaj):
    """True pentru aliajele din Excel care nu au (inca) reteta de dozare.

    Un aliaj devine "real" cand ii adaugi o intrare in config.ALIAJE_SPEC si
    o formula in calcule.DOZATOARE, sub aceeasi cheie.
    """
    return cheie_aliaj not in ALIAJE_SPEC


def aliaje_disponibile():
    """Lista pentru combobox-ul de aliaj:
       [{"cheie", "nume", "placeholder", "are_standarde"}]

    Intai aliajele cu reteta implementata (ca pana acum), apoi cate un
    placeholder pentru fiecare aliaj din Excel care nu e legat de niciunul.
    """
    cat = catalog()
    rezultat = [
        {"cheie": k, "nume": k, "placeholder": False,
         "are_standarde": bool(standarde_pentru(k))}
        for k in ALIAJE_SPEC
    ]
    legate = {_cheie(n) for n in ALIAJE_EXCEL.values()}
    legate |= {_cheie(k) for k in ALIAJE_SPEC}
    for cheie_norm, grup in cat["aliaje"].items():
        if cheie_norm in legate:
            continue
        rezultat.append({"cheie": grup["nume"], "nume": grup["nume"],
                         "placeholder": True, "are_standarde": True})
    return rezultat


def eticheta_aliaj(item):
    """Textul din combobox: placeholder-ele sunt marcate."""
    return item["nume"] + (" (placeholder)" if item["placeholder"] else "")


def standarde_pentru(cheie_aliaj):
    """Standardele aliajului (lista de dict-uri), sau [] daca nu are."""
    nume = _nume_excel_pentru(cheie_aliaj)
    return list(catalog()["aliaje"][nume]["standarde"]) if nume else []


def nume_standarde_alese(order):
    """Numele standardelor alese pe comanda, in ordinea alegerii.

    Sursa e ``order["standarde"]`` (lista). Comenzile mai vechi au doar
    ``order["standard"]`` (un singur nume, text) — ramane valabil ca rezerva
    cand lista nu exista.
    """
    lista = order.get("standarde")
    if isinstance(lista, (list, tuple)):
        nume = [_curata(n) for n in lista]
    else:
        nume = [_curata(order.get("standard"))]
    rezultat, vazute = [], set()
    for n in nume:
        if n and _cheie(n) not in vazute:
            vazute.add(_cheie(n))
            rezultat.append(n)
    return rezultat


def standarde_alese(order):
    """Standardele alese pe comanda (lista de dict-uri din catalog), in
    ordinea alegerii. Numele care nu exista la aliajul comenzii se ignora."""
    disponibile = {_cheie(s["nume"]): s
                   for s in standarde_pentru(order.get("tipAliaj"))}
    return [disponibile[_cheie(n)] for n in nume_standarde_alese(order)
            if _cheie(n) in disponibile]


def seteaza_standarde(order, nume_standarde):
    """Seteaza standardele comenzii (lista de nume). Pastreaza doar numele
    care exista la aliajul comenzii, fara duplicate. Actualizeaza si campul
    vechi ``order["standard"]`` (primul nume), ca sa ramana coerent.

    Returneaza True daca s-a schimbat ceva.
    """
    de_pastrat = {_cheie(s["nume"]): s["nume"]
                  for s in standarde_pentru(order.get("tipAliaj"))}
    noua, vazute = [], set()
    for n in nume_standarde or []:
        k = _cheie(n)
        if k in de_pastrat and k not in vazute:
            vazute.add(k)
            noua.append(de_pastrat[k])
    schimbat = (nume_standarde_alese(order) != noua
                or order.get("standarde") != noua
                or order.get("standard", "") != (noua[0] if noua else ""))
    order["standarde"] = noua
    order["standard"] = noua[0] if noua else ""
    return schimbat


def pastreaza_standarde_valide(order):
    """Dupa schimbarea aliajului: scoate standardele care nu exista la noul
    aliaj (le pastreaza pe cele comune)."""
    return seteaza_standarde(order, nume_standarde_alese(order))


def combina_standarde(lista):
    """Un standard \"virtual\" din mai multe standarde alese impreuna.

    Materialul trebuie sa respecte TOATE standardele, deci pentru fiecare
    element limita efectiva e cea mai stricta: minimul cel mai mare si
    maximul cel mai mic. Standardele N/A (fara nicio limita) nu restring
    nimic. Daca intr-un element minimul unui standard depaseste maximul
    altuia (nu se poate respecta ambele), elementul apare in ``conflicte``.
    """
    if len(lista) == 1:
        return lista[0]
    limite, conflicte = {}, []
    for std in lista:
        for element, lim in std["limite"].items():
            cur = limite.setdefault(element, {})
            if lim.get("min") is not None:
                cur["min"] = max(cur.get("min", lim["min"]), lim["min"])
            if lim.get("max") is not None:
                cur["max"] = min(cur.get("max", lim["max"]), lim["max"])
    for element, lim in limite.items():
        if lim.get("min") is not None and lim.get("max") is not None \
                and lim["min"] > lim["max"] + TOLERANTA:
            conflicte.append(element)
    return {
        "nume": "; ".join(s["nume"] for s in lista),
        "limite": limite,
        "fara_limite": not limite,
        "neaplicabil": all(s["neaplicabil"] for s in lista),
        "componente": list(lista),
        "conflicte": conflicte,
    }


def standard_ales(order):
    """Standardul comenzii (dict) sau None.

    Cu un singur standard ales: chiar acel standard. Cu mai multe: standardul
    COMBINAT (vezi combina_standarde), astfel incat formularele, dialogurile
    si avertismentele sa lucreze cu un singur obiect, indiferent cate
    standarde s-au ales.
    """
    alese = standarde_alese(order)
    if not alese:
        return None
    return combina_standarde(alese)


def text_standard(standard):
    """Limitele unui standard, pe linii, pentru tooltip."""
    if standard is None:
        return ""
    componente = standard.get("componente")
    if standard["neaplicabil"]:
        return f"{standard['nume']}: nu au limite chimice (N/A in Excel)."
    if standard["fara_limite"]:
        return f"{standard['nume']}: nu au inca limite completate in Excel."
    if componente:
        linii = [f"{len(componente)} standarde alese (se respecta toate; se ia limita cea mai stricta):"]
        linii += [f"  \u2022 {s['nume']}" for s in componente]
        linii.append("Limite efective:")
    else:
        linii = [standard["nume"]]
    linii += [f"  {el}: {text_limita(lim)}" for el, lim in standard["limite"].items()]
    for el in standard.get("conflicte", []):
        linii.append(f"  \u26a0 {el}: standardele alese se contrazic (minim > maxim).")
    return "\n".join(linii)


# ---------------------------------------------------------------------------
# Verificarea unei compozitii fata de standard
# ---------------------------------------------------------------------------

def verifica_limite(standard, valori):
    """Compara compozitia data cu limitele standardului.

    ``valori`` = {element: procent}, cu elementul in orice scriere
    ("fe", "Fe"). Se verifica DOAR elementele prezente in ``valori`` si care
    au limita in standard (celelalte — ex. C, N, H, care nu se doza — nu au
    ce fi comparate).

    Returneaza lista de incalcari:
      {"element", "valoare", "limita", "tip": "sub"|"peste", "text"}
    """
    if not standard:
        return []
    componente = standard.get("componente")
    if componente:
        # Mai multe standarde alese: fiecare se verifica separat, ca mesajul
        # sa spuna EXACT care standard e incalcat.
        incalcari = []
        for comp in componente:
            incalcari += verifica_limite(comp, valori)
        return incalcari
    dupa_cheie = {_cheie(el): el for el in standard["limite"]}
    incalcari = []
    for element, valoare in valori.items():
        el = dupa_cheie.get(_cheie(element))
        if el is None or valoare is None:
            continue
        lim = standard["limite"][el]
        v = float(valoare)
        lo, hi = lim.get("min"), lim.get("max")
        if hi is not None and v > hi + TOLERANTA:
            incalcari.append({
                "element": el, "valoare": v, "limita": lim, "tip": "peste",
                "text": f"{el} {v:.3f}% depaseste {text_limita({'max': hi})}"
                        f" ({standard['nume']})",
            })
        elif lo is not None and v < lo - TOLERANTA:
            incalcari.append({
                "element": el, "valoare": v, "limita": lim, "tip": "sub",
                "text": f"{el} {v:.3f}% e sub {text_limita({'min': lo})}"
                        f" ({standard['nume']})",
            })
    return incalcari


# ---------------------------------------------------------------------------
# Limitele pentru formularele tiparite
# ---------------------------------------------------------------------------

def spec_efectiv(order):
    """Specificatia aliajului pentru RetDozare / Fisa limita.

    Fara standard ales pe comanda, ramane EXACT specificatia veche din
    config.ALIAJE_SPEC (comenzile existente se tiparesc ca pana acum). Cu
    standard ales, limitele chimice si titlul formularului vin din standard;
    limitele vechi se scot, ca sa nu ramana pe formular o limita pe care
    standardul ales nu o are. Limitele din spec sunt FRACTII (0.002 = 0.20%).
    """
    tip = order.get("tipAliaj")
    baza = dict(ALIAJE_SPEC.get(tip, {}))
    if not baza:                       # placeholder: doar numele aliajului
        baza = {"nume": tip or "", "grad": tip or ""}
    standard = standard_ales(order)
    if standard is None:
        return baza
    for k in [k for k in baza if k.endswith(("_min", "_max"))]:
        del baza[k]
    for element, lim in standard["limite"].items():
        for capat in ("min", "max"):
            if capat in lim:
                baza[f"{element.lower()}_{capat}"] = lim[capat] / 100.0
    baza["titlu_formular"] = standard["nume"]
    return baza
