# Dozare Lingouri Titan

Aplicatie desktop (PySide6) pentru gestionarea stocului de loturi pe
materiale si calculul retetelor de sarja.

## Rulare

```
python3 main.py
```

## Structura proiectului

```
main.py                             punct de intrare (python3 main.py)
dozare_titan/
    config.py                       materiale, tipuri de aliaj, cai fisiere
    utils.py                        helpere mici (to_float, fmt, uid)
    stiluri.py                      culori si stiluri Qt
    calcule/
        comun.py                    functii comune tuturor aliajelor
        ti6al4v.py                  FORMULA de dozare Ti6Al4V (gata facuta)
        vt9.py                      FORMULA de dozare Ti-VT9 (DE IMPLEMENTAT)
        __init__.py                 calculeaza_bara() -- alege formula dupa tipAliaj
    persistenta.py                  incarcare/salvare date.json, migratii
    export/
        fisa_limita.py              genereaza "Fisa limita -cda" (.xlsx + .pdf)
        retdozare.py                 genereaza "RetDozare" (.xlsx + .pdf)
    dialoguri.py                    ferestre de dialog (login, lot nou, istoric lot)
    fereastra_principala.py         fereastra principala (UI)
```

Fiecare fisier are un singur rol clar; nu mai trebuie sa cauti prin 1500
de linii ca sa gasesti o functie — te uiti la numele fisierului.

## Cum adaugi un aliaj nou (ex. Ti-VT9)

Fisierele sunt organizate exact in ordinea in care trebuie atinse:

### 1. `dozare_titan/config.py`
`ALIAJE_SPEC` are deja o intrare pentru `"Ti -VT9"` cu limitele chimice
(O, Si, Fe, Zr, Mo, Al). Daca aliajul tau are alte materiale de baza
decat cele 5 existente (burete, aliaj Al-V, Al metal, Fe metal, TiO2),
adauga-le in `MATERIALE` aici.

### 2. `dozare_titan/dialoguri.py`
Clasa `DialogLotNou`, variabila `ETICHETE_DOZA` — aici adaugi campurile
de compozitie ale unui lot pentru elementele noi (ex. `dozaSi`, `dozaZr`,
`dozaMo`), daca sunt necesare.

### 3. `dozare_titan/fereastra_principala.py`
Metoda `_construieste_card_reteta` — randul `rand_tinta` contine campurile
de compozitie tinta a retetei (`al`, `v`, `o`, `fe`). Adauga aici campurile
noi (`si`, `zr`, `mo`...) daca formula ta le foloseste.

### 4. `dozare_titan/calcule/vt9.py`  <-- AICI SCRII FORMULA
Fisierul contine deja un stub cu explicatii detaliate si aceeasi
semnatura ca `ti6al4v.py`:

```python
def calculeaza_dozare_kg(target, comps, p):
    ...
    return rezultat, eroare
```

unde:
- `target` = compozitia chimica tinta a retetei (dict)
- `comps`  = compozitia (%) fiecarui lot selectat, pe material
- `p`      = portia, in kg
- `rezultat` = dict `{material_id: kg}` pentru o singura portie (sau `None`)
- `eroare`   = mesaj de eroare (string), sau `None` daca a mers bine

Uita-te la `dozare_titan/calcule/ti6al4v.py` ca model — e formula
existenta pentru Ti6Al4V, cu acelasi tipar (balanta element cu element,
in ordine).

### 5. `dozare_titan/calcule/__init__.py`
Dictionarul `DOZATOARE` mapeaza tipul de aliaj -> functia lui de calcul:

```python
DOZATOARE = {
    "Ti6Al4V": ti6al4v.calculeaza_dozare_kg,
    "Ti -VT9": vt9.calculeaza_dozare_kg,
}
```

Aceasta intrare exista deja, deci in momentul in care completezi
`calculeaza_dozare_kg` din `vt9.py`, formula ta e automat folosita peste
tot in aplicatie (bilant bara, bilant reteta, RetDozare) — nu mai trebuie
sa umbli in `fereastra_principala.py` sau `export/`.

### 6. (optional) `dozare_titan/export/retdozare.py`
Tabelele din RetDozare afiseaza in acest moment coloanele Al/V/O/Fe,
specifice Ti6Al4V. Daca vrei ca raportul RetDozare pentru Ti-VT9 sa
afiseze Si/Zr/Mo in loc de Al/V, va trebui sa adaptezi coloanele din
`_scrie_bloc_retdozare_xlsx` si `genereaza_retdozare_pdf`. Fisa limita
(`export/fisa_limita.py`) nu are nevoie de nicio modificare — lucreaza
direct cu cantitati pe material, indiferent de aliaj.

## Note

- `calcule/comun.py` contine functiile folosite de toate aliajele
  (`lot_ti`, `lot_rest`, `target_ti`, `material_necesar`,
  `_rezumat_calcul`). `lot_ti`/`target_ti` presupun in acest moment ca
  Ti e "restul" din Al+V+O+Fe+N — daca formula ta pentru VT9 calculeaza
  Ti altfel, poti scrie o varianta proprie direct in `vt9.py` fara sa
  atingi `comun.py`.
- Datele salvate (`~/.dozare_titan/date.json`) raman neschimbate ca
  format — acest refactor e doar organizare de cod, nu schimba nimic
  din datele existente sau din comportamentul aplicatiei.
