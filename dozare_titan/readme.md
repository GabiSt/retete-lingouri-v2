# Dozare Lingouri Titan

Aplicatie desktop (PySide6) pentru gestionarea stocului de loturi pe
materiale si calculul retetelor de sarja.

## Rulare

```
python3 main.py
```

## Compilare in .exe (Windows)

```
build_windows.bat
```

sau, manual:

```
pip install pyinstaller
pyinstaller --noconfirm dozare_titan.spec
```

Rezultatul apare in `dist/dozare_titan/dozare_titan.exe`. **Muta/copiaza
tot folderul `dist/dozare_titan`**, nu doar exe-ul — langa el trebuie sa
ramana fisierele lui interne, inclusiv `dozare_titan/assets/logo.png`.

**Foloseste intotdeauna `dozare_titan.spec`, nu `pyinstaller main.py`
direct** — o compilare fara `.spec`/`--add-data` NU include folderul
`assets` in exe, iar documentele generate (Fisa limita, RetDozare) revin
automat la textul de rezerva "ZIROM TITANIUM" in loc de logo (fara nicio
eroare vizibila, deci trece usor neobservat).

## Structura proiectului

```
main.py                             punct de intrare (python3 main.py)
dozare_titan/
    config.py                       materiale, tipuri de aliaj, cai fisiere
    utils.py                        helpere mici (to_float, fmt, uid)
    stiluri.py                      culori si stiluri Qt
    calcule/
        comun.py                    functii comune tuturor aliajelor
        capacitate.py               cate bare incap intr-o reteta + bara de report
        planificare.py              generarea automata a retetelor din loturile
                                    comenzii, distributia barelor, aplicarea si
                                    RETRAGEREA consumului (logica pura, testabila)
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

## Aliaje si standarde (din Excel)

Pe fiecare comanda alegi **aliajul** si, langa el, **standardele** (se pot
alege **mai multe**, ex. AMS 4975 + AMS 4976 — vezi mai jos). Datele vin
din `standarde.xlsx` (un rand pe standard, o coloana pe element chimic; valori
de forma `5.50-6.75`, `max. 0.30`, `-` = fara limita, `N/A`).

- **Unde e fisierul:** aplicatia il cauta intai langa exe / `main.py`
  (`standarde.xlsx` — acesta e cel pe care il editezi), apoi in
  `dozare_titan/assets/` (copia livrata). Se reciteste singur cand il
  salvezi; nu trebuie repornita aplicatia. Coloane noi de elemente se pot
  adauga fara modificari de cod.
- **Mai multe standarde pe o comanda:** butonul din dreapta aliajului deschide
  o lista cu casute de bifat (ramane deschisa cat bifezi). Materialul trebuie
  sa respecte **toate** standardele alese, deci:
  - avertismentele sunt verificate pe fiecare standard in parte, iar mesajul
    numeste exact standardul incalcat (daca o valoare iese la doua standarde,
    apar doua mesaje);
  - pe formularele tiparite (RetDozare, Fisa limita) limita fiecarui element e
    **cea mai stricta** dintre standardele alese (minimul cel mai mare, maximul
    cel mai mic), iar titlul contine numele lor separate prin „;”. Un standard
    N/A nu restrange nimic;
  - daca doua standarde se contrazic la un element (minimul unuia e peste
    maximul celuilalt), tooltip-ul o spune;
  - la schimbarea aliajului raman bifate doar standardele care exista si la
    noul aliaj.
  Pe comanda se salveaza `standarde` (lista); `standard` (primul nume) ramane
  sincronizat pentru datele si codul vechi, iar comenzile salvate cu un singur
  standard sunt migrate automat.
- **Avertismente pe reteta:** compozitia tinta a retetei (Al, V, O, Fe, Mo,
  Si, Zr) se compara cu limitele standardului ales; ce iese din interval
  apare ca banner galben pe reteta, iar compozitia rezultata din calcul e
  verificata la fel pe fiecare bara. Elementele nedozate (C, N, H, Y, Sn, Co)
  nu se pot verifica. La "Loturi si configurare", o tinta in afara
  standardului cere confirmare la salvare.
- **Formulare tiparite:** cu standard ales, limitele chimice si titlul din
  RetDozare / Fisa limita vin din standard. Fara standard ales, comenzile se
  tiparesc ca pana acum (limitele din `config.ALIAJE_SPEC`).
- **Placeholder-e:** orice aliaj din Excel care nu are reteta de dozare
  apare in lista marcat "(placeholder)": il poti alege pe comanda si ii vezi
  standardele si limitele, dar calculul, distribuirea barelor si RetDozare
  nu sunt disponibile. Ca sa devina aliaj real: adauga-l in
  `config.ALIAJE_SPEC` si in `calcule.DOZATOARE`, sub cheia egala cu numele
  din Excel (specificatii de dozare pregatite exista in
  `calcule/specificatii.py`, dar unele cer materiale care nu sunt inca in
  `config.MATERIALE`, ex. Sn metal).
- **Legatura cu Excel-ul:** `config.ALIAJE_EXCEL` spune ce nume din coloana
  "Aliaj" corespunde aliajelor cu reteta (Ti5 si Ti6Al4V-AMS = "Titan grad 5
  (Ti 6Al 4V)"). Daca redenumesti aliajul in Excel, aplicatia te avertizeaza.
  Celulele neintelese si standardele duplicate cu valori diferite sunt
  raportate tot printr-un mesaj.

## Aliajul Ti 6-2-4-2 (Ti-6Al-2Sn-4Zr-2Mo-0.08Si)

Reteta vine din Excel-ul de productie **Cda 24883-Ti6242** (foaia `RetDozare`)
si e scrisa direct in `calcule/ti6242.py`. Materiale: burete, Aliaj AlMo,
Al metal, **Zircaloy 4 (Zy4)**, **Sn metal**, prealiaj SiTi, TiO₂ (Zy4 si Sn
metal sunt materiale noi in `config.MATERIALE`; lotul are un camp nou
`dozaSn`). Calculul merge in ordinea Mo → Al → Zr → Sn → Si → O → Ti; fiecare
material acopera elementul lui minus ce au adus deja cele dinainte (Al din
AlMo, Sn din Zy4, O din toate).

Doua moduri (`ti6242.MOD_IMPLICIT`):

- `"complet"` (implicit, ca restul aplicatiei): balanta de masa completa, se
  scade intai si se imparte dupa; compozitia iese exact pe tinta.
- `"excel"`: reproduce celulele din foaia de calcul, inclusiv o particularitate
  de precedenta (`=B*Al/puritate - Al_adus`, scaderea ramane neimpartita la
  puritate; TiO₂ nu se scade din cei 10 kg). Diferentele: Al metal si Sn cu
  cateva grame, TiO₂ de ~3,5 ori mai mare la Reteta 1 (oxigenul din bara iese
  cu ~48% peste tinta).

Atentie la **tinta**: in Excel, „Dozare in bara” (Al 6,8%) difera de
„Conc. TINTA” (Al 6,0%) — diferenta compenseaza pierderea la topire. Aplicatia
foloseste o singura tinta (cea de dozare), deci Al 6,8% va aparea ca
depasire fata de max 6,50% din AMS 4975/4976.

Verificare: `python3 teste/test_ti6242.py` (compara cu celulele Excel-ului).

## Cum se lucreaza (modul automat)

Tot ce introduce utilizatorul se afla intr-un singur loc: butonul
**"Loturi si configurare"** de pe comanda. Acolo se pun:

- portia (kg), numarul de presari, numarul **total** de bare al comenzii;
- compozitia tinta in bara;
- pe fiecare material, **lista de loturi in ordinea in care intra in
  comanda** (se adauga din stoc, se muta cu sagetile, se scot cu ✕).

La "Salveaza si genereaza retetele", programul construieste singur toate
retetele: toarna bare din loturile de pe primul rand, iar cand un lot se
termina intra automat in urmatorul lot din lista aceluiasi material si
deschide o reteta noua. **Nu se mai creeaza retete de mana si nu se mai
aleg loturi pe fiecare reteta.**

Planul NU se scrie direct pe comanda: apare mai intai o fereastra de
**validare**, cu fiecare reteta propusa (cate bare, ce tip — report sau
noi — si ce loturi foloseste). Operatorul verifica si apasa "Confirma si
salveaza"; daca apasa "Renunta", comanda ramane exact cum era. Daca in
timpul generarii s-a terminat lista de loturi a unui material, planul
partial apare oricum, cu avertismentul, si un buton "Deschide configurarea
si adauga lotul" duce direct inapoi acolo, unde operatorul adauga lotul nou
si genereaza din nou. Mai ramane un singur lucru de facut dupa confirmare:
aplicarea consumului, reteta cu reteta.

Butonul **"Regenereaza retetele"** reface planul cu stocul de la momentul
apasarii. Retetele care au deja consumul aplicat pe stoc raman neatinse — se
rescrie doar ce urmeaza dupa ele.

Comenzile mai vechi, fara loturi alocate pe comanda, raman pe modul manual
(reteta noua + loturi pe reteta + "Distribuie barele") pana cand se deschide
"Loturi si configurare".

## Cum se impart barele pe retete (bara de report)

Numarul de bare se seteaza **pe comanda** (ex. 12), nu pe reteta. Regula
dupa care se umple fiecare reteta, la fel in modul automat si in cel manual:

1. Reteta 1 primeste atatea bare cate incap in loturile ei.
2. Prima bara care nu mai incape **nu se toarna partial**: trece pe reteta
   urmatoare si isi **pastreaza dozarea veche** (cea calculata pe loturile
   retetei din care pleaca) — reteta de presare era deja stabilita in
   momentul in care s-a constatat ca nu mai ajunge materialul.
   - din loturile vechi mai ia doar **restul ramas din materialul care s-a
     terminat** (lotul vechi ramane pe zero, nu se pierde nimic);
   - **diferenta** la acel material, plus **doza intreaga la toate celelalte
     materiale**, se consuma din loturile retetei noi.
3. Restul barelor retetei noi se dozeaza normal, pe loturile noi.
4. Mecanismul se repeta de la reteta la reteta.

Exemplu real (Cda 26858-Ti5, 12 bare de 900 kg), verificat automat:

```
Reteta 1: barele 1-4          se termina prealiajul Al-V (lot 360744) la bara 5
Reteta 2: bara 5 (report, dozarea retetei 1) + barele 6-10
                              se termina buretele (lot 260215-439) la bara 11
Reteta 3: bara 11 (report, dozarea retetei 2) + bara 12
```

Consumul se aplica **o singura data pe reteta**, din "Capacitate reteta /
desfasurare pe bare". Fiecare reteta are si butonul **"Retrage consumul
retetei"**, care anuleaza tot ce a scazut reteta din stoc — inclusiv partea
luata de bara cedata retetei urmatoare, care redevine bara in asteptare.
Retragerea se face in ordine inversa: daca bara cedata a fost deja
consumata pe reteta urmatoare, aplicatia cere sa retragi intai acolo.

## Verificare automata

```
python3 teste/test_cda_26858.py          # distributia, consumul, retragerea
python3 teste/test_generare_automata.py  # generarea automata din loturile comenzii
python3 teste/test_standarde.py          # standarde din Excel, placeholder-e, limite
python3 teste/test_standarde_multiple.py # selectia multipla de standarde
python3 teste/test_ti6242.py             # reteta Ti 6-2-4-2 vs Excel Cda 24883
QT_QPA_PLATFORM=offscreen python3 teste/test_selector_standarde.py   # controlul de bifare
```

Ruleaza fara interfata grafica si compara, pas cu pas, dozarile, distributia
barelor, consumul si restul fiecarui lot cu Excel-ul de productie
"Cda 26858-Ti5-2VAR-d600-L1-L4" (tab-urile "DateIntrare" si "Retete"),
plus simetria retragerii consumului si concordanta cu Fisa limita.

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
