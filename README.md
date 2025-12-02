# Hallgatói Előrehaladás Kimutatás

Ez a projekt egy Excel alapú eszközt biztosít a hallgatói előrehaladás elemzésére. A megoldás VBA makrókat használ a nagy mennyiségű hallgatói adat gyors feldolgozására és formázott kimutatások generálására.

## Használat

1.  Nyissa meg a [`Hallgatoi_Elorehaladas.xlsm`](./Hallgatoi_Elorehaladas.xlsm) fájlt

2.  **Adatforrás kiválasztása**:
    *   A **"Vezérlőpult"** munkalapon kattintson a **"Fájl kiválasztása"** gombra
    *   Tallózza be a hallgatói adatokat tartalmazó Excel fájlt
    *   A fájl útvonala megjelenik a sárga mezőben

3.  **Kurzusok megadása**:
    *   A táblázatban adja meg a vizsgálandó tárgyakat
    *   **Tárgykód**: A tantárgy Neptun kódja (pl. `BME...`)
    *   **Bejegyzés típusa**: Válasszon a legördülő listából:
        *   *Évközi jegy*: Csak egy jegyet keres
        *   *Aláírás és Vizsgajegy*: Külön keresi az aláírást és a vizsgát
        *   *Aláírás*: Csak aláírást keres
        *   *Szigorlat*: Hasonló az évközi jegyhez

4.  **Futtatás**:
    *   Kattintson a **"Kimutatás készítése"** gombra
    *   A program ellenőrzi az adatokat, és létrehoz egy új munkalapot (pl. 14:30-kor `1430_...`)

## Működési Logika és Funkciók

### Kiértékelési Szabályok
*   **Évközi jegy / Szigorlat**:
    *   A legutolsó **érvényes** bejegyzést keresi.
    *   Zöld háttér, ha a bejegyzés "Elismert".
*   **Aláírás és Vizsgajegy**:
    *   **Aláírás**: A legutolsó "Aláírás" bejegyzést keresi. **Fontos**: Itt *nem* vizsgálja az "Érvényes" oszlop értékét (mivel a Neptunban egy vizsga érvénytelenítheti az aláírást, de az aláírás ténye megmarad).
    *   **Vizsgajegy**: Ha van aláírás (és nem "Megtagadva"), keresi a legutolsó **érvényes** vizsgajegyet.
    *   Zöld háttér, ha mindkettő "Elismert".

### Formázás
*   **Színezés**:
    *   A tárgyak oszlopai váltakozó kék árnyalatúak az átláthatóság érdekében.
    *   Hiányzó tárgyteljesítés esetén a cella sárga.
    *   Sikeres (elismert) teljesítés esetén a cella zöld.

### Adatfeldolgozás
*   **Teljesítmény**: A program `Scripting.Dictionary` objektumot használ az adatok memóriában történő gyors kereséséhez.
*   **Intelligens Fájlkezelés**: Ha a forrásfájl már meg van nyitva az Excelben, a program azt használja (nem nyitja meg újra), és futás után nyitva is hagyja. Ha nincs nyitva, megnyitja "Csak olvasható" módban, majd be is zárja.


### Hibakezelés
A program figyelmeztet, ha:
*   Nincs kiválasztva forrásfájl.
*   A forrásfájlból hiányoznak kötelező oszlopok (pl. Neptun kód, Tárgykód).
*   A kurzuslista üres.
*   Egy kurzusnál nincs megadva a "Bejegyzés típusa" (alapértelmezetten Évközi jegyként kezeli).

## Fájlok

<<<<<<< HEAD
A projekt a következő forrásfájlokból áll:

*   [`Hallgatoi_Elorehaladas.xlsm`](./Hallgatoi_Elorehaladas.xlsm): A fő Excel fájl (Vezérlőpult)
*   [`MainModule.bas`](./MainModule.bas): A fő vezérlő logika, fájlkezelés és felhasználói interakció
*   [`DataModule.bas`](./DataModule.bas): Adatok beolvasása, validálása és előkészítése
*   [`LogicModule.bas`](./LogicModule.bas): A tantárgyi követelmények (jegyek, aláírások) kiértékelésének logikája
*   [`ReportModule.bas`](./ReportModule.bas): A kimeneti Excel munkalap generálása és formázása
=======
1. **Hallgatói adatok betöltése:**
   - Kattints a "Tallózás..." gombra
   - Válaszd ki a hallgatói adatokat tartalmazó Excel fájlt
   - A fájlnak tartalmaznia kell a következő oszlopokat:
     - Modulkód
     - Felvétel féléve
     - Neptun kód
     - Nyomtatási név
     - Tárgykód
     - Tárgynév
     - Bejegyzés értéke
     - Bejegyzés típusa
     - Bejegyzés dátuma
     - Érvényes
     - Elismert

2. **Excel fájl létrehozása:**
   - Kattints az "Excel fájl létrehozása" gombra
   - Add meg a mentési helyet és fájlnevet
   - Az alkalmazás automatikusan feldolgozza az adatokat

**Kimeneti fájl szerkezete:**

Az Excel fájl az alábbi oszlopokat tartalmazza:
- **Alapadatok** (minden hallgatónál):
  - Modulkód
  - Felvétel féléve
  - Neptun kód
  - Nyomtatási név
  - Felvételi összes pontszám (üres)
  - Státusz (üres)

- **Kurzus oszlopok** (a hozzáadott kurzusok sorrendjében):
  - Évközi jegy / Szigorlat / Aláírás: 1 oszlop
  - Aláírás + Vizsgajegy: 2 oszlop (aláírás | vizsgajegy)

**Színkódolás:**

Az Excel fájlban a cellák automatikusan színkódolva vannak az alábbiak szerint:

- **Zöld (#92D050)**: A kurzus elismert ("Elismert" oszlop értéke "Igaz")
- **Sárga (#FFFF00)**: A hallgató nem vette fel a kurzust (üres cella)
- **Váltakozó kék árnyalatok**: Nem elismert, de felvett kurzusok

**Adatfeldolgozás logikája:**

- **Évközi jegy / Szigorlat:**
  - A legfrissebb érvényes bejegyzés kerül be

- **Aláírás:**
  - A legfrissebb aláírás bejegyzés kerül be

- **Aláírás + Vizsgajegy:**
  - Bal oszlop: Legfrissebb aláírás (Aláírva/Megtagadva)
  - Jobb oszlop: Ha aláírás = "Aláírva", akkor a legfrissebb érvényes vizsgajegy
  - Ha aláírás = "Megtagadva", a jobb oszlop üres marad

## Naplózás

Az alkalmazás automatikusan naplóz minden műveletet az Excel generálás során:

- **Ha nincsenek hibák/figyelmeztetések:** A naplófájl automatikusan törlődik
- **Ha vannak hibák/figyelmeztetések:** A naplófájl megmarad és a név a kimeneti fájl neve + `_log_ÉÉÉÉHHNN_ÓÓPPMM.txt`

A naplófájl tartalmazza:
- Feldolgozott adatok mennyisége
- Minden egyes kurzus feldolgozásának részletei
- Hiányzó vagy érvénytelen adatok figyelmeztetései
- Hibák részletes leírása

## Fájlformátumok

### Kurzuslista JSON formátum

```json
[
  {
    "course_code": "BMETE90AX21",
    "grading_type": "Aláírás és Vizsgajegy"
  },
  {
    "course_code": "BMETE11AX52",
    "grading_type": "Évközi jegy"
  },
  {
    "course_code": "BMETE90AX20",
    "grading_type": "Szigorlat"
  },
  {
    "course_code": "BMEVIDHKONI",
    "grading_type": "Aláírás"
  }
]
```

## Hibaelhárítás

### "Érvénytelen fájl" hiba
- Ellenőrizd, hogy a hallgatói Excel fájl tartalmazza az összes szükséges oszlopot
- Az oszlopneveknek pontosan egyezniük kell a fent felsoroltakkal

### Üres cellák a kimeneti fájlban
- Ellenőrizd a naplófájlt, ha létrejött
- Lehet, hogy az adott hallgatónak nincs érvényes bejegyzése az adott kurzushoz
- Ellenőrizd, hogy a tárgykód pontosan egyezik a forrás fájlban és a kurzuslistában

### Az alkalmazás nem indul
- Ellenőrizd, hogy telepítve vannak-e a szükséges csomagok: `pip install -r requirements.txt`
- Ellenőrizd a Python verziót: `python --version` (minimum 3.8 szükséges)

## Technikai részletek

- **Programozási nyelv:** Python 3
- **GUI keretrendszer:** tkinter
- **Fő könyvtárak:**
  - pandas: Adatfeldolgozás
  - openpyxl: Excel fájl írás/olvasás
  - logging: Naplózás

## Licensz

Ez a projekt oktatási célra készült.
>>>>>>> origin/master
