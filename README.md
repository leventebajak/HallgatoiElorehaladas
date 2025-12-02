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

A projekt a következő forrásfájlokból áll:

*   [`Hallgatoi_Elorehaladas.xlsm`](./Hallgatoi_Elorehaladas.xlsm): A fő Excel fájl (Vezérlőpult)
*   [`MainModule.bas`](./MainModule.bas): A fő vezérlő logika, fájlkezelés és felhasználói interakció
*   [`DataModule.bas`](./DataModule.bas): Adatok beolvasása, validálása és előkészítése
*   [`LogicModule.bas`](./LogicModule.bas): A tantárgyi követelmények (jegyek, aláírások) kiértékelésének logikája
*   [`ReportModule.bas`](./ReportModule.bas): A kimeneti Excel munkalap generálása és formázása