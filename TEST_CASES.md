# Tesztesetek a Hibakezelés Ellenőrzéséhez

Ez a dokumentum azokat a teszteseteket tartalmazza, amelyekkel ellenőrizhető, hogy a program megfelelően ismeri-e fel és kezeli-e a hibás bemeneteket vagy hiányzó adatokat.

## 1. Bemeneti Validáció (Vezérlőpult)

### 1.1 Nincs forrásfájl kiválasztva
*   **Művelet**: Hagyja üresen a "Hallgatói adatok fájl" mezőt a Vezérlőpulton, majd kattintson a "Kimutatás készítése" gombra.
*   **Elvárt Eredmény**: Hibaüzenet: *"Kérem válasszon ki egy érvényes forrásfájlt!"*

### 1.2 Érvénytelen forrásfájl útvonal
*   **Művelet**: Írjon be kézzel egy nem létező fájl útvonalat (pl. `C:\NemLetezoFajl.xlsx`) a sárga mezőbe, majd kattintson a "Kimutatás készítése" gombra.
*   **Elvárt Eredmény**: Hibaüzenet: *"Kérem válasszon ki egy érvényes forrásfájlt!"*

### 1.3 Üres kurzuslista (Nincsenek sorok)
*   **Művelet**: Törölje ki az összes sort a "KurzusLista" táblázatból (jobb klikk a soron -> Delete -> Table Rows), majd kattintson a "Kimutatás készítése" gombra.
*   **Elvárt Eredmény**: Hibaüzenet: *"Nincsenek megadva kurzusok!"*

### 1.4 Kurzuslista üres sorokkal
*   **Művelet**: Adjon hozzá sorokat a "KurzusLista" táblázathoz, de hagyja üresen a "Tárgykód" oszlopot. Kattintson a "Kimutatás készítése" gombra.
*   **Elvárt Eredmény**: Hibaüzenet: *"Nincsenek megadva kurzusok!"*

## 2. Adatstruktúra Validáció (Forrásfájl)

A teszteléshez készítsen másolatot egy működő hallgatói adatfájlról, és módosítsa az oszlopfejléceket az első sorban.

### 2.1 Hiányzó "Neptun kód" oszlop
*   **Művelet**: Nevezze át vagy törölje a "Neptun kód" oszlopot a forrásfájlban. Futtassa a kimutatást ezzel a fájllal.
*   **Elvárt Eredmény**: Kritikus hibaüzenet: *"A következő oszlopok hiányoznak a fájlból: Neptun kód"*

### 2.2 Több hiányzó oszlop
*   **Művelet**: Törölje vagy nevezze át a "Tárgykód" és "Érvényes" oszlopokat a forrásfájlban.
*   **Elvárt Eredmény**: Kritikus hibaüzenet: *"A következő oszlopok hiányoznak a fájlból: Tárgykód, Érvényes"*

### 2.3 Hiányzó "Bejegyzés típusa" oszlop
*   **Művelet**: Törölje vagy nevezze át a "Bejegyzés típusa" oszlopot.
*   **Elvárt Eredmény**: Kritikus hibaüzenet: *"A következő oszlopok hiányoznak a fájlból: Bejegyzés típusa"*

## 3. Futásidejű Hibák

### 3.1 Érvénytelen fájlformátum
*   **Művelet**: Hozzon létre egy üres szöveges fájlt, és nevezze át `.xlsx` kiterjesztésre. Tallózza be ezt a fájlt és futtassa a kimutatást.
*   **Elvárt Eredmény**: A programnak el kell kapnia a fájl megnyitásakor keletkező hibát. Hibaüzenet: *"Hiba történt: ..."* (pl. fájlformátum hiba).

### 3.2 Sérült vagy zárolt fájl
*   **Művelet**: Próbáljon meg egy olyan fájlt megnyitni, amihez nincs joga, vagy ami sérült.
*   **Elvárt Eredmény**: Általános hibaüzenet a `MainModule` hibakezelőjétől: *"Hiba történt: ..."*
