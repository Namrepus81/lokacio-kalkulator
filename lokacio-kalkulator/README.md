# Lokáció Kalkulátor

Egyszerű, SAP-tól független, böngészőben futó alapverzió.

## Indítás

Dupla kattintás:

`D:\Web\Projektek\lokacio-kalkulator\start.bat`

Vagy nyisd meg ezt a fájlt böngészőben:

`D:\Web\Projektek\lokacio-kalkulator\index.html`

Excel sablon:

`D:\Web\Projektek\lokacio-kalkulator\lokacio-kalkulator-sablon.xlsx`

## Excel oszlopok

Az első sor legyen fejléc. Ezeket az oszlopokat várja:

- `cikkszám`
- `megnevezés`
- `lokáció`
- `min`
- `max`
- `forgás`
- `tároló`
- `súly_kategória`
- `kapacitás`
- `megjegyzés`

Kötelező mezők:

- `cikkszám`
- `lokáció`
- `min`
- `max`
- `forgás`
- `tároló`
- `súly_kategória`
- `kapacitás`

Ékezet nélküli fejléceket is elfogad, például `cikkszam`, `lokacio`, `tarolo`, de a javasolt forma az ékezetes. A felületen a betöltött értékeket ékezetesen írja ki.

## Érték minták

`forgás`:

- `sűrű`
- `közepes`
- `lassú`

`tároló`:

- `kis KLT`
- `nagy KLT`
- `XL láda`
- `raklap`

`súly_kategória`:

- `könnyű`
- `közepes`
- `nehéz`

`lokáció`:

- forma: `A-A-1-1`, `A-A-1-2`, `C-A-1-1`, `B-A-1-1`
- az utolsó szám a szint
- `A-A-1-1` = 1. szint
- `A-A-1-2` = 2. szint
- `A-A-1-3` = 3. szint
- `A-A-1-4` = 4. szint, ez a maximum

`kapacitás`:

- szám, például `2000`
- azt jelenti, hogy az adott lokációban/szinten mennyi fér el abból a cikkszámból

## Jelenlegi szabályok

- A/B lokáció: elöl
- C/D lokáció: közép
- E/F/G/H lokáció: hátul
- sűrű anyag hátul: figyelmeztetés
- lassú anyag elöl: figyelmeztetés
- nehéz anyag 2. vagy magasabb szinten: kritikus
- nagy vagy XL láda 2. vagy magasabb szinten: figyelmeztetés
- raklap 2. vagy magasabb szinten: kritikus
- azonos cikkszám lokációinál csak az 1-2. szint kapacitását adja össze kalkulációhoz
- kalkulációba csak ezek a lokációkezdetek számítanak: `A-A`, `A-B`, `A-C`, `A-D`
- `A-E`, `B-A`, `C-A`, `C-B` és minden más lokáció csak tájékoztató
- a 3-4. szint kapacitása csak tájékoztató jellegű
- ha az alsó két szint kapacitása kisebb, mint a max készlet: figyelmeztetés
- ha az alsó két szint kapacitása kisebb, mint a min készlet: kritikus

## Logikus elhelyezés

A `Logikus elhelyezés` rész cikkszámonként pontozza a meglévő lokációkat.

Figyelembe veszi:

- forgás: sűrű / közepes / lassú
- zóna: elöl / közép / hátul
- szint: a lokáció utolsó száma alapján
- tároló: kis KLT / nagy KLT / XL láda / raklap
- súly_kategória: könnyű / közepes / nehéz
- min / max készlet
- alsó két szint kapacitása
- csak `A-A`, `A-B`, `A-C`, `A-D` lokációkat javasol
- a többi lokáció és a 3-4. szint csak információként jelenik meg, nem számít bele a max kalkulációba

Ez az első verzió még nem talál ki teljesen új üres lokációkat. A betöltött Excelben szereplő lokációkat rendezi és értékeli logikus sorrendbe.

## SAP export jellegű fájl

Az app felismeri ezeket az oszlopokat is:

- `Raktárhely` -> lokáció
- `Termék` -> cikkszám
- `Termék rövid leírása` -> megnevezés
- `Mennyiség` -> aktuális mennyiség

Ez önmagában keresésre jó, de a teljes kalkulációhoz továbbra is kell a min, max, forgás, tároló, súly_kategória és kapacitás adat.

## Két Exceles működés

Az app két külön fájlt tud kezelni:

1. `Napi export`
   - SAP/export jellegű aktuális lista
   - ebből jön a cikkszám, lokáció, megnevezés, mennyiség

2. `Törzsadat`
   - ritkán módosított adatbázis
   - ebből jön a min, max, forgás, tároló, súly_kategória, kapacitás

A program cikkszám alapján fűzi össze a két fájlt. Ha a napi exportban új cikkszám van, de a törzsadatban nincs meg hozzá minden szükséges adat, a `Hiányzó adatok` részben listázza, hogy mit kell még felvinni.

## Megjegyzés

Az `.xlsx` olvasáshoz a SheetJS könyvtárat CDN-ről tölti be a böngésző. Ha nincs internet, a CSV feltöltés akkor is működik.
