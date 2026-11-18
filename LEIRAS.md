<!--
LEIRAS.md
EZ A FÁJL: rövid, felhasználói leírás a programról (nem kell hozzá programozói tudás).
A részletes telepítési és műszaki leírás a README.md-ben van.
-->

# Minőségügyi dashboard – rövid leírás

## Mire való?

A nemmegfelelőségi (NMF) nyilvántartás Excel-táblájából grafikonokat és rangsorokat készít, hogy gyorsan látszódjon:
- hogyan alakul a hibák száma és költsége hónapról hónapra,
- mely hibatípusok, okok, vevők vagy gépek adják a problémák nagy részét,
- stabil-e a havi hibaarány,
- mely cikkek és munkaszámok kerülnek be újra és újra.

A program **csak a saját gépen fut**, semmilyen adatot nem küld ki az internetre.
Online bemutató (csak kitalált adattal): https://minosegugyi-dashboard.streamlit.app

## Indítás

Dupla kattintás az **`inditas.bat`** fájlra. A böngészőben magától megnyílik a dashboard. Leállítás: a fekete ablak bezárása.

## Adatok

| Mappa / fájl | Mi ez |
|---|---|
| `data/real/` – NMF Excel (.xlsx) | A valós nemmegfelelőségi tábla |
| `data/real/TermelesBeerkezes*.csv` | Késztermék-bevételezés (késztermékraktár) – ebből számolja a havi gyártott darabszámot a p-kártyához. A selejt- és a másodosztály-raktárba került darabok nincsenek benne, ezeket az NMF táblából adja hozzá. Csak a CSV időszakába eső NMF-sorok látszanak. |
| `data/synthetic/` | Kitalált próbaadat, bemutatáshoz és gyakorláshoz |

Alapból a próbaadat látszik. A valós adathoz az oldalsávban kapcsold be a **„Valós adat használata”** kapcsolót.

## Oldalsáv – szűrők

**Hónap** (csúszka: -tól, -ig), **Megrendelő**, **Nemmegfelelőség jellege**, **Keletkezés helye (művelet)**, **NMF keletkezés helye (gép)**, **Következmény**.
Üres szűrő = *Mind*. A **Szűrők törlése** gomb mindent visszaállít.

## Fülek

| Fül | Mit mutat |
|---|---|
| 📈 **Havi trend** | Havi esetszám és költség, trendvonallal. |
| 📊 **Pareto-diagram** | Rangsor választható bontásban (jelleg, ok, vevő, művelet, gép, következmény), darabszám vagy költség szerint, 80%-os vonallal. |
| 🔁 **Ismétlődések** | Többször felvett cikkek és munkaszámok rangsora (esetszám, becsült költség vagy hibás db szerint). Felül az **első 3** tétel és a részesedésük, valamint hogy hány tétel adja a 80%-ot. |
| 🎯 **p-szabályozókártya** | Havi hibaarány (hibás db / gyártott db) szabályozási határokkal. Választható: **Laney p′** (ajánlott) vagy **klasszikus p**, valamint hogy **összes nemmegfelelő** vagy **csak selejt + másodosztály** (végleges veszteség) számítson hibásnak. A hónap-szűrő hat rá, a többi szűrő nem. |
| 📝 **AI-ellenőrzési napló** | Itt rögzíthető, mit ellenőriztél kézzel egy AI által adott eredményen. |

## Beszámoló készítése

Az oldalsáv alján: **📄 Beszámoló készítése**, majd **⬇ PowerPoint** vagy **⬇ PDF**. A beszámoló a most beállított szűrőkkel készül (pl. egy hónapra vagy egy vevőre). Tartalma: kulcsszámok, trend, Pareto-diagramok, az első 3 cikk és munkaszám, ismétlődő cikk–hiba párok, p-kártya és egy üres intézkedési táblázat. A PowerPointba beírhatod a saját értékelésedet.

## Hasznos tippek

- **Hibatípus okainak feltárása:** szűrj egy hibatípusra, majd a Pareto-diagramnál válaszd a *Nemmegfelelőség oka* bontást.
- **Visszatérő problémák:** az Ismétlődések fülön az *„Ugyanaz a cikk, ugyanazzal a hibával”* táblázat mutatja azokat, ahol a korábbi intézkedés valószínűleg nem volt hatásos.
- **Klasszikus p- vagy Laney p′-kártya:** ha a klasszikus kártyán szinte minden hónap kilóg, az a nagy havi darabszám miatt van. Ilyenkor a Laney p′ a megbízhatóbb jelző.
- **Hibás dátum:** ha egy sor dátuma hiányzik vagy hibás, a dashboard tetején sárga sáv írja ki az Excel-sorszámát.
- **Excel módosítása:** mentés után elég a böngészőt frissíteni (F5), a program az új adatot olvassa be.

## Fontos

- A valós adatok (`data/real/`) és az AI-napló nem kerülhetnek fel sehova. Ezt a `.gitignore` is biztosítja.
- A p-kártya hibaaránya közelítő érték. A bevételezés csak a raktárba került készterméket tartalmazza, az „Nmf db” viszont belső selejtet is.
