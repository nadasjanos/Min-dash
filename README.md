<!--
README.md
EZ A FÁJL: a projekt leírása – mit csinál, hogyan kell telepíteni és futtatni.
-->

# Minőségügyi dashboard (Streamlit)

Egyszerű, a saját gépeden futó webes felület a nemmegfelelőségek elemzéséhez.

> 🌐 **Kipróbálás telepítés nélkül (kitalált adatokkal):** https://minosegugyi-dashboard.streamlit.app
> A saját adataiddal töltsd le a programot, és a saját gépeden futtasd – lásd [Kipróbálás a saját adataiddal](#kipróbálás-a-saját-adataiddal).
>
> 🇬🇧 **English:** the dashboard and the PowerPoint/PDF report are also available in English – switch in the sidebar, or open the demo directly in English: https://minosegugyi-dashboard.streamlit.app/?lang=en

**Nyelv:** az oldalsáv tetején választható **magyar** vagy **angol** felület. A beszámoló (PowerPoint / PDF) is a választott nyelven készül. A próbaadat kategóriái és a státuszok angolra fordítva látszanak; a valós Excel tartalmát (vevőnevek, hibaleírások) a program nem fordítja. Nyelvváltáskor a szűrők törlődnek.

## Mit csinál?

Négy fül (nézet) van benne:

| Fül | Mit mutat |
|---|---|
| 📈 **Havi trend** | Havonta hány nemmegfelelőség volt és mennyi költséget okozott, trendvonallal. |
| 📊 **Pareto-diagram** | Választható bontásban (jelleg, ok, vevő, művelet, gép, következmény) rangsorolva, darabszám vagy költség szerint, kumulált %-kal és 80%-os vonallal. 15-nél több kategóriánál a többit „Egyéb” oszlopba vonja. Tipp: szűrj egy hibatípusra, és bonts ok szerint. |
| 🔁 **Ismétlődések** | Választható rangsor: *esetek száma*, *becsült költség* vagy *hibás db*. Felül **🏆 Első 3** összefoglaló cikkre és munkaszámra: az első 3 tétel, a részesedésük a teljes (szűrt) értékből, és hogy hány tétel adja ténylegesen a 80%-ot (ebben az egyszer szereplő tételek is számítanak). Alatta a többször felvett tételek rangsorolva: grafikon a 15 legtöbbször hibás cikkről hibatípusonként színezve, valamint táblázatok: *ugyanaz a cikk ugyanazzal a hibával* (pl. egy cikk többször kerülethibával), *ismétlődő cikkek* és *ismétlődő munkaszámok* (pl. egy munkaszám több hibával). Csúszkával állítható, legalább hányszor szerepeljen. |
| 🎯 **p-szabályozókártya** | A havi hibaarány (hibás db / gyártott db) a szabályozási határokkal. A határon kívüli hónapokat piros X jelöli. Kapcsolóval választható: **Laney p′** (alapértelmezett) vagy **klasszikus p** (p̄ ± 3σ). Valós adatnál azt is választhatod, mit számoljon hibásnak: **összes nemmegfelelő** (első menetes hibaarány, a javíthatókkal együtt) vagy **csak selejt + másodosztály** (végleges veszteség) – a gyártott db mindkettőnél ugyanaz. Nagy darabszámnál (több ezer db/hó) a klasszikus kártya határai nagyon szűkek, ezért szinte minden hónap kilóg; a Laney p′ a hónapok közötti valódi ingadozáshoz (σz) igazítja a határokat, így csak a tényleg kiugró hónapot jelzi. |
| 📝 **AI-ellenőrzési napló** | Űrlap, amivel rögzítheted, mit ellenőriztél kézzel egy AI által adott kimeneten. Helyben, a `data/ai_naplo.csv` fájlba ment. |

Az oldalsávban (*Beállítások*) van az adatforrás-kapcsoló és a **szűrők**:

- Hónap – csúszka, a két végével adod meg a „-tól” és „-ig” hónapot
- Megrendelő (Vevő) – szintetikus adatnál: termékcsalád
- Nemmegfelelőség jellege – szintetikus adatnál: hibatípus
- Keletkezés helye (művelet), NMF keletkezés helye (gép), Következmény – csak valós adatnál

Minden szűrő üresen indul, ez azt jelenti: **Mind**. Ha kiválasztasz egy vagy több értéket, csak azok látszanak. A *Szűrők törlése* gomb mindent visszaállít. A szűrők a trendre, a Pareto-diagramra és a kulcsszámokra hatnak. A p-kártyára csak a hónap-szűrő hat (a középvonal és a határok a kiválasztott hónapokból számolódnak), a többi szűrő nem, mert a gyártott mennyiség csak havi összesítésben ismert.

### Beszámoló (PowerPoint + PDF)

Az oldalsáv alján a **📄 Beszámoló készítése** gombbal egy prezentálható beszámoló készül a **most beállított szűrőkkel** (és a p-kártyánál kiválasztott típussal). Utána két letöltés gomb jelenik meg: **PowerPoint (.pptx)** – szerkeszthető, értekezletre – és **PDF** – archiváláshoz, e-mailhez.

Tartalom (9–10 dia / oldal): címlap (időszak, forrás, szűrők) · kulcsszámok · havi trend · Pareto jelleg szerint · Pareto ok szerint · első 3 cikk és munkaszám becsült költség szerint · ismétlődő cikk–hiba párok (top 10) · p-szabályozókártya (összes nemmegfelelő) · valós adatnál p-kártya a végleges veszteségre (selejt + másodosztály) · üres intézkedési táblázat (kézzel kitöltendő). A grafikonok alatt automatikus, tényszerű megállapítások vannak; a saját értékelésedet a PowerPointba írhatod be.

Ha a szűrők a gombnyomás után változnak, az app jelzi, hogy a beszámolót újra kell készíteni. A fájl a saját gépeden készül, semmi nem megy ki az internetre – de a valós adatos beszámolóban vevőnevek és költségek vannak, ezért megosztás előtt nézd át, kinek küldöd.

## Adatok és adatvédelem

- **Alapból a szintetikus (kitalált) adatot** használja: `data/synthetic/`.
- **Valós adat**: a `data/real/` mappába másold az Excel-fájlt. Ezt az app **csak akkor** olvassa, ha az oldalsávban bekapcsolod a *„Valós adat használata”* kapcsolót.
- A `data/real/` mappa és az AI-napló benne van a `.gitignore`-ban, így Gittel sem kerülhet fel sehova.
- A program semmit nem tölt fel az internetre. A Streamlit használati statisztika-küldése ki van kapcsolva (`.streamlit/config.toml`), és az `inditas.bat` úgy indítja, hogy csak a saját gépedről legyen elérhető (`localhost`).

### Valós Excel oszlopai

A `valos_adat.py` fájl az Excel oszlopait a belső formátumra alakítja:

| Belső mező | Excel-oszlop (alapbeállítás) |
|---|---|
| dátum | `Dátum` |
| hibatípus | `Nem megfelelőség jellege` |
| ok | `Nem megfelelőség oka` |
| termékcsalád | `Vevő` (nincs külön termékcsalád oszlop – átírható pl. `Cikkszám`-ra) |
| költség | `Becsült költségvonzat Ft` |
| státusz | `n=nincs teendő f=folyamtban, …` oszlop kódja: n → *Lezárt*, f → *Folyamatban*, o / b / t / i → *Nyitott – oktatás / beszállító / technológia / intézkedés szükséges*, üres → *Nincs megadva* |
| hibás darab | `Nmf db` |
| keletkezés helye *(nem kötelező)* | `Keletkezés helye (Művelet)` |
| gép *(nem kötelező)* | `NMF keletkezés helye (gép)` |
| következmény *(nem kötelező)* | `Következmény` |
| munkaszám *(nem kötelező)* | `Munkaszám` |
| cikkszám *(nem kötelező)* | `Cikkszám` |

A szöveges mezőkben (vevő, jelleg, ok, művelet, gép, következmény) a csak írásmódban eltérő értékeket – szóköz, kis/nagybetű, ékezet, írásjel – az app egy névre vonja össze, a leggyakoribb írásmóddal (pl. „Belsőgyártás” → „Belső gyártás”). Az Excelt ez nem módosítja.

Ha valamelyik oszlop máshogy szerepel, írd át a `valos_adat.py` tetején lévő `OSZLOPOK` szótárban. Az app hibaüzenetben jelzi, ha nem talál egy oszlopot.

**p-kártya valós adatra:** ehhez kell a havi összes gyártott darabszám is. Az app ezt automatikusan a késztermék-bevételezés exportjából számolja, ha van `data/real/TermelesBeerkezes*.csv` fájl (pontosvesszős CSV; a „Bevét dátum” szerint havonta összeadja a „Mennyiség” oszlopot). A CSV a **késztermékraktár** bevételezése: a jó termékeket és a selejtpótló munkaszámokat tartalmazza. A selejt- és a másodosztály-raktárba érkeztetett darabok nincsenek benne – ezért az app hozzáadja az NMF táblából a `Következmény` = *Teljes selejt* vagy *Másodosztály* sorok `Nmf db` értékét. A javított darabok a javítás után a bevételezésben szerepelnek, őket nem adja hozzá még egyszer. (A listát a `valos_adat.py` `NEM_BEERKEZETT_KOVETKEZMENYEK` sora állítja.) **Közös időszak:** ha van bevételezési CSV, az app csak azokat az NMF-sorokat mutatja, amelyek dátuma a CSV első és utolsó bevételezési napja közé esik, hogy a hibás és a gyártott darab ugyanarra az időszakra vonatkozzon. Az oldal tetején látszik a vizsgált időszak, és ha valami kimaradt, hány sor. Ha inkább kézzel adnád meg, készíts egy `data/real/gyartott_mennyiseg.csv` fájlt – ha ez létezik, az app ezt használja:

```
hónap,gyártott_db
2025-01-01,5230
2025-02-01,4980
```

## Kipróbálás a saját adataiddal

A program a **saját gépeden** fut, az adataid nem kerülnek fel sehova.

1. **Letöltés:** a GitHub-oldalon kattints a zöld **Code** gombra, majd a **Download ZIP**-re. Csomagold ki egy mappába.
2. **Első indítás próbaadattal:** kattints duplán az `inditas.bat`-ra (Windows). Mindent feltelepít, és a böngészőben megnyílik a dashboard a kitalált adatokkal.
3. **Saját NMF-tábla:** nézd meg a `minta/NMF_minta.xlsx` fájlt. Ilyen oszlopnevekkel kell a táblának lennie:

   | Oszlop | Kötelező? | Mit írj bele |
   |---|---|---|
   | `Dátum` | igen | a nemmegfelelőség dátuma |
   | `Nem megfelelőség jellege` | igen | hibatípus (pl. Méreteltérés) |
   | `Nem megfelelőség oka` | igen | ok (pl. Gép, Ember) |
   | `Vevő` | igen | megrendelő vagy termékcsalád |
   | `Becsült költségvonzat Ft` | igen | becsült költség forintban |
   | `n=nincs teendő …` | igen | intézkedési kód: n, f, o, b, t vagy i (lásd lent) |
   | `Nmf db` | igen | hány darab nem megfelelő |
   | `Munkaszám`, `Cikkszám` | nem | az Ismétlődések fülhöz |
   | `Keletkezés helye (Művelet)`, `NMF keletkezés helye (gép)`, `Következmény` | nem | további szűrők; a `Következmény` „Teljes selejt” / „Másodosztály” értékei a veszteség-p-kártyához kellenek |

   Ha a te tábládban másképp hívják az oszlopokat, két lehetőséged van: átnevezed őket az Excelben, vagy átírod a nevüket a `valos_adat.py` tetején, az `OSZLOPOK` részben.
4. **Gyártott mennyiség (a p-kártyához):** a `minta/gyartott_mennyiseg.csv` mintájára készíts egy táblát (hónap, gyártott_db), és ugyanezzel a névvel mentsd a `data/real/` mappába. Ha ez hiányzik, a p-kártya kivételével minden működik.
5. Másold a saját Exceledet a `data/real/` mappába, indítsd újra a programot, és az oldalsávban kapcsold be a **„Valós adat használata”** kapcsolót.

## A projekt feltöltése a GitHubra (a készítőnek)

A `github_csomag.py` egy külön mappába (`github_csomag/`) másolja azokat a fájlokat, amelyek felkerülhetnek: a kódot, a leírásokat, a próbaadatot és a mintát. A valós adat soha nem kerül bele, ezt a szkript ellenőrzi is.

```bash
.venv\Scripts\python.exe github_csomag.py
```

Utána a GitHubon: **New repository**, majd **uploading an existing file**, és oda húzd be a `github_csomag/` mappa **tartalmát** (ne az egész projektmappát!). A webes feltöltés nem nézi a `.gitignore`-t, ezért fontos, hogy csak ebből a mappából tölts fel.

## Leggyorsabb indítás: dupla kattintás

Kattints duplán az **`inditas.bat`** fájlra. Mindent telepít, ami kell:
1. ha nincs Python a gépen, feltelepíti (winget-tel, csak a te felhasználódnak, rendszergazdai jog nélkül; ehhez internet kell),
2. létrehozza a `.venv` környezetet (ha hiányzik vagy sérült, újra),
3. feltelepíti / ellenőrzi a csomagokat (első alkalommal pár perc, utána pár másodperc),
4. létrehozza a `data/real/` mappát, és ha kell, legenerálja a szintetikus adatot,
5. elindítja a dashboardot, ami magától megnyílik a böngészőben.

Ha a Python automatikus telepítése nem sikerül (pl. nincs winget a gépen), az ablak kiírja, mit kell kézzel tenni.

Leállítás: zárd be a fekete ablakot. Tipp: jobb klikk az `inditas.bat`-on → *Küldés* → *Asztal (parancsikon létrehozása)*, így az asztalról is indítható.

Ha inkább kézzel, parancssorból indítanád, lásd lent.

> Fontos: parancssorba mindig **csak a parancsot** másold be (pl. `pip install -r requirements.txt`), az előtte álló magyarázó szöveget (pl. „Telepítés:”) ne.

## Telepítés kézzel (egyszer kell)

1. Telepítsd a Pythont (3.10 vagy újabb): https://www.python.org/downloads/
2. Nyiss egy terminált (PowerShell) ebben a mappában.
3. (Ajánlott) Hozz létre egy külön környezetet, és aktiváld:

```bash
python -m venv .venv
```

```bash
.venv\Scripts\Activate.ps1
```

   > Ha az aktiválás „running scripts is disabled” hibát ad, futtasd előbb ezt (csak az aktuális ablakra érvényes):
   > `Set-ExecutionPolicy -Scope Process Bypass`
   >
   > A `.venv` azért is ajánlott, mert a Microsoft Store-os Pythonnál a túl hosszú telepítési útvonal miatt a `pip install` hibát adhat.

4. Telepítsd a csomagokat:

```bash
pip install -r requirements.txt
```

## Futtatás

1. Szintetikus adat előállítása (elég egyszer; mindig ugyanazt adja a fix seed miatt):

```bash
python generate_data.py
```

2. A dashboard indítása:

```bash
streamlit run app.py --server.address localhost
```

   (A `--server.address localhost` miatt csak a saját gépedről érhető el.)

A böngészőben megnyílik a `http://localhost:8501` cím. Leállítás: a terminálban **Ctrl+C**.

## Fájlok

| Fájl | Mit csinál |
|---|---|
| `inditas.bat` | Dupla kattintásos indító (telepít, ha kell, és elindítja az appot). |
| `app.py` | Maga a dashboard. |
| `generate_data.py` | Szintetikus adatot készít (24 hónap, enyhén csökkenő hibaarány, seed = 42). |
| `valos_adat.py` | A valós Excel beolvasása és oszlop-megfeleltetése. |
| `szamitasok.py` | Közös számítások (trend, Pareto, p-kártya, ismétlődések) – a képernyő és a beszámoló is ezt használja. |
| `beszamolo.py` | A PowerPoint- és PDF-beszámoló elkészítése. |
| `forditas.py` | Magyar / angol nyelv: szám- és pénzformátum, a próbaadat és a státuszok fordítása. |
| `minta_keszit.py` | Elkészíti a `minta/` kitöltési mintát (kitalált adattal). |
| `minta/` | Kitöltési minta: `NMF_minta.xlsx` és `gyartott_mennyiseg.csv`. |
| `github_csomag.py` | Feltöltésre kész mappát készít (`github_csomag/`), valós adat nélkül. |
| `requirements.txt` | A szükséges Python-csomagok. |
| `.streamlit/config.toml` | Streamlit-beállítások (nincs statisztika-küldés). |
| `.gitignore` | Amit a Git soha ne vegyen fel (valós adat, AI-napló). |
