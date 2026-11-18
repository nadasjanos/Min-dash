# -----------------------------------------------------------------------------
# valos_adat.py
# EZ A FÁJL: beolvassa a VALÓS nemmegfelelőségi Excel-táblát a data/real/
# mappából, és átalakítja a dashboard belső formátumára:
#     dátum, hibatípus, ok, termékcsalád, költség, státusz, nmf_db
# Csak a saját gépeden olvas fájlt, semmit nem küld sehova.
#
# Ha az Excelben máshogy hívnak egy oszlopot, csak az OSZLOPOK szótárban
# (lent) kell átírni a nevét – a kód többi részéhez nem kell nyúlni.
# -----------------------------------------------------------------------------

import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# Az ennél korábbi vagy jövő évnél későbbi dátumot elgépelésnek tekintjük
LEGKORABBI_EV = 2000

# --- Beállítások -------------------------------------------------------------
# Melyik munkalapot olvassa (0 = az első lap, vagy írd be a lap nevét idézőjelben)
MUNKALAP = 0
# Hányadik sorban vannak az oszlopnevek (0 = az első sor)
FEJLEC_SOR = 0

# Belső név  ->  oszlopnév az Excelben.
# Az oszlopneveket a program "kisimítja": a sortöréseket és a több szóközt
# egyetlen szóközre cseréli, ezért itt is egy szóközzel írd őket.
# Elég az oszlopnév ELEJÉT megadni, ha az egyértelmű (pl. a hosszú
# "n=nincs teendő f=folyamtban, ..." oszlopnál).
OSZLOPOK = {
    "dátum": "Dátum",
    "hibatípus": "Nem megfelelőség jellege",
    "ok": "Nem megfelelőség oka",
    # Nincs külön "termékcsalád" oszlop, ezért a vevő szerint csoportosítunk.
    # Átírhatod pl. "Cikkszám"-ra vagy "Keletkezés helye (Művelet)"-re.
    "termékcsalád": "Vevő",
    "költség": "Becsült költségvonzat Ft",
    "státusz": "n=nincs teendő",  # az intézkedési kód oszlopa (n/f/o/b/t/i)
    "nmf_db": "Nmf db",            # hány darab volt nem megfelelő az adott esetben
}

# Nem kötelező oszlopok: ha megvannak az Excelben, szűrhetsz rájuk az
# oldalsávban; ha hiányoznak, egyszerűen nem jelenik meg a szűrő.
OPCIONALIS_OSZLOPOK = {
    "keletkezés_helye": "Keletkezés helye (Művelet)",
    "gép": "NMF keletkezés helye (gép)",
    "következmény": "Következmény",
    "munkaszám": "Munkaszám",   # az "Ismétlődések" fülhöz
    "cikkszám": "Cikkszám",     # az "Ismétlődések" fülhöz
}
# Azonosítók: ezeknél NEM vonjuk össze az írásmódokat (csak a szóközöket
# tisztítjuk), mert két kicsit eltérő szám két külön tétel is lehet.
AZONOSITOK = {"munkaszám", "cikkszám"}

# Az intézkedési kódok jelentése (az Excel fejlécéből) -> státusz a dashboardon.
# "Lezárt" = nincs teendő; minden más nyitott ügynek számít.
STATUSZ_KODOK = {
    "n": "Lezárt",
    "f": "Folyamatban",
    "o": "Nyitott – oktatás",
    "b": "Nyitott – beszállító",
    "t": "Nyitott – technológia",
    "i": "Nyitott – intézkedés szükséges",
}


def _oszlop_keres(oszlopok, nev: str):
    """Megkeresi az oszlopot pontos névre, ha nincs, akkor a név elejére."""
    if nev in oszlopok:
        return nev
    talalat = [o for o in oszlopok if o.startswith(nev)]
    return talalat[0] if len(talalat) == 1 else None


def _kisimit(nev) -> str:
    """Oszlopnévből eltünteti a sortöréseket és a felesleges szóközöket."""
    return " ".join(str(nev).split())


def _osszehasonlito_kulcs(szoveg: str) -> str:
    """Egy név "lényege": kisbetűs, ékezet, szóköz és írásjel nélkül.
    Pl. "Belső gyártás", "Belsőgyártás", "BELSO GYARTAS" -> "belsogyartas"."""
    ekezet_nelkul = "".join(c for c in unicodedata.normalize("NFKD", szoveg)
                            if not unicodedata.combining(c))
    return re.sub(r"[\W_]+", "", ekezet_nelkul.lower())


def _irasmodok_osszevonasa(oszlop: pd.Series) -> pd.Series:
    """A csak írásmódban eltérő értékeket egy névre vonja össze.
    Az összevont csoport neve a leggyakoribb írásmód lesz."""
    kulcs = oszlop.map(_osszehasonlito_kulcs)
    # minden kulcshoz a leggyakoribb eredeti írásmód
    leggyakoribb = (pd.DataFrame({"kulcs": kulcs, "nev": oszlop})
                    .groupby("kulcs")["nev"].agg(lambda s: s.value_counts().index[0]))
    return kulcs.map(leggyakoribb)


def _datum_atalakit(ertek, alap_ev: int):
    """Egy cella értékét dátummá alakítja. Kezeli:
    - a rendes Excel-dátumot,
    - a szövegként beírt dátumot: "2025.01.15", "2025. 01. 15.", "01.15.", "01. 15."
      (év nélkül az alap_ev-et használja),
    - az Excel belső dátumszámát (pl. 45672).
    Ha nem értelmezhető, vagy nem reális az év, NaT-t (üres dátumot) ad."""
    if ertek is None or pd.isna(ertek):
        return pd.NaT
    if isinstance(ertek, (datetime, date)):
        if not LEGKORABBI_EV <= ertek.year <= date.today().year + 1:
            return pd.NaT  # pl. 1900 vagy 0001 – elgépelés
        eredmeny = pd.Timestamp(ertek)
    elif isinstance(ertek, (int, float)):
        if not 30000 < ertek < 80000:   # kb. 1982 és 2119 közötti Excel-dátumszám
            return pd.NaT
        eredmeny = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(ertek))
    else:
        szoveg = re.sub(r"\s+", "", str(ertek)).rstrip(".")  # szóközök és záró pont le
        teljes = re.fullmatch(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", szoveg)
        ev_nelkul = re.fullmatch(r"(\d{1,2})[.\-/](\d{1,2})", szoveg)
        if teljes:
            ev, ho, nap = map(int, teljes.groups())
        elif ev_nelkul:
            ev, (ho, nap) = alap_ev, map(int, ev_nelkul.groups())
            if ho > 12 >= nap:          # ha nap.hónap sorrendben írták
                ho, nap = nap, ho
        else:
            return pd.NaT
        try:
            eredmeny = pd.Timestamp(year=ev, month=ho, day=nap)
        except ValueError:              # pl. 02.30. nem létező nap
            return pd.NaT
    if not LEGKORABBI_EV <= eredmeny.year <= date.today().year + 1:
        return pd.NaT
    return eredmeny


# --- Gyártott mennyiség a p-kártyához -----------------------------------------
# A késztermék-bevételezés exportja (CSV, pontosvesszővel elválasztva). A fájl
# elején cím- és üres sorok vannak, a fejléc oldalanként ismétlődik – ezért a
# program magától megkeresi a fejlécet, és csak a dátumos sorokat veszi.
TERMELES_FAJL_MINTA = "TermelesBeerkezes*.csv"
TERMELES_DATUM = "Bevét dátum"
TERMELES_MENNYISEG = "Mennyiség"


def termeles_fajl(mappa: Path):
    """Visszaadja a termelési (bevételezési) CSV-t, ha van ilyen a mappában."""
    talalat = sorted(mappa.glob(TERMELES_FAJL_MINTA)) if mappa.exists() else []
    return talalat[0] if talalat else None


def termeles_sorok(fajl: Path) -> pd.DataFrame:
    """A bevételezési lista adatsorai: dátum, mennyiség (cím- és fejlécsorok nélkül)."""
    nyers = pd.read_csv(fajl, sep=";", encoding="utf-8-sig", header=None, dtype=str)
    # A fejlécsor az, amelyikben szerepel a "Bevét dátum" felirat
    for i, sor in nyers.iterrows():
        cellak = [str(c).strip() for c in sor]
        if TERMELES_DATUM in cellak and TERMELES_MENNYISEG in cellak:
            d_oszlop = cellak.index(TERMELES_DATUM)
            m_oszlop = cellak.index(TERMELES_MENNYISEG)
            break
    else:
        raise ValueError(f"A {fajl.name} fájlban nem találom a '{TERMELES_DATUM}' és "
                         f"'{TERMELES_MENNYISEG}' oszlopot.")

    datum = pd.to_datetime(nyers[d_oszlop].str.strip(), format="%Y.%m.%d", errors="coerce")
    # Magyar számformátum: "1 234,5" -> 1234.5
    mennyiseg = pd.to_numeric(nyers[m_oszlop].str.replace(r"\s", "", regex=True)
                              .str.replace(",", "."), errors="coerce")
    jo = datum.notna() & mennyiseg.notna()  # csak az igazi adatsorok (nem cím / fejléc)
    return pd.DataFrame({"dátum": datum[jo], "mennyiség": mennyiseg[jo]}).reset_index(drop=True)


def termeles_betolt(sorok: pd.DataFrame) -> pd.DataFrame:
    """A bevételezési sorokból havi gyártott mennyiséget számol.
    Eredmény oszlopai: hónap, gyártott_db (ugyanaz, mint a gyartott_mennyiseg.csv-é)."""
    return (pd.DataFrame({"hónap": sorok["dátum"].dt.to_period("M").dt.to_timestamp(),
                          "gyártott_db": sorok["mennyiség"]})
            .groupby("hónap", as_index=False)["gyártott_db"].sum())


def kozos_idoszak(nm: pd.DataFrame, sorok: pd.DataFrame):
    """Csak azok az NMF-sorok maradnak, amelyek dátuma a bevételezési CSV
    időszakába esik (első és utolsó bevételezési nap között).
    Így a hibás és a gyártott darab ugyanarra az időszakra vonatkozik.
    Visszaad: (megmaradt NMF-sorok, kiesett sorok száma, (első nap, utolsó nap))."""
    elso, utolso = sorok["dátum"].min().normalize(), sorok["dátum"].max().normalize()
    nap = nm["dátum"].dt.normalize()
    benne = nap.between(elso, utolso)
    return nm[benne].reset_index(drop=True), int((~benne).sum()), (elso, utolso)


# A bevételezési CSV a késztermékraktár beérkeztetése: a jó termékek és a
# selejtpótló munkaszámok. A selejt- vagy másodosztály-raktárba érkeztetett
# darabok NEM szerepelnek benne – azok itt csak az NMF táblából ismertek. Ezért a teljes gyártott
# mennyiséghez ezeket hozzá kell adni. (A javított darabok a javítás után a
# bevételezésben már benne vannak, őket nem adjuk hozzá még egyszer.)
NEM_BEERKEZETT_KOVETKEZMENYEK = ["Teljes selejt", "Másodosztály"]


def teljes_gyartott(havi: pd.DataFrame, nm: pd.DataFrame) -> pd.DataFrame:
    """Havi gyártott db = bevételezés + a selejtre / másodosztályra került Nmf db.
    Új oszlopok: 'bevételezés_db', 'selejt_2o_db'; a 'gyártott_db' a kettő összege."""
    havi = havi.copy()
    havi["bevételezés_db"] = havi["gyártott_db"]
    if "következmény" in nm.columns:
        kiesett = nm[nm["következmény"].isin(NEM_BEERKEZETT_KOVETKEZMENYEK)]
        havonta = (kiesett.groupby(kiesett["dátum"].dt.to_period("M").dt.to_timestamp())
                   ["nmf_db"].sum())
        havi["selejt_2o_db"] = havi["hónap"].map(havonta).fillna(0)
    else:
        havi["selejt_2o_db"] = 0
    havi["gyártott_db"] = havi["bevételezés_db"] + havi["selejt_2o_db"]
    return havi


def valos_fajlok(mappa: Path) -> list[Path]:
    """Visszaadja a data/real/ mappában lévő Excel-fájlokat."""
    if not mappa.exists():
        return []
    return sorted(p for p in mappa.iterdir()
                  if p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
                  and not p.name.startswith("~$"))  # ~$ = Excel ideiglenes fájlja


def betolt(fajl: Path) -> tuple[pd.DataFrame, list[int]]:
    """Beolvassa a valós Excelt és belső formátumra alakítja.
    Visszaadja: (adattábla, a hibás/hiányzó dátum miatt kihagyott Excel-sorok száma)."""
    nyers = pd.read_excel(fajl, sheet_name=MUNKALAP, header=FEJLEC_SOR)
    nyers.columns = [_kisimit(c) for c in nyers.columns]

    # Belső név -> a ténylegesen megtalált Excel-oszlop
    oszlop = {belso: _oszlop_keres(list(nyers.columns), nev) for belso, nev in OSZLOPOK.items()}
    hianyzo = [OSZLOPOK[belso] for belso, talalt in oszlop.items() if talalt is None]
    if hianyzo:
        raise ValueError("Ezek az oszlopok nem találhatók az Excelben: "
                         + ", ".join(hianyzo)
                         + ". Javítsd az OSZLOPOK szótárat a valos_adat.py fájlban.")

    # Az "alap év" a rendes dátumok leggyakoribb éve – ezt kapják az év nélkül
    # beírt dátumok (pl. "12.05.").
    rendes = pd.to_datetime(nyers[oszlop["dátum"]].map(
        lambda x: x if isinstance(x, (datetime, date)) else None), errors="coerce")
    jo_evek = rendes.dt.year[(rendes.dt.year >= LEGKORABBI_EV)]
    alap_ev = int(jo_evek.mode()[0]) if len(jo_evek) else date.today().year

    df = pd.DataFrame()
    df["dátum"] = pd.to_datetime(nyers[oszlop["dátum"]].map(
        lambda x: _datum_atalakit(x, alap_ev)))
    # Szöveges oszlopok: felesleges szóközök le, üres cella -> "Nincs megadva"
    szoveges = {b: oszlop[b] for b in ["hibatípus", "ok", "termékcsalád"]}
    for belso, nev in OPCIONALIS_OSZLOPOK.items():
        talalt = _oszlop_keres(list(nyers.columns), nev)
        if talalt is not None:
            szoveges[belso] = talalt
    for belso, excel_oszlop in szoveges.items():
        tiszta = (nyers[excel_oszlop].astype("string")
                  .str.replace(r"\s+", " ", regex=True).str.strip()
                  .replace("", pd.NA).fillna("Nincs megadva"))
        if belso in AZONOSITOK:
            df[belso] = tiszta.astype(str)
        else:
            # "Belső gyártás" és "Belsőgyártás" -> egy név (a leggyakoribb írásmód)
            df[belso] = _irasmodok_osszevonasa(tiszta.astype(str))
    df["költség"] = pd.to_numeric(nyers[oszlop["költség"]], errors="coerce").fillna(0)
    # Státusz: az intézkedési kód első betűje (kisbetűsítve) alapján
    kod = nyers[oszlop["státusz"]].astype("string").str.strip().str.lower().str[:1]
    df["státusz"] = kod.map(STATUSZ_KODOK).fillna("Nincs megadva")
    df["nmf_db"] = pd.to_numeric(nyers[oszlop["nmf_db"]], errors="coerce").fillna(1)

    # Hibás dátumú sorok: van bennük adat, de a dátum hiányzik / nem értelmezhető.
    # (A teljesen üres sorokat, pl. az Excel alján, nem jelezzük.)
    van_adat = nyers[[oszlop[b] for b in ["hibatípus", "ok", "termékcsalád", "nmf_db"]]].notna().any(axis=1)
    hibas = df["dátum"].isna() & van_adat
    # Excel-sorszám: a pandas 0-tól számol, + fejlécsor + 1-től számozás
    hibas_sorok = [int(i) + FEJLEC_SOR + 2 for i in df.index[hibas]]

    return df.dropna(subset=["dátum"]).reset_index(drop=True), hibas_sorok
