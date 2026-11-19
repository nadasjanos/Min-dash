# -----------------------------------------------------------------------------
# app.py
# EZ A FÁJL: maga a minőségügyi dashboard (Streamlit webes felület).
# Csak a saját gépeden fut (localhost), semmit nem tölt fel az internetre.
#
# Indítás:  streamlit run app.py
#
# Nézetek (fülek):
#   1) Havi trend            – hány nemmegfelelőség és mennyi költség volt havonta
#   2) Pareto-diagram        – melyik hibatípus okozza a legtöbb problémát
#   3) Ismétlődések          – többször felvett cikkek és munkaszámok
#   4) p-szabályozókártya    – a havi hibaarány statisztikailag stabil-e
#   5) AI-ellenőrzési napló  – kézzel rögzítheted, mit ellenőriztél az AI kimenetén
#
# Adatforrás: alapból a data/synthetic/ (kitalált) adat. A data/real/ (valós)
# adatot csak az oldalsávban lévő kapcsolóval lehet bekapcsolni.
#
# Nyelv: magyar vagy angol (oldalsáv teteje). A szövegek mindkét nyelven itt
# vannak egymás mellett:  L("magyar szöveg", "English text").
# Angol nyitás linkkel: ...?lang=en
# -----------------------------------------------------------------------------

import os
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

import beszamolo  # a saját fájlunk: PowerPoint / PDF beszámoló készítése
import forditas  # a saját fájlunk: magyar / angol nyelv
import szamitasok as sz  # a saját fájlunk: közös számítások (képernyő + beszámoló)
import valos_adat  # a saját fájlunk: a valós Excel beolvasása

# --- Mappák és fájlnevek -----------------------------------------------------
ALAP_MAPPA = Path(__file__).parent / "data"
SZINTETIKUS_MAPPA = ALAP_MAPPA / "synthetic"
VALOS_MAPPA = ALAP_MAPPA / "real"
NM_FAJL = "nemmegfelelosegek.csv"          # nemmegfelelőségek listája
MENNYISEG_FAJL = "gyartott_mennyiseg.csv"  # havi gyártott darabszám
AI_NAPLO_FAJL = ALAP_MAPPA / "ai_naplo.csv"  # az AI-ellenőrzési napló (helyi fájl)

# Online bemutató mód (Streamlit Community Cloud): ott nincs valós adat, és az
# AI-naplót sem írjuk közös fájlba, mert azt minden látogató látná.
# A Community Cloud a repót a /mount/src/ mappába teszi – erről ismerjük fel.
# Helyben is bekapcsolható: MINOSEG_DEMO=1 környezeti változóval.
DEMO = (os.environ.get("MINOSEG_DEMO") == "1"
        or Path(__file__).resolve().as_posix().startswith("/mount/src/"))

KOTELEZO_OSZLOPOK = ["dátum", "hibatípus", "ok", "termékcsalád", "költség", "státusz"]
AI_NAPLO_OSZLOPOK = ["dátum", "AI eszköz", "feladat / AI kimenet", "ellenőrzött szempontok",
                     "eredmény", "megjegyzés"]

# --- Nyelv -------------------------------------------------------------------
# Az első megnyitáskor a link dönt (?lang=en -> angol), utána az oldalsáv választója.
NYELVEK = ["Magyar", "English"]
if "nyelv" not in st.session_state:
    st.session_state["nyelv"] = ("English" if st.query_params.get("lang", "").lower() == "en"
                                 else "Magyar")
EN = st.session_state["nyelv"] == "English"
L = forditas.nyelv(EN)  # L("magyar", "angol") -> a választott nyelvű szöveg


def szam(x):
    return forditas.szam(x, EN)


def ft(x):
    return forditas.ft(x, EN)


def pct(x, tizedes=1):
    return forditas.pct(x, tizedes, EN)


st.set_page_config(page_title=L("Minőségügyi dashboard", "Quality dashboard"), layout="wide")


# --- Adatbetöltés ------------------------------------------------------------
# A @st.cache_data miatt a fájlt csak egyszer olvassa be, utána a memóriából veszi
# (gyorsabb). Ha a fájl módosul, az "R" billentyűvel / újraindítással frissül.
@st.cache_data
def adat_betoltes(mappa: Path, valos_fajl: Path | None = None, verzio: float = 0.0):
    # A "verzio" a fájlok utolsó módosítási ideje: ha átírod az Excelt / CSV-t,
    # ez megváltozik, és az app újraolvassa a fájlt (nem a régit mutatja).
    """Beolvassa a nemmegfelelőségeket és (ha van) a havi gyártott mennyiséget.
    Szintetikus adatnál CSV-t olvas, valós adatnál Excelt (valos_adat.py)."""
    hibas_sorok = []  # a hibás dátum miatt kihagyott Excel-sorok
    if valos_fajl is not None:
        nm, hibas_sorok = valos_adat.betolt(valos_fajl)
    else:
        nm = pd.read_csv(mappa / NM_FAJL, encoding="utf-8-sig", parse_dates=["dátum"])
    if "nmf_db" not in nm.columns:
        nm["nmf_db"] = 1  # szintetikus adatban minden sor 1 hibás darab
    mennyiseg = None
    kozos = None  # valós adatnál: (kiesett NMF-sorok száma, (első nap, utolsó nap))
    if (mappa / MENNYISEG_FAJL).exists():
        # 1. választás: kézzel készített gyartott_mennyiseg.csv
        mennyiseg = pd.read_csv(mappa / MENNYISEG_FAJL, encoding="utf-8-sig",
                                parse_dates=["hónap"])
    elif valos_fajl is not None and valos_adat.termeles_fajl(mappa) is not None:
        # 2. választás: a késztermék-bevételezés exportja (TermelesBeerkezes*.csv).
        sorok = valos_adat.termeles_sorok(valos_adat.termeles_fajl(mappa))
        # Csak a közös időszak marad: az NMF-ből kiesik, ami a CSV időszakán kívül van
        nm, kiesett, idoszak = valos_adat.kozos_idoszak(nm, sorok)
        kozos = (kiesett, idoszak)
        # kiegészítve a selejtre / másodosztályra került darabokkal (azok nincsenek benne)
        mennyiseg = valos_adat.teljes_gyartott(valos_adat.termeles_betolt(sorok), nm)
    return nm, mennyiseg, hibas_sorok, kozos


def ai_naplo_betoltes() -> pd.DataFrame:
    """Beolvassa az AI-ellenőrzési naplót; ha még nincs, üres táblát ad.
    Online bemutatóban csak a látogató saját böngészőmenetében él (nem fájlban)."""
    if DEMO:
        return st.session_state.get("ai_naplo", pd.DataFrame(columns=AI_NAPLO_OSZLOPOK))
    if AI_NAPLO_FAJL.exists():
        return pd.read_csv(AI_NAPLO_FAJL, encoding="utf-8-sig")
    return pd.DataFrame(columns=AI_NAPLO_OSZLOPOK)


def nyelvvaltas():
    """Nyelvváltáskor a szűrők törlődnek (a kategóriák neve is megváltozik)."""
    for kulcs in [k for k in st.session_state if str(k).startswith("szuro_")]:
        del st.session_state[kulcs]
    st.session_state.pop("beszamolo", None)


# --- Oldalsáv: nyelv, adatforrás és szűrők -----------------------------------
st.sidebar.radio("Nyelv / Language", NYELVEK, key="nyelv", horizontal=True,
                 on_change=nyelvvaltas)
st.sidebar.header(L("Beállítások", "Settings"))
if DEMO:
    # Online bemutató: csak a kitalált adat, saját adathoz le kell tölteni
    valos_adat_kell = False
    st.sidebar.info(L(
        "🌐 **Online bemutató** kitalált adatokkal. A saját adataiddal a programot "
        "töltsd le, és a saját gépeden futtasd: ",
        "🌐 **Online demo** with made-up data. To use your own data, download the "
        "program and run it on your own computer: ")
        + "[GitHub – Min-dash](https://github.com/nadasjanos/Min-dash)")
else:
    valos_adat_kell = st.sidebar.toggle(
        L("Valós adat használata (data/real/)", "Use real data (data/real/)"),
        value=False,  # alapból KI van kapcsolva -> szintetikus adat
        key="valos_kapcsolo",  # állandó kulcs: nyelvváltáskor is megmarad a beállítás
        help=L("Kikapcsolva a kitalált (szintetikus) adatot látod.",
               "When off, you see made-up (synthetic) data."),
    )
mappa = VALOS_MAPPA if valos_adat_kell else SZINTETIKUS_MAPPA

valos_fajl = None
if valos_adat_kell:
    fajlok = valos_adat.valos_fajlok(VALOS_MAPPA)
    if not fajlok:
        st.warning(L("Nem található Excel-fájl a `data/real/` mappában. "
                     "Másold oda a nemmegfelelőségi táblát (.xlsx).",
                     "No Excel file found in the `data/real/` folder. "
                     "Copy your nonconformity table (.xlsx) there."))
        st.stop()
    valos_fajl = st.sidebar.selectbox(L("Valós fájl", "Real data file"), fajlok,
                                      format_func=lambda p: p.name, key="valos_fajl")
elif not (mappa / NM_FAJL).exists():
    st.warning(L("Még nincs szintetikus adat. Futtasd le: `python generate_data.py`",
                 "No synthetic data yet. Run: `python generate_data.py`"))
    st.stop()

try:
    # a mappa összes fájljának legutóbbi módosítási ideje -> ha bármi változik, újraolvas
    verzio = max((f.stat().st_mtime for f in mappa.iterdir() if f.is_file()), default=0.0)
    nm, mennyiseg, hibas_sorok, kozos = adat_betoltes(mappa, valos_fajl, verzio)
except ValueError as hiba:  # pl. hiányzó oszlop a valós Excelben
    st.error(str(hiba))
    st.stop()

hianyzo = [o for o in KOTELEZO_OSZLOPOK if o not in nm.columns]
if hianyzo:
    st.error(L(f"A {NM_FAJL} fájlból hiányoznak ezek az oszlopok: {', '.join(hianyzo)}",
               f"These columns are missing from {NM_FAJL}: {', '.join(hianyzo)}"))
    st.stop()

# Angol nézetben a státuszok és (próbaadatnál) a kategóriák angolul látszanak
if EN:
    nm = forditas.adat_forditas(nm, szintetikus=not valos_adat_kell)

# --- Szűrők ------------------------------------------------------------------
nm["hónap"] = nm["dátum"].dt.to_period("M").dt.to_timestamp()

# belső oszlopnév -> felirat (csak a létező oszlopok jelennek meg).
# Ugyanezt a listát használja a Pareto-diagram bontás-választója is.
BONTASOK = {
    "termékcsalád": (L("Megrendelő (Vevő)", "Customer") if valos_adat_kell
                     else L("Termékcsalád", "Product family")),
    "hibatípus": (L("Nemmegfelelőség jellege", "Nonconformity type") if valos_adat_kell
                  else L("Hibatípus", "Defect type")),
    "ok": (L("Nemmegfelelőség oka", "Nonconformity cause") if valos_adat_kell
           else L("Ok", "Cause")),
    "keletkezés_helye": L("Keletkezés helye (művelet)", "Origin (operation)"),
    "gép": L("NMF keletkezés helye (gép)", "Origin (machine)"),
    "következmény": L("Következmény", "Consequence"),
}
BONTASOK = {o: f for o, f in BONTASOK.items() if o in nm.columns}
SZURO_OSZLOPOK = [o for o in BONTASOK if o != "ok"]  # az okra nem kértünk szűrőt

honapok = list(pd.DatetimeIndex(nm["hónap"].unique()).sort_values())


def szurok_torlese(teljes_tartomany):
    """A 'Szűrők törlése' gomb: a hónap a teljes időszak lesz, a többi 'Mind'."""
    st.session_state["szuro_hónap"] = teljes_tartomany
    for oszlop in SZURO_OSZLOPOK:
        st.session_state[f"szuro_{oszlop}"] = []


st.sidebar.subheader(L("Szűrők", "Filters"))
szurt = nm.copy()

# Hónap: csúszka, két végével adod meg a "-tól" és "-ig" hónapot
tol, ig = honapok[0], honapok[-1]  # alapból a teljes időszak (a p-kártya is ezt használja)
if len(honapok) > 1:
    tol, ig = st.sidebar.select_slider(
        L("Hónap (-tól, -ig)", "Month (from, to)"), options=honapok,
        value=(honapok[0], honapok[-1]),
        format_func=lambda d: pd.Timestamp(d).strftime("%Y. %m." if not EN else "%b %Y"),
        key="szuro_hónap")
    szurt = szurt[(szurt["hónap"] >= tol) & (szurt["hónap"] <= ig)]

# A többi szűrő üresen indul = "Mind". Ha kiválasztasz értékeket, csak azok maradnak.
aktiv_szurok = {}  # a beszámoló címlapjára: melyik szűrő mire van állítva
for oszlop in SZURO_OSZLOPOK:
    valasztott = st.sidebar.multiselect(BONTASOK[oszlop], sorted(nm[oszlop].unique()),
                                        key=f"szuro_{oszlop}", placeholder=L("Mind", "All"))
    if valasztott:
        szurt = szurt[szurt[oszlop].isin(valasztott)]
        aktiv_szurok[BONTASOK[oszlop]] = valasztott
st.sidebar.button(L("Szűrők törlése", "Clear filters"), on_click=szurok_torlese,
                  args=((honapok[0], honapok[-1]),))
st.sidebar.caption(L(f"Szűrés után {len(szurt)} / {len(nm)} sor látszik.",
                     f"{len(szurt)} / {len(nm)} rows shown after filtering."))

st.sidebar.caption(L("Adatforrás: ", "Data source: ")
                   + (L("🔴 VALÓS (data/real/)", "🔴 REAL (data/real/)") if valos_adat_kell
                      else L("🟢 szintetikus (data/synthetic/)", "🟢 synthetic (data/synthetic/)")))

# --- Fejléc és kulcsszámok ---------------------------------------------------
st.title(L("Minőségügyi dashboard – nemmegfelelőségek", "Quality dashboard – nonconformities"))
if hibas_sorok:
    st.warning(L(f"{len(hibas_sorok)} sor kimaradt, mert a dátuma hiányzik vagy hibás "
                 "(pl. 1900-as év). Javítsd ezeket az Excelben – sorok: ",
                 f"{len(hibas_sorok)} rows were skipped because their date is missing or "
                 "invalid (e.g. year 1900). Fix them in Excel – rows: ")
               + ", ".join(map(str, hibas_sorok)))
if kozos is not None:
    kiesett, (elso, utolso) = kozos
    szoveg = L(f"Vizsgált időszak: {elso:%Y.%m.%d.} – {utolso:%Y.%m.%d.} "
               "(az NMF tábla és a bevételezési CSV közös időszaka).",
               f"Period analysed: {elso:%Y-%m-%d} – {utolso:%Y-%m-%d} "
               "(common period of the NC table and the receipts CSV).")
    if kiesett:
        st.warning(szoveg + L(f" {kiesett} NMF-sor kimaradt, mert a dátuma ezen kívül esik.",
                              f" {kiesett} NC rows were skipped because they fall outside it."))
    else:
        st.caption(szoveg + L(" Minden NMF-sor ebbe az időszakba esik.",
                              " All NC rows fall within this period."))
k1, k2, k3, k4 = st.columns(4)
k1.metric(L("Nemmegfelelőségek száma", "Number of nonconformities"), szam(len(szurt)))
k2.metric(L("Összköltség", "Total cost"), ft(szurt["költség"].sum()))
k3.metric(L("Nyitott / folyamatban", "Open / in progress"),
          int(forditas.nyitott_e(szurt["státusz"]).sum()))
k4.metric(L("Átlagköltség / eset", "Average cost / case"),
          ft(szurt["költség"].mean()) if len(szurt) else "–")

tab_trend, tab_pareto, tab_ism, tab_p, tab_ai = st.tabs([
    L("📈 Havi trend", "📈 Monthly trend"), L("📊 Pareto-diagram", "📊 Pareto chart"),
    L("🔁 Ismétlődések", "🔁 Repeat issues"), L("🎯 p-szabályozókártya", "🎯 p control chart"),
    L("📝 AI-ellenőrzési napló", "📝 AI verification log"),
])

# --- 1) Havi trend -----------------------------------------------------------
with tab_trend:
    havi = sz.havi_osszesites(szurt)  # havi esetszám, költség és lineáris trend
    trend = havi["trend"]

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=havi["hónap"], y=havi["darab"],
                name=L("Nemmegfelelőségek (db)", "Nonconformities (count)"),
                marker_color="#4C78A8")
    fig.add_scatter(x=havi["hónap"], y=trend, name=L("Trend (db)", "Trend (count)"),
                    mode="lines", line=dict(color="#4C78A8", dash="dash"))
    fig.add_scatter(x=havi["hónap"], y=havi["koltseg"], name=L("Költség (Ft)", "Cost (HUF)"),
                    mode="lines+markers", line=dict(color="#E45756"), secondary_y=True)
    fig.update_yaxes(title_text=L("darab", "count"), secondary_y=False)
    fig.update_yaxes(title_text=L("Ft", "HUF"), secondary_y=True)
    fig.update_layout(title=L("Havi nemmegfelelőségek és költség",
                              "Monthly nonconformities and cost"),
                      hovermode="x unified", legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, width="stretch")
    st.caption(L("A szaggatott vonal a lineáris trend: ha lefelé tart, javul a helyzet.",
                 "The dashed line is the linear trend: if it goes down, things are improving."))

# --- 2) Pareto-diagram -------------------------------------------------------
with tab_pareto:
    c1, c2 = st.columns(2)
    bontas = c1.selectbox(L("Mi szerint bontsuk?", "Break down by"), list(BONTASOK),
                          index=list(BONTASOK).index("hibatípus"),
                          format_func=lambda o: BONTASOK[o], key="pareto_bontas")
    # A belső érték mindig magyar ("Darabszám" / "Költség"), csak a felirat fordított
    MERTEKEK = {"Darabszám": L("Darabszám", "Count"), "Költség": L("Költség", "Cost")}
    mertek = c2.radio(L("Mi alapján rangsoroljunk?", "Rank by"), list(MERTEKEK),
                      format_func=lambda m: MERTEKEK[m], horizontal=True, key="pareto_mertek")
    mertek_felirat = MERTEKEK[mertek]

    # Pareto-sorrend; 15-nél több kategóriánál a többi "Egyéb"-be kerül
    p_adat = sz.pareto_adat(szurt, bontas, mertek, en=EN)

    if len(p_adat) < 2:
        st.info(L(f"A szűrés után csak {len(p_adat)} érték maradt ebben a bontásban, így nincs "
                  "mit rangsorolni. Válassz másik bontást (pl. a hiba okát), vagy lazíts a "
                  "szűrőkön.",
                  f"Only {len(p_adat)} value(s) left in this breakdown after filtering, so there "
                  "is nothing to rank. Choose another breakdown (e.g. cause) or relax the "
                  "filters."))
    else:
        kumulalt = p_adat.cumsum() / p_adat.sum() * 100  # kumulált százalék

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=p_adat.index, y=p_adat.values, name=mertek_felirat, marker_color="#4C78A8")
        fig.add_scatter(x=p_adat.index, y=kumulalt.values, name=L("Kumulált %", "Cumulative %"),
                        mode="lines+markers", line=dict(color="#E45756"), secondary_y=True)
        fig.add_hline(y=80, line_dash="dot", line_color="gray", secondary_y=True,
                      annotation_text="80%")
        fig.update_yaxes(title_text=mertek_felirat, secondary_y=False)
        fig.update_yaxes(title_text=L("Kumulált %", "Cumulative %"), range=[0, 105],
                         tickvals=[0, 20, 40, 60, 80, 100], showgrid=False, secondary_y=True)
        fig.update_layout(title=L(f"Pareto-diagram – {BONTASOK[bontas].lower()} szerint "
                                  f"({mertek_felirat.lower()})",
                                  f"Pareto chart – by {BONTASOK[bontas].lower()} "
                                  f"({mertek_felirat.lower()})"),
                          legend=dict(orientation="h", y=-0.35))
        st.plotly_chart(fig, width="stretch")
        st.caption(L("80/20 szabály: a 80%-os vonal alatti néhány tételre érdemes először "
                     "fókuszálni. Tipp: szűrj egy hibatípusra, és bonts ok szerint.",
                     "80/20 rule: focus first on the few items below the 80% line. "
                     "Tip: filter for one defect type and break it down by cause."))

# --- Ismétlődések: többször felvett munkaszámok és cikkszámok ----------------
# A számítások a szamitasok.py-ban vannak (a beszámoló is ezeket használja)
RANGSOR = sz.RANGSOR  # magyar felirat -> belső név (esetek / koltseg / hibas_db)
RANGSOR_EN = {"Esetek száma": "Number of cases", "Becsült költség": "Estimated cost",
              "Hibás db": "Defective qty"}
ismetlodok = sz.ismetlodok


def rangsor_cimke(hu_felirat: str) -> str:
    return RANGSOR_EN[hu_felirat] if EN else hu_felirat


def top3_doboz(df: pd.DataFrame, kulcs: str, cim: str, rendezes: str, felirat: str):
    """Megmutatja az első 3 tételt, és hogy együtt hány %-át adják az egésznek.
    Itt MINDEN tétel számít (az egyszer szereplők is), különben a % félrevezető lenne."""
    ertek = sz.rangsor_ertek(df, kulcs, rendezes)
    if ertek.sum() <= 0:
        st.info(L(f"Nincs adat ({cim}).", f"No data ({cim})."))
        return
    reszesedes = ertek / ertek.sum()
    kell_80 = sz.hany_adja_a_80_szazalekot(ertek)  # hány tétel adja a 80%-ot

    st.metric(L(f"Első 3 {cim} részesedése", f"Share of top 3 {cim}"),
              pct(reszesedes.head(3).sum(), 0),
              help=L(f"A szűrt adatok teljes {felirat.lower()} értékéből ennyit ad az első 3.",
                     f"Share of the top 3 in the total {felirat.lower()} of the filtered data."))
    top = pd.DataFrame({
        "#": range(1, min(3, len(ertek)) + 1),
        kulcs: ertek.index[:3],
        felirat: ertek.values[:3],
        L("részesedés", "share"): [pct(x) for x in reszesedes.values[:3]],
        L("kumulált", "cumulative"): [pct(x) for x in reszesedes.cumsum().values[:3]],
    })
    st.dataframe(forditas.oszlopok(top, EN), hide_index=True, width="stretch")
    st.caption(L(f"A 80%-ot az első **{kell_80}** {cim} adja (összesen {len(ertek)} van).",
                 f"The top **{kell_80}** {cim} make up 80% (out of {len(ertek)})."))


with tab_ism:
    if not {"munkaszám", "cikkszám"} & set(szurt.columns):
        st.info(L("Ehhez az adathoz nincs munkaszám / cikkszám oszlop.",
                  "This data has no job number / item number column."))
    else:
        st.write(L("Azok a munkaszámok és cikkek, amelyek többször is bekerültek a "
                   "nemmegfelelőségek közé – rangsorolva. Az oldalsáv szűrői itt is érvényesek.",
                   "Job numbers and items that were recorded as nonconforming more than once – "
                   "ranked. The sidebar filters apply here too."))
        c1, c2 = st.columns(2)
        rangsor_hu = c1.radio(L("Mi szerint rangsoroljunk?", "Rank by"), list(RANGSOR),
                              format_func=rangsor_cimke, horizontal=True, key="ism_rangsor")
        rendezes = RANGSOR[rangsor_hu]
        rangsor_felirat = rangsor_cimke(rangsor_hu)
        min_db = c2.slider(L("Legalább ennyiszer szerepeljen (a táblázatokban)",
                             "Minimum occurrences (in the tables)"), 2, 10, 2, key="ism_min")

        # --- Top 3 összefoglaló ---
        st.subheader(L(f"🏆 Első 3 – {rangsor_felirat.lower()} szerint",
                       f"🏆 Top 3 – by {rangsor_felirat.lower()}"))
        bal, jobb = st.columns(2)
        with bal:
            if "cikkszám" in szurt.columns:
                top3_doboz(szurt, "cikkszám", L("cikk", "items"), rendezes, rangsor_felirat)
        with jobb:
            if "munkaszám" in szurt.columns:
                top3_doboz(szurt, "munkaszám", L("munkaszám", "job numbers"), rendezes,
                           rangsor_felirat)
        st.divider()

        if "cikkszám" in szurt.columns:
            # 1) Grafikon: a 15 legtöbbször szereplő cikk, hibatípusonként színezve
            cikk_ism = ismetlodok(szurt, ["cikkszám"], min_db, rendezes)
            if len(cikk_ism):
                top = cikk_ism.head(15)["cikkszám"].tolist()
                resz = szurt[szurt["cikkszám"].isin(top)].groupby(["cikkszám", "hibatípus"])
                if rendezes == "esetek":
                    resz = resz.size()
                else:
                    resz = resz["költség" if rendezes == "koltseg" else "nmf_db"].sum()
                bontas = resz.unstack(fill_value=0).reindex(top[::-1])  # a legnagyobb legfelül
                fig = go.Figure()
                for hibatipus in bontas.columns:
                    fig.add_bar(y=bontas.index, x=bontas[hibatipus], name=hibatipus,
                                orientation="h")
                fig.update_layout(barmode="stack", height=150 + 28 * len(top),
                                  title=L(f"Ismétlődő cikkek – top 15 {rangsor_felirat.lower()} "
                                          "szerint, hibatípusonként",
                                          f"Repeat items – top 15 by {rangsor_felirat.lower()}, "
                                          "by defect type"),
                                  xaxis_title=rangsor_felirat, legend=dict(orientation="h", y=-0.2))
                st.plotly_chart(fig, width="stretch")

            # 2) Ugyanaz a cikk UGYANAZZAL a hibával (pl. egy cikk többször kerülethibával)
            st.subheader(L("Ugyanaz a cikk, ugyanazzal a hibával", "Same item, same defect"))
            par = ismetlodok(szurt, ["cikkszám", "hibatípus"], min_db, rendezes)
            st.caption(L(f"{len(par)} cikk–hiba pár szerepel legalább {min_db}-szer.",
                         f"{len(par)} item–defect pairs occur at least {min_db} times."))
            st.dataframe(forditas.oszlopok(par, EN), hide_index=True, width="stretch")

            # 3) Ismétlődő cikkek (bármilyen hibával)
            st.subheader(L("Ismétlődő cikkek (bármilyen hibával)", "Repeat items (any defect)"))
            st.caption(L(f"{len(cikk_ism)} cikk szerepel legalább {min_db}-szer.",
                         f"{len(cikk_ism)} items occur at least {min_db} times."))
            st.dataframe(forditas.oszlopok(cikk_ism, EN), hide_index=True, width="stretch")

        if "munkaszám" in szurt.columns:
            # 4) Ismétlődő munkaszámok (pl. egy munkaszám több hibával)
            st.subheader(L("Ismétlődő munkaszámok", "Repeat job numbers"))
            msz = ismetlodok(szurt, ["munkaszám"], min_db, rendezes)
            st.caption(L(f"{len(msz)} munkaszám szerepel legalább {min_db}-szer.",
                         f"{len(msz)} job numbers occur at least {min_db} times."))
            st.dataframe(forditas.oszlopok(msz, EN), hide_index=True, width="stretch")

# --- 3) p-szabályozókártya ---------------------------------------------------
with tab_p:
    if mennyiseg is None:
        st.info(L(f"A p-kártyához kell a havi gyártott (ellenőrzött) összmennyiség is. "
                  f"Tegyél egy `{MENNYISEG_FAJL}` fájlt a `data/real/` mappába, ezekkel az "
                  "oszlopokkal: `hónap` (pl. 2025-01-01), `gyártott_db`. Az Excel \"Össz db\" "
                  "oszlopa erre nem jó, mert az csak az érintett tételek mennyisége.",
                  f"The p chart also needs the total monthly produced (inspected) quantity. "
                  f"Put a `{MENNYISEG_FAJL}` file into the `data/real/` folder with these "
                  "columns: `hónap` (month, e.g. 2025-01-01), `gyártott_db` (produced qty)."))
    else:
        # FONTOS: a gyártott mennyiség csak havi összesítésben ismert, ezért itt csak
        # a HÓNAP-szűrő érvényes; a többi szűrő (vevő, gép stb.) nem.
        st.caption(L("A p-kártyára csak a hónap-szűrő hat: a középvonal és a határok a "
                     "kiválasztott hónapokból számolódnak. A többi szűrő (vevő, gép stb.) itt "
                     "nem érvényes, mert a gyártott mennyiség csak havi összesítésben ismert.",
                     "Only the month filter applies to the p chart: the centre line and limits "
                     "are calculated from the selected months. Other filters do not apply here, "
                     "because the produced quantity is only known as a monthly total."))
        if valos_adat_kell and not (mappa / MENNYISEG_FAJL).exists():
            kov = " és ".join(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)
            st.caption(L(
                "Gyártott mennyiség = bevételezés "
                f"(`{valos_adat.termeles_fajl(mappa).name}`: jó termékek + selejtpótló "
                f"munkaszámok) + az NMF táblából a {kov} tételek db-ja "
                "(ezek nem kerülnek a bevételezésbe). A javított darabok a javítás után "
                "a bevételezésben szerepelnek, ezért őket nem adjuk hozzá még egyszer.",
                "Produced quantity = receipts "
                f"(`{valos_adat.termeles_fajl(mappa).name}`: good products + scrap "
                "replacement jobs) + scrap and 2nd-class quantities from the NC table "
                "(these are not received into the finished goods store). Repaired parts are "
                "received after repair, so they are not added again."))
        # A belső érték nyelvtől független ("laney" / "klasszikus"), csak a felirat fordított
        TIPUSOK = {"laney": L("Laney p′ (ajánlott)", "Laney p′ (recommended)"),
                   "klasszikus": L("Klasszikus p", "Classic p")}
        tipus = st.radio(
            L("Kártya típusa", "Chart type"), list(TIPUSOK), format_func=lambda k: TIPUSOK[k],
            horizontal=True, key="p_tipus",  # a beszámoló is ezt a választást használja
            help=L("Klasszikus p: a határ csak a darabszámból jön – több ezer darabnál nagyon "
                   "szűk, ezért szinte minden hónap kilóg. Laney p′: a határt a hónapok "
                   "közötti valódi ingadozáshoz igazítja, így csak a tényleg kiugró hónapot "
                   "jelzi.",
                   "Classic p: the limits come from the sample size only – with thousands of "
                   "parts they are very narrow, so almost every month is flagged. Laney p′ "
                   "adjusts the limits to the real month-to-month variation, so it only flags "
                   "truly unusual months."))
        laney = tipus == "laney"

        # Mit számoljunk hibásnak? (csak ha van "Következmény" oszlop, vagyis valós adatnál)
        szamlalo = nm
        mit_cim = L("hibaarány", "defect rate")
        if "következmény" in nm.columns:
            MIT = {"osszes": L("Összes nemmegfelelő (első menetes hibaarány)",
                               "All nonconforming (first-pass defect rate)"),
                   "veszteseg": L("Csak selejt + másodosztály (végleges veszteség)",
                                  "Scrap + 2nd class only (final loss)")}
            mit = st.radio(
                L("Mit számoljunk hibásnak?", "What counts as defective?"), list(MIT),
                format_func=lambda k: MIT[k], horizontal=True,
                key="p_mit",  # a beszámoló is ezt a választást használja
                help=L("Összes: minden NMF darab, a javíthatók is – a gyártás első menetes "
                       "minőségét mutatja. Végleges veszteség: csak ami selejt vagy "
                       "másodosztály lett – ez a valódi anyag- és értékveszteség.",
                       "All: every nonconforming part, including repairable ones – shows "
                       "first-pass quality. Final loss: only parts scrapped or downgraded – "
                       "the real material and value loss."))
            if mit == "veszteseg":
                szamlalo = nm[nm["következmény"].isin(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)]
                mit_cim = L("selejt + másodosztály arány", "scrap + 2nd class rate")

        # A számítás (p, p̄, határok, Laney σz) a szamitasok.py-ban van.
        # A nevező (gyártott db) mindkét esetben ugyanaz, csak a számláló más.
        pk, p_atlag, sigma_z = sz.p_kartya(szamlalo, mennyiseg, tol, ig, laney)

        fig = go.Figure()
        fig.add_scatter(x=pk.index, y=pk["p"] * 100, name=L("Hibaarány (p)", "Defect rate (p)"),
                        mode="lines+markers", line=dict(color="#4C78A8"))
        fig.add_scatter(x=pk.index, y=pk["FSH"] * 100, name=L("FSH (UCL)", "UCL"),
                        line=dict(color="#E45756", dash="dash", shape="hv"))
        fig.add_scatter(x=pk.index, y=pk["ASH"] * 100, name=L("ASH (LCL)", "LCL"),
                        line=dict(color="#E45756", dash="dash", shape="hv"))
        fig.add_hline(y=p_atlag * 100, line_color="green",
                      annotation_text=f"p̄ = {pct(p_atlag, 2)}")
        kivul = pk[pk["kívül"]]
        fig.add_scatter(x=kivul.index, y=kivul["p"] * 100,
                        name=L("Határon kívül", "Out of limits"),
                        mode="markers", marker=dict(color="red", size=12, symbol="x"))
        fig.update_layout(title=L(f"{TIPUSOK[tipus].split(' (')[0]}-szabályozókártya "
                                  f"(havi {mit_cim}, %)",
                                  f"{TIPUSOK[tipus].split(' (')[0]} control chart "
                                  f"(monthly {mit_cim}, %)"),
                          yaxis_title=L("hibaarány (%)", "defect rate (%)"),
                          legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig, width="stretch")
        st.caption(L(f"σz = {forditas.tized(sigma_z, 2, EN)}: a hónapok közötti ingadozás "
                     "ennyiszerese annak, amit a klasszikus p-kártya véletlennek tekint. "
                     + ("1 körüli érték esetén a két kártya gyakorlatilag ugyanaz."
                        if sigma_z < 1.5
                        else "Ilyen nagy értéknél a klasszikus kártya túl sok hamis riasztást ad."),
                     f"σz = {forditas.tized(sigma_z, 2, EN)}: the month-to-month variation is "
                     "this many times what the classic p chart treats as random. "
                     + ("With a value around 1 the two charts are practically the same."
                        if sigma_z < 1.5
                        else "At such a high value the classic chart gives too many false "
                             "alarms.")))

        if len(kivul):
            st.warning(L(f"{len(kivul)} hónap esik a szabályozási határokon kívül – "
                         "ez speciális (nem véletlen) okra utalhat. Tartós javuló trendnél "
                         "érdemes a határokat az új időszakra újraszámolni.",
                         f"{len(kivul)} month(s) fall outside the control limits – this may "
                         "indicate a special (non-random) cause. With a lasting improving "
                         "trend, recalculate the limits for the new period."))
        else:
            st.success(L("Minden hónap a szabályozási határokon belül van.",
                         "All months are within the control limits."))
        with st.expander(L("Számított táblázat", "Calculated table")):
            st.dataframe(forditas.oszlopok(
                pk.assign(p=pk["p"].round(4), FSH=pk["FSH"].round(4), ASH=pk["ASH"].round(4)),
                EN))

# --- 4) AI-ellenőrzési napló -------------------------------------------------
with tab_ai:
    if DEMO:
        st.write(L("Itt rögzítheted, mit ellenőriztél kézzel egy AI által adott kimeneten. "
                   "**Online bemutató:** a bejegyzések csak a te böngészőablakodban "
                   "látszanak, és az oldal bezárásakor törlődnek – mások nem látják. "
                   "Megtartáshoz használd a „Napló letöltése” gombot.",
                   "Record here what you checked manually in an AI-generated output. "
                   "**Online demo:** entries are visible only in your own browser window and "
                   "are deleted when you close the page – others cannot see them. "
                   "Use the “Download log” button to keep them."))
    else:
        st.write(L("Itt rögzítheted, mit ellenőriztél kézzel egy AI által adott kimeneten. "
                   f"A bejegyzések helyben, a `{AI_NAPLO_FAJL.relative_to(Path(__file__).parent)}` "
                   "fájlba mentődnek.",
                   "Record here what you checked manually in an AI-generated output. "
                   "Entries are saved locally to "
                   f"`{AI_NAPLO_FAJL.relative_to(Path(__file__).parent)}`."))

    # A st.form miatt csak a "Mentés" gombra kattintva fut le a mentés
    with st.form("ai_naplo_urlap", clear_on_submit=True):
        c1, c2 = st.columns(2)
        datum = c1.date_input(L("Dátum", "Date"), value=date.today())
        eszkoz = c2.text_input(L("AI eszköz", "AI tool"),
                               placeholder=L("pl. Claude, ChatGPT, Copilot",
                                             "e.g. Claude, ChatGPT, Copilot"))
        feladat = st.text_area(L("Feladat / AI kimenet rövid leírása",
                                 "Task / short description of the AI output"),
                               placeholder=L("pl. Pareto-elemzés a 2025. Q3 adatokra",
                                             "e.g. Pareto analysis of 2025 Q3 data"))
        szempontok = st.multiselect(
            L("Mit ellenőriztem?", "What did I check?"),
            [L("Számok, számítások", "Numbers, calculations"),
             L("Forrásadat helyessége", "Correctness of source data"),
             L("Logika, következtetések", "Logic, conclusions"),
             L("Képletek / kód", "Formulas / code"),
             L("Mértékegységek", "Units"),
             L("Szabványnak / előírásnak megfelelés", "Compliance with standards / specs"),
             L("Hiányzó információ", "Missing information"),
             L("Adatvédelem (nincs érzékeny adat)", "Data protection (no sensitive data)")],
            placeholder=L("Válassz", "Choose"),
        )
        c3, c4 = st.columns(2)
        eredmeny = c3.selectbox(L("Eredmény", "Result"),
                                [L("Rendben, elfogadva", "OK, accepted"),
                                 L("Javítva elfogadva", "Accepted after correction"),
                                 L("Elvetve", "Rejected")])
        megjegyzes = c4.text_input(L("Megjegyzés", "Note"),
                                   placeholder=L("pl. két hónap összege hibás volt",
                                                 "e.g. two monthly totals were wrong"))
        mentes = st.form_submit_button(L("Mentés", "Save"))

    if mentes:
        if not feladat.strip():
            st.error(L("A feladat leírása kötelező.", "The task description is required."))
        else:
            uj_sor = pd.DataFrame([{
                "dátum": datum.isoformat(),
                "AI eszköz": eszkoz,
                "feladat / AI kimenet": feladat,
                "ellenőrzött szempontok": "; ".join(szempontok),
                "eredmény": eredmeny,
                "megjegyzés": megjegyzes,
            }])
            naplo = pd.concat([ai_naplo_betoltes(), uj_sor], ignore_index=True)
            if DEMO:
                st.session_state["ai_naplo"] = naplo  # csak ebben a böngészőmenetben
            else:
                AI_NAPLO_FAJL.parent.mkdir(parents=True, exist_ok=True)
                naplo.to_csv(AI_NAPLO_FAJL, index=False, encoding="utf-8-sig")
            st.success(L("Bejegyzés elmentve.", "Entry saved."))

    naplo = ai_naplo_betoltes()
    st.subheader(L(f"Eddigi bejegyzések ({len(naplo)})", f"Entries so far ({len(naplo)})"))
    st.dataframe(forditas.oszlopok(naplo, EN), width="stretch", hide_index=True)
    if len(naplo):
        # A letöltés a saját böngésződbe ment fájlt, nem tölt fel semmit
        st.download_button(L("Napló letöltése (CSV)", "Download log (CSV)"),
                           naplo.to_csv(index=False).encode("utf-8-sig"),
                           file_name="ai_naplo.csv", mime="text/csv")

# --- Beszámoló (PowerPoint + PDF) az oldalsávban -----------------------------
# A beszámoló a gomb megnyomásakor érvényes szűrőkkel készül. A fájl a saját
# gépeden jön létre, a böngésző letöltésként menti – semmit nem küld sehova.
st.sidebar.divider()
st.sidebar.subheader(L("📄 Beszámoló", "📄 Report"))
laney_besz = st.session_state.get("p_tipus", "laney") == "laney"
# "Ujjlenyomat": ha bármelyik beállítás változik, a korábban elkészült fájl elavult
allapot = (valos_adat_kell, str(valos_fajl), str(tol), str(ig), laney_besz, EN,
           tuple((k, tuple(v)) for k, v in aktiv_szurok.items()), verzio)

if st.sidebar.button(L("Beszámoló készítése", "Create report"), type="primary"):
    with st.spinner(L("Beszámoló készítése...", "Creating report...")):
        adat = beszamolo.adat_osszeallitas(
            szurt, nm, mennyiseg, tol, ig, laney_besz,
            szurok=[f"{k}: {', '.join(map(str, v))}"[:120] for k, v in aktiv_szurok.items()],
            forras=(L(f"valós adat ({valos_fajl.name})", f"real data ({valos_fajl.name})")
                    if valos_adat_kell else L("szintetikus próbaadat", "synthetic sample data")),
            cimkek=BONTASOK,
            en=EN,
        )
        st.session_state["beszamolo"] = {"allapot": allapot,
                                         "pptx": beszamolo.pptx_keszit(adat),
                                         "pdf": beszamolo.pdf_keszit(adat)}

kesz = st.session_state.get("beszamolo")
if kesz and kesz["allapot"] == allapot:
    fajlnev = (L("minosegugyi_beszamolo", "quality_report")
               + f"_{pd.Timestamp(tol):%Y%m}-{pd.Timestamp(ig):%Y%m}")
    st.sidebar.download_button("⬇ PowerPoint (.pptx)", kesz["pptx"], file_name=f"{fajlnev}.pptx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "presentationml.presentation")
    st.sidebar.download_button("⬇ PDF", kesz["pdf"], file_name=f"{fajlnev}.pdf",
                               mime="application/pdf")
    if valos_adat_kell:
        st.sidebar.caption(L("⚠ Valós adatokat tartalmaz (vevőnevek, költségek) – "
                             "megosztás előtt nézd át, kinek küldöd.",
                             "⚠ Contains real data (customer names, costs) – check who you "
                             "send it to before sharing."))
elif kesz:
    st.sidebar.caption(L("A szűrők változtak – készítsd el újra a beszámolót.",
                         "The filters have changed – create the report again."))
else:
    st.sidebar.caption(L("A beszámoló a most beállított szűrőkkel készül.",
                         "The report uses the filters currently set."))
