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
#   3) p-szabályozókártya    – a havi hibaarány statisztikailag stabil-e
#   4) AI-ellenőrzési napló  – kézzel rögzítheted, mit ellenőriztél az AI kimenetén
#
# Adatforrás: alapból a data/synthetic/ (kitalált) adat. A data/real/ (valós)
# adatot csak az oldalsávban lévő kapcsolóval lehet bekapcsolni.
# -----------------------------------------------------------------------------

from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

import beszamolo  # a saját fájlunk: PowerPoint / PDF beszámoló készítése
import szamitasok as sz  # a saját fájlunk: közös számítások (képernyő + beszámoló)
import valos_adat  # a saját fájlunk: a valós Excel beolvasása

# --- Mappák és fájlnevek -----------------------------------------------------
ALAP_MAPPA = Path(__file__).parent / "data"
SZINTETIKUS_MAPPA = ALAP_MAPPA / "synthetic"
VALOS_MAPPA = ALAP_MAPPA / "real"
NM_FAJL = "nemmegfelelosegek.csv"          # nemmegfelelőségek listája
MENNYISEG_FAJL = "gyartott_mennyiseg.csv"  # havi gyártott darabszám
AI_NAPLO_FAJL = ALAP_MAPPA / "ai_naplo.csv"  # az AI-ellenőrzési napló (helyi fájl)

KOTELEZO_OSZLOPOK = ["dátum", "hibatípus", "ok", "termékcsalád", "költség", "státusz"]
AI_NAPLO_OSZLOPOK = ["dátum", "AI eszköz", "feladat / AI kimenet", "ellenőrzött szempontok",
                     "eredmény", "megjegyzés"]

st.set_page_config(page_title="Minőségügyi dashboard", layout="wide")


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
    """Beolvassa az AI-ellenőrzési naplót; ha még nincs, üres táblát ad."""
    if AI_NAPLO_FAJL.exists():
        return pd.read_csv(AI_NAPLO_FAJL, encoding="utf-8-sig")
    return pd.DataFrame(columns=AI_NAPLO_OSZLOPOK)


# --- Oldalsáv: adatforrás és szűrők ------------------------------------------
st.sidebar.header("Beállítások")
valos_adat_kell = st.sidebar.toggle(
    "Valós adat használata (data/real/)",
    value=False,  # alapból KI van kapcsolva -> szintetikus adat
    help="Kikapcsolva a kitalált (szintetikus) adatot látod.",
)
mappa = VALOS_MAPPA if valos_adat_kell else SZINTETIKUS_MAPPA

valos_fajl = None
if valos_adat_kell:
    fajlok = valos_adat.valos_fajlok(VALOS_MAPPA)
    if not fajlok:
        st.warning("Nem található Excel-fájl a `data/real/` mappában. "
                   "Másold oda a nemmegfelelőségi táblát (.xlsx).")
        st.stop()
    valos_fajl = st.sidebar.selectbox("Valós fájl", fajlok, format_func=lambda p: p.name)
elif not (mappa / NM_FAJL).exists():
    st.warning("Még nincs szintetikus adat. Futtasd le: `python generate_data.py`")
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
    st.error(f"A {NM_FAJL} fájlból hiányoznak ezek az oszlopok: {', '.join(hianyzo)}")
    st.stop()

# --- Szűrők ------------------------------------------------------------------
nm["hónap"] = nm["dátum"].dt.to_period("M").dt.to_timestamp()

# belső oszlopnév -> felirat (csak a létező oszlopok jelennek meg).
# Ugyanezt a listát használja a Pareto-diagram bontás-választója is.
BONTASOK = {
    "termékcsalád": "Megrendelő (Vevő)" if valos_adat_kell else "Termékcsalád",
    "hibatípus": "Nemmegfelelőség jellege" if valos_adat_kell else "Hibatípus",
    "ok": "Nemmegfelelőség oka" if valos_adat_kell else "Ok",
    "keletkezés_helye": "Keletkezés helye (művelet)",
    "gép": "NMF keletkezés helye (gép)",
    "következmény": "Következmény",
}
BONTASOK = {o: f for o, f in BONTASOK.items() if o in nm.columns}
SZURO_OSZLOPOK = [o for o in BONTASOK if o != "ok"]  # az okra nem kértünk szűrőt

honapok = list(pd.DatetimeIndex(nm["hónap"].unique()).sort_values())


def szurok_torlese(teljes_tartomany):
    """A 'Szűrők törlése' gomb: a hónap a teljes időszak lesz, a többi 'Mind'."""
    st.session_state["szuro_hónap"] = teljes_tartomany
    for oszlop in SZURO_OSZLOPOK:
        st.session_state[f"szuro_{oszlop}"] = []


st.sidebar.subheader("Szűrők")
szurt = nm.copy()

# Hónap: csúszka, két végével adod meg a "-tól" és "-ig" hónapot
tol, ig = honapok[0], honapok[-1]  # alapból a teljes időszak (a p-kártya is ezt használja)
if len(honapok) > 1:
    tol, ig = st.sidebar.select_slider(
        "Hónap (-tól, -ig)", options=honapok, value=(honapok[0], honapok[-1]),
        format_func=lambda d: pd.Timestamp(d).strftime("%Y. %m."), key="szuro_hónap")
    szurt = szurt[(szurt["hónap"] >= tol) & (szurt["hónap"] <= ig)]

# A többi szűrő üresen indul = "Mind". Ha kiválasztasz értékeket, csak azok maradnak.
aktiv_szurok = {}  # a beszámoló címlapjára: melyik szűrő mire van állítva
for oszlop in SZURO_OSZLOPOK:
    valasztott = st.sidebar.multiselect(BONTASOK[oszlop], sorted(nm[oszlop].unique()),
                                        key=f"szuro_{oszlop}", placeholder="Mind")
    if valasztott:
        szurt = szurt[szurt[oszlop].isin(valasztott)]
        aktiv_szurok[BONTASOK[oszlop]] = valasztott
st.sidebar.button("Szűrők törlése", on_click=szurok_torlese,
                  args=((honapok[0], honapok[-1]),))
st.sidebar.caption(f"Szűrés után {len(szurt)} / {len(nm)} sor látszik.")

st.sidebar.caption("Adatforrás: " + ("🔴 VALÓS (data/real/)" if valos_adat_kell
                                      else "🟢 szintetikus (data/synthetic/)"))

# --- Fejléc és kulcsszámok ---------------------------------------------------
st.title("Minőségügyi dashboard – nemmegfelelőségek")
if hibas_sorok:
    st.warning(f"{len(hibas_sorok)} sor kimaradt, mert a dátuma hiányzik vagy hibás "
               f"(pl. 1900-as év). Javítsd ezeket az Excelben – sorok: "
               + ", ".join(map(str, hibas_sorok)))
if kozos is not None:
    kiesett, (elso, utolso) = kozos
    szoveg = (f"Vizsgált időszak: {elso:%Y.%m.%d.} – {utolso:%Y.%m.%d.} "
              "(az NMF tábla és a bevételezési CSV közös időszaka).")
    if kiesett:
        st.warning(szoveg + f" {kiesett} NMF-sor kimaradt, mert a dátuma ezen kívül esik.")
    else:
        st.caption(szoveg + " Minden NMF-sor ebbe az időszakba esik.")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Nemmegfelelőségek száma", f"{len(szurt):,}".replace(",", " "))
k2.metric("Összköltség", f"{szurt['költség'].sum():,.0f} Ft".replace(",", " "))
k3.metric("Nyitott / folyamatban", int((szurt["státusz"].str.startswith("Nyitott")
                                        | (szurt["státusz"] == "Folyamatban")).sum()))
k4.metric("Átlagköltség / eset", f"{szurt['költség'].mean():,.0f} Ft".replace(",", " ")
          if len(szurt) else "–")

tab_trend, tab_pareto, tab_ism, tab_p, tab_ai = st.tabs(
    ["📈 Havi trend", "📊 Pareto-diagram", "🔁 Ismétlődések", "🎯 p-szabályozókártya",
     "📝 AI-ellenőrzési napló"]
)

# --- 1) Havi trend -----------------------------------------------------------
with tab_trend:
    havi = sz.havi_osszesites(szurt)  # havi esetszám, költség és lineáris trend
    trend = havi["trend"]

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=havi["hónap"], y=havi["darab"], name="Nemmegfelelőségek (db)",
                marker_color="#4C78A8")
    fig.add_scatter(x=havi["hónap"], y=trend, name="Trend (db)", mode="lines",
                    line=dict(color="#4C78A8", dash="dash"))
    fig.add_scatter(x=havi["hónap"], y=havi["koltseg"], name="Költség (Ft)",
                    mode="lines+markers", line=dict(color="#E45756"), secondary_y=True)
    fig.update_yaxes(title_text="darab", secondary_y=False)
    fig.update_yaxes(title_text="Ft", secondary_y=True)
    fig.update_layout(title="Havi nemmegfelelőségek és költség", hovermode="x unified",
                      legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, width="stretch")
    st.caption("A szaggatott vonal a lineáris trend: ha lefelé tart, javul a helyzet.")

# --- 2) Pareto-diagram -------------------------------------------------------
with tab_pareto:
    c1, c2 = st.columns(2)
    bontas = c1.selectbox("Mi szerint bontsuk?", list(BONTASOK),
                          index=list(BONTASOK).index("hibatípus"),
                          format_func=lambda o: BONTASOK[o])
    mertek = c2.radio("Mi alapján rangsoroljunk?", ["Darabszám", "Költség"], horizontal=True)

    # Pareto-sorrend; 15-nél több kategóriánál a többi "Egyéb"-be kerül
    p_adat = sz.pareto_adat(szurt, bontas, mertek)

    if len(p_adat) < 2:
        st.info(f"A szűrés után csak {len(p_adat)} érték maradt ebben a bontásban, így nincs "
                "mit rangsorolni. Válassz másik bontást (pl. a hiba okát), vagy lazíts a "
                "szűrőkön.")
    else:
        kumulalt = p_adat.cumsum() / p_adat.sum() * 100  # kumulált százalék

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=p_adat.index, y=p_adat.values, name=mertek, marker_color="#4C78A8")
        fig.add_scatter(x=p_adat.index, y=kumulalt.values, name="Kumulált %",
                        mode="lines+markers", line=dict(color="#E45756"), secondary_y=True)
        fig.add_hline(y=80, line_dash="dot", line_color="gray", secondary_y=True,
                      annotation_text="80%")
        fig.update_yaxes(title_text=mertek, secondary_y=False)
        fig.update_yaxes(title_text="Kumulált %", range=[0, 105], tickvals=[0, 20, 40, 60, 80, 100],
                         showgrid=False, secondary_y=True)
        fig.update_layout(title=f"Pareto-diagram – {BONTASOK[bontas].lower()} szerint "
                                f"({mertek.lower()})",
                          legend=dict(orientation="h", y=-0.35))
        st.plotly_chart(fig, width="stretch")
        st.caption("80/20 szabály: a 80%-os vonal alatti néhány tételre érdemes először "
                   "fókuszálni. Tipp: szűrj egy hibatípusra, és bonts ok szerint.")

# --- Ismétlődések: többször felvett munkaszámok és cikkszámok ----------------
# A számítások a szamitasok.py-ban vannak (a beszámoló is ezeket használja)
RANGSOR = sz.RANGSOR
ismetlodok = sz.ismetlodok


def top3_doboz(df: pd.DataFrame, kulcs: str, cim: str, rendezes: str, felirat: str):
    """Megmutatja az első 3 tételt, és hogy együtt hány %-át adják az egésznek.
    Itt MINDEN tétel számít (az egyszer szereplők is), különben a % félrevezető lenne."""
    ertek = sz.rangsor_ertek(df, kulcs, rendezes)
    if ertek.sum() <= 0:
        st.info(f"Nincs adat ({cim}).")
        return
    reszesedes = ertek / ertek.sum()
    kell_80 = sz.hany_adja_a_80_szazalekot(ertek)  # hány tétel adja a 80%-ot

    st.metric(f"Első 3 {cim} részesedése", f"{reszesedes.head(3).sum():.0%}",
              help=f"A szűrt adatok teljes {felirat.lower()} értékéből ennyit ad az első 3.")
    top = pd.DataFrame({
        "#": range(1, min(3, len(ertek)) + 1),
        kulcs: ertek.index[:3],
        felirat: ertek.values[:3],
        "részesedés": [f"{x:.1%}" for x in reszesedes.values[:3]],
        "kumulált": [f"{x:.1%}" for x in reszesedes.cumsum().values[:3]],
    })
    st.dataframe(top, hide_index=True, width="stretch")
    st.caption(f"A 80%-ot az első **{kell_80}** {cim} adja (összesen {len(ertek)} van).")


with tab_ism:
    if not {"munkaszám", "cikkszám"} & set(szurt.columns):
        st.info("Ehhez az adathoz nincs munkaszám / cikkszám oszlop.")
    else:
        st.write("Azok a munkaszámok és cikkek, amelyek többször is bekerültek a "
                 "nemmegfelelőségek közé – rangsorolva. Az oldalsáv szűrői itt is érvényesek.")
        c1, c2 = st.columns(2)
        rangsor_felirat = c1.radio("Mi szerint rangsoroljunk?", list(RANGSOR), horizontal=True)
        rendezes = RANGSOR[rangsor_felirat]
        min_db = c2.slider("Legalább ennyiszer szerepeljen (a táblázatokban)", 2, 10, 2)

        # --- Top 3 összefoglaló ---
        st.subheader(f"🏆 Első 3 – {rangsor_felirat.lower()} szerint")
        bal, jobb = st.columns(2)
        with bal:
            if "cikkszám" in szurt.columns:
                top3_doboz(szurt, "cikkszám", "cikk", rendezes, rangsor_felirat)
        with jobb:
            if "munkaszám" in szurt.columns:
                top3_doboz(szurt, "munkaszám", "munkaszám", rendezes, rangsor_felirat)
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
                                  title=f"Ismétlődő cikkek – top 15 {rangsor_felirat.lower()} "
                                        "szerint, hibatípusonként",
                                  xaxis_title=rangsor_felirat, legend=dict(orientation="h", y=-0.2))
                st.plotly_chart(fig, width="stretch")

            # 2) Ugyanaz a cikk UGYANAZZAL a hibával (pl. egy cikk többször kerülethibával)
            st.subheader("Ugyanaz a cikk, ugyanazzal a hibával")
            par = ismetlodok(szurt, ["cikkszám", "hibatípus"], min_db, rendezes)
            st.caption(f"{len(par)} cikk–hiba pár szerepel legalább {min_db}-szer.")
            st.dataframe(par, hide_index=True, width="stretch")

            # 3) Ismétlődő cikkek (bármilyen hibával)
            st.subheader("Ismétlődő cikkek (bármilyen hibával)")
            st.caption(f"{len(cikk_ism)} cikk szerepel legalább {min_db}-szer.")
            st.dataframe(cikk_ism, hide_index=True, width="stretch")

        if "munkaszám" in szurt.columns:
            # 4) Ismétlődő munkaszámok (pl. egy munkaszám több hibával)
            st.subheader("Ismétlődő munkaszámok")
            msz = ismetlodok(szurt, ["munkaszám"], min_db, rendezes)
            st.caption(f"{len(msz)} munkaszám szerepel legalább {min_db}-szer.")
            st.dataframe(msz, hide_index=True, width="stretch")

# --- 3) p-szabályozókártya ---------------------------------------------------
with tab_p:
    if mennyiseg is None:
        st.info(f"A p-kártyához kell a havi gyártott (ellenőrzött) összmennyiség is. "
                f"Tegyél egy `{MENNYISEG_FAJL}` fájlt a `data/real/` mappába, ezekkel az "
                "oszlopokkal: `hónap` (pl. 2025-01-01), `gyártott_db`. Az Excel \"Össz db\" "
                "oszlopa erre nem jó, mert az csak az érintett tételek mennyisége.")
    else:
        # FONTOS: a gyártott mennyiség csak havi összesítésben ismert, ezért itt csak
        # a HÓNAP-szűrő érvényes; a többi szűrő (vevő, gép stb.) nem.
        st.caption("A p-kártyára csak a hónap-szűrő hat: a középvonal és a határok a "
                   "kiválasztott hónapokból számolódnak. A többi szűrő (vevő, gép stb.) itt "
                   "nem érvényes, mert a gyártott mennyiség csak havi összesítésben ismert.")
        if valos_adat_kell and not (mappa / MENNYISEG_FAJL).exists():
            st.caption(
                "Gyártott mennyiség = bevételezés "
                f"(`{valos_adat.termeles_fajl(mappa).name}`: jó termékek + selejtpótló "
                "munkaszámok) + az NMF táblából a "
                f"{' és '.join(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)} tételek db-ja "
                "(ezek nem kerülnek a bevételezésbe). A javított darabok a javítás után "
                "a bevételezésben szerepelnek, ezért őket nem adjuk hozzá még egyszer.")
        tipus = st.radio(
            "Kártya típusa", ["Laney p′ (ajánlott)", "Klasszikus p"], horizontal=True,
            key="p_tipus",  # a beszámoló is ezt a választást használja
            help="Klasszikus p: a határ csak a darabszámból jön – több ezer darabnál nagyon "
                 "szűk, ezért szinte minden hónap kilóg. Laney p′: a határt a hónapok "
                 "közötti valódi ingadozáshoz igazítja, így csak a tényleg kiugró hónapot jelzi.")
        laney = tipus.startswith("Laney")

        # Mit számoljunk hibásnak? (csak ha van "Következmény" oszlop, vagyis valós adatnál)
        szamlalo = nm
        mit_cim = "hibaarány"
        if "következmény" in nm.columns:
            mit = st.radio(
                "Mit számoljunk hibásnak?",
                ["Összes nemmegfelelő (első menetes hibaarány)",
                 "Csak selejt + másodosztály (végleges veszteség)"],
                horizontal=True, key="p_mit",  # a beszámoló is ezt a választást használja
                help="Összes: minden NMF darab, a javíthatók is – a gyártás első menetes "
                     "minőségét mutatja. Végleges veszteség: csak ami selejt vagy "
                     "másodosztály lett – ez a valódi anyag- és értékveszteség.")
            if mit.startswith("Csak"):
                szamlalo = nm[nm["következmény"].isin(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)]
                mit_cim = "selejt + másodosztály arány"

        # A számítás (p, p̄, határok, Laney σz) a szamitasok.py-ban van.
        # A nevező (gyártott db) mindkét esetben ugyanaz, csak a számláló más.
        pk, p_atlag, sigma_z = sz.p_kartya(szamlalo, mennyiseg, tol, ig, laney)

        fig = go.Figure()
        fig.add_scatter(x=pk.index, y=pk["p"] * 100, name="Hibaarány (p)",
                        mode="lines+markers", line=dict(color="#4C78A8"))
        fig.add_scatter(x=pk.index, y=pk["FSH"] * 100, name="FSH (UCL)",
                        line=dict(color="#E45756", dash="dash", shape="hv"))
        fig.add_scatter(x=pk.index, y=pk["ASH"] * 100, name="ASH (LCL)",
                        line=dict(color="#E45756", dash="dash", shape="hv"))
        fig.add_hline(y=p_atlag * 100, line_color="green",
                      annotation_text=f"p̄ = {p_atlag:.2%}")
        kivul = pk[pk["kívül"]]
        fig.add_scatter(x=kivul.index, y=kivul["p"] * 100, name="Határon kívül",
                        mode="markers", marker=dict(color="red", size=12, symbol="x"))
        fig.update_layout(title=f"{'Laney p′' if laney else 'Klasszikus p'}-szabályozókártya "
                                f"(havi {mit_cim}, %)",
                          yaxis_title="hibaarány (%)", legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig, width="stretch")
        st.caption(f"σz = {sigma_z:.2f}: a hónapok közötti ingadozás ennyiszerese annak, amit "
                   "a klasszikus p-kártya véletlennek tekint. "
                   + ("1 körüli érték esetén a két kártya gyakorlatilag ugyanaz." if sigma_z < 1.5
                      else "Ilyen nagy értéknél a klasszikus kártya túl sok hamis riasztást ad."))

        if len(kivul):
            st.warning(f"{len(kivul)} hónap esik a szabályozási határokon kívül – "
                       "ez speciális (nem véletlen) okra utalhat. Tartós javuló trendnél "
                       "érdemes a határokat az új időszakra újraszámolni.")
        else:
            st.success("Minden hónap a szabályozási határokon belül van.")
        with st.expander("Számított táblázat"):
            st.dataframe(pk.assign(p=pk["p"].round(4), FSH=pk["FSH"].round(4),
                                   ASH=pk["ASH"].round(4)))

# --- 4) AI-ellenőrzési napló -------------------------------------------------
with tab_ai:
    st.write("Itt rögzítheted, mit ellenőriztél kézzel egy AI által adott kimeneten. "
             f"A bejegyzések helyben, a `{AI_NAPLO_FAJL.relative_to(Path(__file__).parent)}` "
             "fájlba mentődnek.")

    # A st.form miatt csak a "Mentés" gombra kattintva fut le a mentés
    with st.form("ai_naplo_urlap", clear_on_submit=True):
        c1, c2 = st.columns(2)
        datum = c1.date_input("Dátum", value=date.today())
        eszkoz = c2.text_input("AI eszköz", placeholder="pl. Claude, ChatGPT, Copilot")
        feladat = st.text_area("Feladat / AI kimenet rövid leírása",
                               placeholder="pl. Pareto-elemzés a 2025. Q3 adatokra")
        szempontok = st.multiselect(
            "Mit ellenőriztem?",
            ["Számok, számítások", "Forrásadat helyessége", "Logika, következtetések",
             "Képletek / kód", "Mértékegységek", "Szabványnak / előírásnak megfelelés",
             "Hiányzó információ", "Adatvédelem (nincs érzékeny adat)"],
        )
        c3, c4 = st.columns(2)
        eredmeny = c3.selectbox("Eredmény", ["Rendben, elfogadva", "Javítva elfogadva",
                                             "Elvetve"])
        megjegyzes = c4.text_input("Megjegyzés", placeholder="pl. két hónap összege hibás volt")
        mentes = st.form_submit_button("Mentés")

    if mentes:
        if not feladat.strip():
            st.error("A feladat leírása kötelező.")
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
            AI_NAPLO_FAJL.parent.mkdir(parents=True, exist_ok=True)
            naplo.to_csv(AI_NAPLO_FAJL, index=False, encoding="utf-8-sig")
            st.success("Bejegyzés elmentve.")

    naplo = ai_naplo_betoltes()
    st.subheader(f"Eddigi bejegyzések ({len(naplo)})")
    st.dataframe(naplo, width="stretch", hide_index=True)
    if len(naplo):
        # A letöltés a saját böngésződbe ment fájlt, nem tölt fel semmit
        st.download_button("Napló letöltése (CSV)",
                           naplo.to_csv(index=False).encode("utf-8-sig"),
                           file_name="ai_naplo.csv", mime="text/csv")

# --- Beszámoló (PowerPoint + PDF) az oldalsávban -----------------------------
# A beszámoló a gomb megnyomásakor érvényes szűrőkkel készül. A fájl a saját
# gépeden jön létre, a böngésző letöltésként menti – semmit nem küld sehova.
st.sidebar.divider()
st.sidebar.subheader("📄 Beszámoló")
laney_besz = st.session_state.get("p_tipus", "Laney").startswith("Laney")
# "Ujjlenyomat": ha bármelyik beállítás változik, a korábban elkészült fájl elavult
allapot = (valos_adat_kell, str(valos_fajl), str(tol), str(ig), laney_besz,
           tuple((k, tuple(v)) for k, v in aktiv_szurok.items()), verzio)

if st.sidebar.button("Beszámoló készítése", type="primary"):
    with st.spinner("Beszámoló készítése..."):
        adat = beszamolo.adat_osszeallitas(
            szurt, nm, mennyiseg, tol, ig, laney_besz,
            szurok=[f"{k}: {', '.join(map(str, v))}"[:120] for k, v in aktiv_szurok.items()],
            forras=(f"valós adat ({valos_fajl.name})" if valos_adat_kell
                    else "szintetikus próbaadat"),
            cimkek=BONTASOK,
        )
        st.session_state["beszamolo"] = {"allapot": allapot,
                                         "pptx": beszamolo.pptx_keszit(adat),
                                         "pdf": beszamolo.pdf_keszit(adat)}

kesz = st.session_state.get("beszamolo")
if kesz and kesz["allapot"] == allapot:
    fajlnev = f"minosegugyi_beszamolo_{pd.Timestamp(tol):%Y%m}-{pd.Timestamp(ig):%Y%m}"
    st.sidebar.download_button("⬇ PowerPoint (.pptx)", kesz["pptx"], file_name=f"{fajlnev}.pptx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "presentationml.presentation")
    st.sidebar.download_button("⬇ PDF", kesz["pdf"], file_name=f"{fajlnev}.pdf",
                               mime="application/pdf")
    if valos_adat_kell:
        st.sidebar.caption("⚠ Valós adatokat tartalmaz (vevőnevek, költségek) – "
                           "megosztás előtt nézd át, kinek küldöd.")
elif kesz:
    st.sidebar.caption("A szűrők változtak – készítsd el újra a beszámolót.")
else:
    st.sidebar.caption("A beszámoló a most beállított szűrőkkel készül.")
