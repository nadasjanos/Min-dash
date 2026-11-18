# -----------------------------------------------------------------------------
# beszamolo.py
# EZ A FÁJL: a dashboard adataiból prezentálható beszámolót készít:
#   - PowerPoint (.pptx) – szerkeszthető, értekezletre
#   - PDF – ugyanaz a tartalom, archiváláshoz / e-mailhez
# A fájl a saját gépeden készül el (a böngésző letöltésként menti), semmit
# nem küld sehova. A grafikonokat matplotlib rajzolja (internet nem kell hozzá).
#
# Diák / oldalak:
#   1. Címlap   2. Kulcsszámok   3. Havi trend   4. Pareto – jelleg
#   5. Pareto – ok   6. Első 3 (becsült költség)   7. Ismétlődő cikk–hiba párok
#   8. p-szabályozókártya   9. Intézkedések (üres táblázat, kézzel kitöltendő)
# -----------------------------------------------------------------------------

import io
from datetime import date

import matplotlib

matplotlib.use("Agg")  # képernyő nélküli rajzolás (csak fájlba)
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

import szamitasok as sz  # noqa: E402

KEK, PIROS, ZOLD, SZURKE = "#4C78A8", "#E45756", "#54A24B", "#6B6B6B"


# --- Segédfüggvények ------------------------------------------------------------
def ft(x) -> str:
    """Pénzösszeg magyar formában: 1234567 -> '1 234 567 Ft'."""
    return f"{x:,.0f} Ft".replace(",", " ")


def szam(x) -> str:
    """Egész szám ezres tagolással: 1234 -> '1 234'."""
    return f"{x:,.0f}".replace(",", " ")


def pct(x, tizedes=1) -> str:
    """Százalék magyar formában (tizedesvessző): 0.0259 -> '2,59%'."""
    return f"{x:.{tizedes}%}".replace(".", ",")


def tized(x, tizedes=1) -> str:
    """Tizedes tört magyar formában: 1.25 -> '1,25'."""
    return f"{x:.{tizedes}f}".replace(".", ",")


def rovid(szoveg, hossz=30) -> str:
    """Hosszú szöveg levágása, hogy elférjen a táblázatban / tengelyen."""
    szoveg = str(szoveg)
    return szoveg if len(szoveg) <= hossz else szoveg[: hossz - 1] + "…"


# --- 1) Adatok összeállítása ---------------------------------------------------------
def adat_osszeallitas(szurt, nm, mennyiseg, tol, ig, laney, szurok, forras, cimkek) -> dict:
    """Minden számot és táblázatot kiszámol, ami a beszámolóba kerül.
    szurok: a beállított szűrők szöveges leírása (lista)
    cimkek: belső oszlopnév -> felirat (pl. 'hibatípus' -> 'Nemmegfelelőség jellege')"""
    a = {
        "idoszak": f"{pd.Timestamp(tol):%Y.%m.} – {pd.Timestamp(ig):%Y.%m.}",
        "szurok": szurok or ["nincs (minden adat)"],
        "forras": forras,
        "keszult": date.today().strftime("%Y.%m.%d"),
        "cimkek": cimkek,
        "esetek": len(szurt),
        "koltseg": szurt["költség"].sum(),
        "nyitott": int((szurt["státusz"].str.startswith("Nyitott")
                        | (szurt["státusz"] == "Folyamatban")).sum()),
        "hibas_db": szurt["nmf_db"].sum(),
        "atlag": szurt["költség"].mean() if len(szurt) else 0,
    }

    # Havi trend és meredekség (eset / hónap)
    a["havi"] = sz.havi_osszesites(szurt) if len(szurt) else None
    if a["havi"] is not None and len(a["havi"]) > 1:
        a["meredekseg"] = np.polyfit(np.arange(len(a["havi"])), a["havi"]["darab"], 1)[0]
    else:
        a["meredekseg"] = None

    # Pareto: jelleg és ok szerint (darabszám)
    for kulcs in ["hibatípus", "ok"]:
        teljes = szurt[kulcs].value_counts()
        a[f"pareto_{kulcs}"] = sz.pareto_adat(szurt, kulcs, "Darabszám", max_oszlop=12)
        a[f"pareto_{kulcs}_n80"] = sz.hany_adja_a_80_szazalekot(teljes) if len(teljes) else 0
        a[f"pareto_{kulcs}_osszes"] = len(teljes)

    # Első 3 cikk / munkaszám becsült költség szerint (minden tétel számít)
    for kulcs in ["cikkszám", "munkaszám"]:
        if kulcs in szurt.columns:
            ertek = sz.rangsor_ertek(szurt, kulcs, "koltseg")
            if ertek.sum() > 0:
                resz = ertek / ertek.sum()
                a[f"top3_{kulcs}"] = pd.DataFrame({
                    "#": range(1, min(3, len(ertek)) + 1),
                    kulcs.capitalize(): [rovid(x, 28) for x in ertek.index[:3]],
                    "Becsült költség": [ft(x) for x in ertek.values[:3]],
                    "Részesedés": [f"{pct(x, 1)}" for x in resz.values[:3]],
                    "Kumulált": [f"{pct(x, 1)}" for x in resz.cumsum().values[:3]],
                })
                a[f"top3_{kulcs}_resz"] = resz.head(3).sum()
                a[f"top3_{kulcs}_n80"] = sz.hany_adja_a_80_szazalekot(ertek)
                a[f"top3_{kulcs}_osszes"] = len(ertek)

    # Ugyanaz a cikk ugyanazzal a hibával (legalább 2×), esetszám szerint
    if "cikkszám" in szurt.columns:
        par = sz.ismetlodok(szurt, ["cikkszám", "hibatípus"], 2, "esetek")
        a["par_db"] = len(par)
        a["par"] = pd.DataFrame({
            "#": par["#"],
            "Cikkszám": par["cikkszám"].map(lambda x: rovid(x, 22)),
            cimkek.get("hibatípus", "Hibatípus"): par["hibatípus"].map(lambda x: rovid(x, 28)),
            "Esetek": par["esetek száma"],
            "Hibás db": par["hibás db"].map(szam),
            "Becsült költség": par["költség (Ft)"].map(ft),
            "Utolsó": par["utolsó"],
        }).head(10)

    # p-szabályozókártya (csak ha ismert a gyártott mennyiség)
    if mennyiseg is not None:
        pk, p_atlag, sigma_z = sz.p_kartya(nm, mennyiseg, tol, ig, laney)
        a.update(pk=pk, p_atlag=p_atlag, sigma_z=sigma_z, laney=laney,
                 kivul=pk.index[pk["kívül"]].strftime("%Y.%m.").tolist())
        if "selejt_2o_db" in pk.columns:  # a gyártott db összetétele (valós adatnál)
            a["p_nevezo"] = ("Gyártott db = bevételezés (jó termék + selejtpótlás) + "
                             f"selejt / másodosztály az NMF táblából "
                             f"({szam(pk['selejt_2o_db'].sum())} db).")
        # Második kártya: végleges veszteség (csak selejt + másodosztály a számlálóban)
        if "következmény" in nm.columns:
            import valos_adat  # itt importáljuk, mert csak valós adatnál kell
            veszt = nm[nm["következmény"].isin(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)]
            pkv, pv_atlag, pv_sigma = sz.p_kartya(veszt, mennyiseg, tol, ig, laney)
            a.update(pkv=pkv, pv_atlag=pv_atlag, pv_sigma=pv_sigma,
                     pv_kivul=pkv.index[pkv["kívül"]].strftime("%Y.%m.").tolist())
    return a


# --- 2) Grafikonok (matplotlib) -------------------------------------------------------
def _honap_cimkek(ax, honapok):
    ax.set_xticks(range(len(honapok)))
    ax.set_xticklabels([pd.Timestamp(h).strftime("%Y.%m") for h in honapok],
                       rotation=45, ha="right", fontsize=8)


def rajz_trend(ax, havi):
    x = np.arange(len(havi))
    ax.bar(x, havi["darab"], color=KEK, label="Nemmegfelelőségek (db)")
    ax.plot(x, havi["trend"], color=KEK, linestyle="--", label="Trend")
    ax.set_ylabel("esetek száma")
    _honap_cimkek(ax, havi["hónap"])
    ax2 = ax.twinx()
    ax2.plot(x, havi["koltseg"] / 1e6, color=PIROS, marker="o", label="Költség (M Ft)")
    ax2.set_ylabel("becsült költség (millió Ft)")
    ax2.set_ylim(bottom=0)
    sorok = ax.get_legend_handles_labels()[0] + ax2.get_legend_handles_labels()[0]
    ax.legend(sorok, [s.get_label() for s in sorok], loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=4, frameon=False, fontsize=8)
    ax.spines[["top"]].set_visible(False)


def rajz_pareto(ax, adat, ylabel="esetek száma"):
    x = np.arange(len(adat))
    ax.bar(x, adat.values, color=KEK)
    ax.set_xticks(x)
    ax.set_xticklabels([rovid(i, 22) for i in adat.index], rotation=35, ha="right", fontsize=8)
    ax.set_ylabel(ylabel)
    ax2 = ax.twinx()
    ax2.plot(x, adat.cumsum() / adat.sum() * 100, color=PIROS, marker="o")
    ax2.axhline(80, color=SZURKE, linestyle=":", linewidth=1)
    ax2.text(len(adat) - 0.5, 81, "80%", color=SZURKE, fontsize=8, ha="right")
    ax2.set_ylim(0, 105)
    ax2.set_ylabel("kumulált %")
    ax.spines[["top"]].set_visible(False)


def rajz_pkartya(ax, pk, p_atlag, laney):
    x = np.arange(len(pk))
    ax.plot(x, pk["p"] * 100, color=KEK, marker="o", label="Hibaarány (p)")
    ax.step(x, pk["FSH"] * 100, where="mid", color=PIROS, linestyle="--", label="FSH / ASH")
    ax.step(x, pk["ASH"] * 100, where="mid", color=PIROS, linestyle="--")
    ax.axhline(p_atlag * 100, color=ZOLD, label=f"átlag = {pct(p_atlag, 2)}")
    kivul = np.where(pk["kívül"])[0]
    ax.scatter(kivul, pk["p"].iloc[kivul] * 100, color="red", marker="x", s=90, zorder=5,
               label="Határon kívül")
    _honap_cimkek(ax, pk.index)
    ax.set_ylabel("hibaarány (%)")
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=4, frameon=False, fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)


def _png(rajzolo, *args, meret=(11, 4.8)) -> io.BytesIO:
    """Egy grafikont PNG képpé alakít (a PowerPointba)."""
    fig, ax = plt.subplots(figsize=meret, dpi=150)
    rajzolo(ax, *args)
    fig.tight_layout()
    kep = io.BytesIO()
    fig.savefig(kep, format="png")
    plt.close(fig)
    kep.seek(0)
    return kep


# --- 3) Megállapítások (automatikus, tényszerű mondatok) --------------------------------
def megallapitasok(a: dict) -> dict:
    m = {}
    if a["meredekseg"] is not None:
        irany = "csökken" if a["meredekseg"] < 0 else "nő"
        m["trend"] = (f"A trend szerint az esetszám havonta átlagosan "
                      f"{tized(abs(a['meredekseg']), 1)} esettel {irany}.")
    for kulcs in ["hibatípus", "ok"]:
        adat = a[f"pareto_{kulcs}"]
        if len(adat):
            nev = a["cimkek"].get(kulcs, kulcs).lower()
            m[f"pareto_{kulcs}"] = (
                f"A leggyakoribb: {adat.index[0]} ({pct(adat.iloc[0] / adat.sum(), 0)}). "
                f"Az esetek 80%-át az első {a[f'pareto_{kulcs}_n80']} {nev} adja "
                f"(összesen {a[f'pareto_{kulcs}_osszes']} különböző).")
    if "top3_cikkszám" in a:
        m["top3"] = (f"Az első 3 cikk a becsült költség {pct(a['top3_cikkszám_resz'], 0)}-át adja; "
                     f"a 80%-hoz {a['top3_cikkszám_n80']} cikk kell "
                     f"(összesen {a['top3_cikkszám_osszes']}).")
        if "top3_munkaszám" in a:
            m["top3"] += (f" Munkaszámoknál az első 3 részesedése "
                          f"{pct(a['top3_munkaszám_resz'], 0)}, a 80%-hoz "
                          f"{a['top3_munkaszám_n80']} munkaszám kell.")
    if "par_db" in a:
        m["par"] = (f"{a['par_db']} olyan cikk–hiba pár van, amely legalább kétszer előfordult – "
                    "ezeknél a korábbi intézkedés hatékonyságát érdemes felülvizsgálni.")
    if "pk" in a:
        kivul = a["kivul"]
        m["p"] = (f"Átlagos hibaarány: {pct(a['p_atlag'], 2)}. "
                  + (f"Határon kívüli hónap(ok): {', '.join(kivul)}." if kivul
                     else "Minden hónap a szabályozási határokon belül van.")
                  + f" (σz = {tized(a['sigma_z'], 2)})"
                  + (f" {a['p_nevezo']}" if "p_nevezo" in a else ""))
    if "pkv" in a:
        kivul = a["pv_kivul"]
        m["pv"] = (f"Átlagos végleges veszteség (selejt + másodosztály): "
                   f"{pct(a['pv_atlag'], 2)}. "
                   + (f"Határon kívüli hónap(ok): {', '.join(kivul)}." if kivul
                      else "Minden hónap a szabályozási határokon belül van.")
                   + f" (σz = {tized(a['pv_sigma'], 2)}) A javítható hibák itt nem számítanak.")
    return m


def p_cim(a: dict, veszteseg: bool = False) -> str:
    """A p-kártya oldal címe a kártya típusával."""
    tipus = "Laney p′" if a.get("laney", True) else "Klasszikus p"
    mit = "végleges veszteség (selejt + másodosztály)" if veszteseg else "havi hibaarány"
    return f"{tipus}-kártya – {mit}"


# --- 4) PowerPoint ---------------------------------------------------------------------
def _szoveg(dia, szoveg, bal, fent, szel, mag, meret=14, felkover=False, szin=None):
    doboz = dia.shapes.add_textbox(Inches(bal), Inches(fent), Inches(szel), Inches(mag))
    keret = doboz.text_frame
    keret.word_wrap = True
    for i, sor in enumerate(szoveg if isinstance(szoveg, list) else [szoveg]):
        bek = keret.paragraphs[0] if i == 0 else keret.add_paragraph()
        bek.text = sor
        bek.font.size = Pt(meret)
        bek.font.bold = felkover
        if szin:
            bek.font.color.rgb = RGBColor.from_string(szin.lstrip("#"))
    return doboz


def _uj_dia(prs, cim, a):
    dia = prs.slides.add_slide(prs.slide_layouts[6])  # üres elrendezés
    _szoveg(dia, cim, 0.5, 0.3, 12.3, 0.8, meret=26, felkover=True)
    _szoveg(dia, f"Időszak: {a['idoszak']}  |  Forrás: {a['forras']}  |  Készült: {a['keszult']}",
            0.5, 7.0, 12.3, 0.4, meret=9, szin=SZURKE)
    return dia


def _pptx_tablazat(dia, df, bal, fent, szel, sor_mag=0.32, meret=10):
    sorok, oszlopok = len(df) + 1, len(df.columns)
    tabla = dia.shapes.add_table(sorok, oszlopok, Inches(bal), Inches(fent), Inches(szel),
                                 Inches(sor_mag * sorok)).table
    if df.columns[0] == "#":  # a sorszám oszlop keskeny, a többi osztozik a helyen
        tabla.columns[0].width = Inches(0.5)
        for j in range(1, oszlopok):
            tabla.columns[j].width = Inches((szel - 0.5) / (oszlopok - 1))
    for j, nev in enumerate(df.columns):
        tabla.cell(0, j).text = str(nev)
    for i, sor in enumerate(df.itertuples(index=False), start=1):
        for j, ertek in enumerate(sor):
            tabla.cell(i, j).text = str(ertek)
    for cella in (tabla.cell(i, j) for i in range(sorok) for j in range(oszlopok)):
        for bek in cella.text_frame.paragraphs:
            bek.font.size = Pt(meret)
    return tabla


def pptx_keszit(a: dict) -> bytes:
    m = megallapitasok(a)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)  # 16:9

    # 1. Címlap
    dia = prs.slides.add_slide(prs.slide_layouts[6])
    _szoveg(dia, "Minőségügyi beszámoló", 0.8, 2.0, 11.7, 1.0, meret=40, felkover=True)
    _szoveg(dia, "Nemmegfelelőségek elemzése", 0.8, 3.0, 11.7, 0.6, meret=22, szin=SZURKE)
    _szoveg(dia, [f"Időszak: {a['idoszak']}", f"Adatforrás: {a['forras']}",
                  "Szűrők: " + "; ".join(a["szurok"]), f"Készült: {a['keszult']}"],
            0.8, 4.2, 11.7, 2.0, meret=14)

    # 2. Kulcsszámok
    dia = _uj_dia(prs, "Kulcsszámok", a)
    kpi = [("Nemmegfelelőségek", szam(a["esetek"])), ("Becsült összköltség", ft(a["koltseg"])),
           ("Hibás darab", szam(a["hibas_db"])), ("Nyitott / folyamatban", szam(a["nyitott"])),
           ("Átlagköltség / eset", ft(a["atlag"]))]
    if "pk" in a:
        kpi.append(("Átlagos hibaarány (p)", f"{pct(a['p_atlag'], 2)}"))
    if "pkv" in a:
        kpi.append(("Végleges veszteség (selejt + 2. o.)", f"{pct(a['pv_atlag'], 2)}"))
    for i, (nev, ertek) in enumerate(kpi):  # 4 doboz egy sorban
        bal, fent = 0.6 + (i % 4) * 3.1, 1.6 + (i // 4) * 2.4
        _szoveg(dia, ertek, bal, fent, 3.0, 0.9, meret=24, felkover=True, szin=KEK)
        _szoveg(dia, nev, bal, fent + 0.9, 3.0, 0.8, meret=13, szin=SZURKE)

    # 3–5. Grafikonok
    grafikonok = []
    if a["havi"] is not None:
        grafikonok.append(("Havi trend – esetszám és becsült költség",
                           _png(rajz_trend, a["havi"]), m.get("trend", "")))
    for kulcs in ["hibatípus", "ok"]:
        if len(a[f"pareto_{kulcs}"]) >= 2:
            grafikonok.append((f"Pareto – {a['cimkek'].get(kulcs, kulcs).lower()} szerint",
                               _png(rajz_pareto, a[f"pareto_{kulcs}"]),
                               m.get(f"pareto_{kulcs}", "")))
    for cim, kep, szoveg in grafikonok:
        dia = _uj_dia(prs, cim, a)
        dia.shapes.add_picture(kep, Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, szoveg, 0.6, 6.4, 12.1, 0.6, meret=13)

    # 6. Első 3 – becsült költség
    if "top3_cikkszám" in a:
        dia = _uj_dia(prs, "Első 3 – becsült költség szerint", a)
        _szoveg(dia, "Cikkek", 0.6, 1.2, 6, 0.4, meret=16, felkover=True)
        _pptx_tablazat(dia, a["top3_cikkszám"], 0.6, 1.65, 12.1)
        if "top3_munkaszám" in a:
            _szoveg(dia, "Munkaszámok", 0.6, 3.3, 6, 0.4, meret=16, felkover=True)
            _pptx_tablazat(dia, a["top3_munkaszám"], 0.6, 3.75, 12.1)
        _szoveg(dia, m.get("top3", ""), 0.6, 5.6, 12.1, 1.0, meret=13)

    # 7. Ismétlődő cikk–hiba párok
    if a.get("par") is not None and len(a["par"]):
        dia = _uj_dia(prs, "Ugyanaz a cikk, ugyanazzal a hibával (top 10)", a)
        _pptx_tablazat(dia, a["par"], 0.6, 1.2, 12.1, sor_mag=0.4, meret=10)
        _szoveg(dia, m.get("par", ""), 0.6, 6.2, 12.1, 0.7, meret=13)

    # 8. p-szabályozókártya
    dia = _uj_dia(prs, p_cim(a), a)
    if "pk" in a:
        dia.shapes.add_picture(_png(rajz_pkartya, a["pk"], a["p_atlag"], a["laney"]),
                               Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, m["p"], 0.6, 6.4, 12.1, 0.6, meret=13)
    else:
        _szoveg(dia, "Nincs gyártott mennyiség adat, ezért a p-kártya nem számolható.",
                0.6, 3.0, 12.1, 1.0, meret=16)

    # 8b. Végleges veszteség kártya (csak valós adatnál, ahol van "Következmény")
    if "pkv" in a:
        dia = _uj_dia(prs, p_cim(a, veszteseg=True), a)
        dia.shapes.add_picture(_png(rajz_pkartya, a["pkv"], a["pv_atlag"], a["laney"]),
                               Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, m["pv"], 0.6, 6.4, 12.1, 0.6, meret=13)

    # 9. Intézkedések (kézzel kitöltendő)
    dia = _uj_dia(prs, "Intézkedések", a)
    ures = pd.DataFrame([[""] * 5] * 6,
                        columns=["Probléma", "Intézkedés", "Felelős", "Határidő", "Státusz"])
    _pptx_tablazat(dia, ures, 0.6, 1.3, 12.1, sor_mag=0.6, meret=12)

    kimenet = io.BytesIO()
    prs.save(kimenet)
    return kimenet.getvalue()


# --- 5) PDF (ugyanaz a tartalom, fekvő A4 oldalakon) ----------------------------------
def _pdf_oldal(pdf, cim, a, rajz=None, szoveg="", tablazatok=(), cimlap=False):
    fig = plt.figure(figsize=(11.69, 8.27))  # fekvő A4
    if cimlap:
        fig.text(0.07, 0.62, "Minőségügyi beszámoló", fontsize=32, weight="bold")
        fig.text(0.07, 0.55, "Nemmegfelelőségek elemzése", fontsize=18, color=SZURKE)
        sorok = [f"Időszak: {a['idoszak']}", f"Adatforrás: {a['forras']}",
                 "Szűrők: " + "; ".join(a["szurok"]), f"Készült: {a['keszult']}"]
        for i, sor in enumerate(sorok):
            fig.text(0.07, 0.42 - i * 0.05, rovid(sor, 140), fontsize=12)
    else:
        fig.text(0.05, 0.93, cim, fontsize=20, weight="bold")
        fig.text(0.05, 0.03, f"Időszak: {a['idoszak']}  |  Forrás: {a['forras']}  |  "
                             f"Készült: {a['keszult']}", fontsize=8, color=SZURKE)
    if rajz is not None:
        ax = fig.add_axes([0.08, 0.22, 0.84, 0.64])
        rajz[0](ax, *rajz[1:])
    for felirat, df, fent, magas in tablazatok:
        if felirat:
            fig.text(0.05, fent + magas + 0.01, felirat, fontsize=13, weight="bold")
        ax = fig.add_axes([0.05, fent, 0.9, magas])
        ax.axis("off")
        t = ax.table(cellText=df.astype(str).values, colLabels=list(df.columns),
                     loc="upper center", cellLoc="left")
        if df.columns[0] == "#":  # a sorszám oszlop keskeny, a többi osztozik
            n = len(df.columns) - 1
            for (_, oszl), cella in t.get_celld().items():
                cella.set_width(0.05 if oszl == 0 else 0.95 / n)
        t.auto_set_font_size(False)
        t.set_fontsize(9)
        t.scale(1, 1.4)
        for (sor, _), cella in t.get_celld().items():
            if sor == 0:
                cella.set_text_props(weight="bold")
                cella.set_facecolor("#E8EEF5")
    if szoveg:
        fig.text(0.05, 0.09, szoveg, fontsize=11, wrap=True)
    pdf.savefig(fig)
    plt.close(fig)


def pdf_keszit(a: dict) -> bytes:
    m = megallapitasok(a)
    kimenet = io.BytesIO()
    with PdfPages(kimenet) as pdf:
        _pdf_oldal(pdf, "", a, cimlap=True)

        # Kulcsszámok
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.text(0.05, 0.93, "Kulcsszámok", fontsize=20, weight="bold")
        kpi = [("Nemmegfelelőségek", szam(a["esetek"])), ("Becsült összköltség", ft(a["koltseg"])),
               ("Hibás darab", szam(a["hibas_db"])), ("Nyitott / folyamatban", szam(a["nyitott"])),
               ("Átlagköltség / eset", ft(a["atlag"]))]
        if "pk" in a:
            kpi.append(("Átlagos hibaarány (p)", f"{pct(a['p_atlag'], 2)}"))
        if "pkv" in a:
            kpi.append(("Végleges veszteség\n(selejt + 2. o.)", f"{pct(a['pv_atlag'], 2)}"))
        for i, (nev, ertek) in enumerate(kpi):  # 4 doboz egy sorban
            x, y = 0.05 + (i % 4) * 0.235, 0.68 - (i // 4) * 0.3
            fig.text(x, y, ertek, fontsize=20, weight="bold", color=KEK)
            fig.text(x, y - 0.03, nev, fontsize=12, color=SZURKE, va="top")
        fig.text(0.05, 0.03, f"Időszak: {a['idoszak']}  |  Forrás: {a['forras']}  |  "
                             f"Készült: {a['keszult']}", fontsize=8, color=SZURKE)
        pdf.savefig(fig)
        plt.close(fig)

        if a["havi"] is not None:
            _pdf_oldal(pdf, "Havi trend – esetszám és becsült költség", a,
                       rajz=(rajz_trend, a["havi"]), szoveg=m.get("trend", ""))
        for kulcs in ["hibatípus", "ok"]:
            if len(a[f"pareto_{kulcs}"]) >= 2:
                _pdf_oldal(pdf, f"Pareto – {a['cimkek'].get(kulcs, kulcs).lower()} szerint", a,
                           rajz=(rajz_pareto, a[f"pareto_{kulcs}"]),
                           szoveg=m.get(f"pareto_{kulcs}", ""))
        if "top3_cikkszám" in a:
            tablak = [("Cikkek", a["top3_cikkszám"], 0.62, 0.2)]
            if "top3_munkaszám" in a:
                tablak.append(("Munkaszámok", a["top3_munkaszám"], 0.3, 0.2))
            _pdf_oldal(pdf, "Első 3 – becsült költség szerint", a, tablazatok=tablak,
                       szoveg=m.get("top3", ""))
        if a.get("par") is not None and len(a["par"]):
            _pdf_oldal(pdf, "Ugyanaz a cikk, ugyanazzal a hibával (top 10)", a,
                       tablazatok=[("", a["par"], 0.2, 0.68)], szoveg=m.get("par", ""))
        if "pk" in a:
            _pdf_oldal(pdf, p_cim(a), a,
                       rajz=(rajz_pkartya, a["pk"], a["p_atlag"], a["laney"]), szoveg=m["p"])
        if "pkv" in a:
            _pdf_oldal(pdf, p_cim(a, veszteseg=True), a,
                       rajz=(rajz_pkartya, a["pkv"], a["pv_atlag"], a["laney"]), szoveg=m["pv"])
        ures = pd.DataFrame([[""] * 5] * 6,
                            columns=["Probléma", "Intézkedés", "Felelős", "Határidő", "Státusz"])
        _pdf_oldal(pdf, "Intézkedések", a, tablazatok=[("", ures, 0.3, 0.55)])
    return kimenet.getvalue()
