# -----------------------------------------------------------------------------
# beszamolo.py
# EZ A FÁJL: a dashboard adataiból prezentálható beszámolót készít:
#   - PowerPoint (.pptx) – szerkeszthető, értekezletre
#   - PDF – ugyanaz a tartalom, archiváláshoz / e-mailhez
# A fájl a saját gépeden készül el (a böngésző letöltésként menti), semmit
# nem küld sehova. A grafikonokat matplotlib rajzolja (internet nem kell hozzá).
# Nyelv: a dashboardon választott nyelven (magyar / angol) készül.
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

import forditas  # noqa: E402
import szamitasok as sz  # noqa: E402

KEK, PIROS, ZOLD, SZURKE = "#4C78A8", "#E45756", "#54A24B", "#6B6B6B"


# --- Segédfüggvények ------------------------------------------------------------
def rovid(szoveg, hossz=30) -> str:
    """Hosszú szöveg levágása, hogy elférjen a táblázatban / tengelyen."""
    szoveg = str(szoveg)
    return szoveg if len(szoveg) <= hossz else szoveg[: hossz - 1] + "…"


def _honap(h, en) -> str:
    """Hónap felirat: magyarul 2025.03., angolul Mar 2025."""
    return pd.Timestamp(h).strftime("%b %Y" if en else "%Y.%m.")


# --- 1) Adatok összeállítása ---------------------------------------------------------
def adat_osszeallitas(szurt, nm, mennyiseg, tol, ig, laney, szurok, forras, cimkek,
                      en=False) -> dict:
    """Minden számot és táblázatot kiszámol, ami a beszámolóba kerül.
    szurok: a beállított szűrők szöveges leírása (lista)
    cimkek: belső oszlopnév -> felirat (pl. 'hibatípus' -> 'Nemmegfelelőség jellege')
    en: True = angol nyelvű beszámoló"""
    L = forditas.nyelv(en)

    def ft(x):
        return forditas.ft(x, en)

    def pct(x, t=1):
        return forditas.pct(x, t, en)

    a = {
        "en": en,
        "idoszak": f"{_honap(tol, en)} – {_honap(ig, en)}",
        "szurok": szurok or [L("nincs (minden adat)", "none (all data)")],
        "forras": forras,
        "keszult": date.today().strftime("%Y-%m-%d" if en else "%Y.%m.%d"),
        "cimkek": cimkek,
        "esetek": len(szurt),
        "koltseg": szurt["költség"].sum(),
        "nyitott": int(forditas.nyitott_e(szurt["státusz"]).sum()),
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
        a[f"pareto_{kulcs}"] = sz.pareto_adat(szurt, kulcs, "Darabszám", max_oszlop=12, en=en)
        a[f"pareto_{kulcs}_n80"] = sz.hany_adja_a_80_szazalekot(teljes) if len(teljes) else 0
        a[f"pareto_{kulcs}_osszes"] = len(teljes)

    # Első 3 cikk / munkaszám becsült költség szerint (minden tétel számít)
    kulcs_nev = {"cikkszám": L("Cikkszám", "Item no."), "munkaszám": L("Munkaszám", "Job no.")}
    for kulcs in ["cikkszám", "munkaszám"]:
        if kulcs in szurt.columns:
            ertek = sz.rangsor_ertek(szurt, kulcs, "koltseg")
            if ertek.sum() > 0:
                resz = ertek / ertek.sum()
                a[f"top3_{kulcs}"] = pd.DataFrame({
                    "#": range(1, min(3, len(ertek)) + 1),
                    kulcs_nev[kulcs]: [rovid(x, 28) for x in ertek.index[:3]],
                    L("Becsült költség", "Estimated cost"): [ft(x) for x in ertek.values[:3]],
                    L("Részesedés", "Share"): [pct(x) for x in resz.values[:3]],
                    L("Kumulált", "Cumulative"): [pct(x) for x in resz.cumsum().values[:3]],
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
            L("Cikkszám", "Item no."): par["cikkszám"].map(lambda x: rovid(x, 22)),
            cimkek.get("hibatípus", L("Hibatípus", "Defect type")):
                par["hibatípus"].map(lambda x: rovid(x, 28)),
            L("Esetek", "Cases"): par["esetek száma"],
            L("Hibás db", "Defective qty"): par["hibás db"].map(lambda x: forditas.szam(x, en)),
            L("Becsült költség", "Estimated cost"): par["költség (Ft)"].map(ft),
            L("Utolsó", "Last"): par["utolsó"],
        }).head(10)

    # p-szabályozókártya (csak ha ismert a gyártott mennyiség)
    if mennyiseg is not None:
        pk, p_atlag, sigma_z = sz.p_kartya(nm, mennyiseg, tol, ig, laney)
        a.update(pk=pk, p_atlag=p_atlag, sigma_z=sigma_z, laney=laney,
                 kivul=[_honap(h, en) for h in pk.index[pk["kívül"]]])
        if "selejt_2o_db" in pk.columns:  # a gyártott db összetétele (valós adatnál)
            db = forditas.szam(pk["selejt_2o_db"].sum(), en)
            a["p_nevezo"] = L("Gyártott db = bevételezés (jó termék + selejtpótlás) + "
                              f"selejt / másodosztály az NMF táblából ({db} db).",
                              "Produced qty = receipts (good products + scrap replacement) + "
                              f"scrap / 2nd class from the NC table ({db} pcs).")
        # Második kártya: végleges veszteség (csak selejt + másodosztály a számlálóban)
        if "következmény" in nm.columns:
            import valos_adat  # itt importáljuk, mert csak valós adatnál kell
            veszt = nm[nm["következmény"].isin(valos_adat.NEM_BEERKEZETT_KOVETKEZMENYEK)]
            pkv, pv_atlag, pv_sigma = sz.p_kartya(veszt, mennyiseg, tol, ig, laney)
            a.update(pkv=pkv, pv_atlag=pv_atlag, pv_sigma=pv_sigma,
                     pv_kivul=[_honap(h, en) for h in pkv.index[pkv["kívül"]]])
    return a


# --- 2) Grafikonok (matplotlib) -------------------------------------------------------
def _honap_cimkek(ax, honapok, en):
    ax.set_xticks(range(len(honapok)))
    ax.set_xticklabels([pd.Timestamp(h).strftime("%b %Y" if en else "%Y.%m") for h in honapok],
                       rotation=45, ha="right", fontsize=8)


def rajz_trend(ax, havi, en=False):
    L = forditas.nyelv(en)
    x = np.arange(len(havi))
    ax.bar(x, havi["darab"], color=KEK, label=L("Nemmegfelelőségek (db)", "Nonconformities"))
    ax.plot(x, havi["trend"], color=KEK, linestyle="--", label="Trend")
    ax.set_ylabel(L("esetek száma", "number of cases"))
    _honap_cimkek(ax, havi["hónap"], en)
    ax2 = ax.twinx()
    ax2.plot(x, havi["koltseg"] / 1e6, color=PIROS, marker="o",
             label=L("Költség (M Ft)", "Cost (M HUF)"))
    ax2.set_ylabel(L("becsült költség (millió Ft)", "estimated cost (million HUF)"))
    ax2.set_ylim(bottom=0)
    sorok = ax.get_legend_handles_labels()[0] + ax2.get_legend_handles_labels()[0]
    ax.legend(sorok, [s.get_label() for s in sorok], loc="lower center",
              bbox_to_anchor=(0.5, 1.0), ncol=4, frameon=False, fontsize=8)
    ax.spines[["top"]].set_visible(False)


def rajz_pareto(ax, adat, en=False):
    L = forditas.nyelv(en)
    x = np.arange(len(adat))
    ax.bar(x, adat.values, color=KEK)
    ax.set_xticks(x)
    ax.set_xticklabels([rovid(i, 22) for i in adat.index], rotation=35, ha="right", fontsize=8)
    ax.set_ylabel(L("esetek száma", "number of cases"))
    ax2 = ax.twinx()
    ax2.plot(x, adat.cumsum() / adat.sum() * 100, color=PIROS, marker="o")
    ax2.axhline(80, color=SZURKE, linestyle=":", linewidth=1)
    ax2.text(len(adat) - 0.5, 81, "80%", color=SZURKE, fontsize=8, ha="right")
    ax2.set_ylim(0, 105)
    ax2.set_ylabel(L("kumulált %", "cumulative %"))
    ax.spines[["top"]].set_visible(False)


def rajz_pkartya(ax, pk, p_atlag, laney, en=False):
    L = forditas.nyelv(en)
    x = np.arange(len(pk))
    ax.plot(x, pk["p"] * 100, color=KEK, marker="o", label=L("Hibaarány (p)", "Defect rate (p)"))
    ax.step(x, pk["FSH"] * 100, where="mid", color=PIROS, linestyle="--",
            label=L("FSH / ASH", "UCL / LCL"))
    ax.step(x, pk["ASH"] * 100, where="mid", color=PIROS, linestyle="--")
    ax.axhline(p_atlag * 100, color=ZOLD,
               label=L("átlag", "average") + f" = {forditas.pct(p_atlag, 2, en)}")
    kivul = np.where(pk["kívül"])[0]
    ax.scatter(kivul, pk["p"].iloc[kivul] * 100, color="red", marker="x", s=90, zorder=5,
               label=L("Határon kívül", "Out of limits"))
    _honap_cimkek(ax, pk.index, en)
    ax.set_ylabel(L("hibaarány (%)", "defect rate (%)"))
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
    en = a["en"]
    L = forditas.nyelv(en)

    def pct(x, t=1):
        return forditas.pct(x, t, en)

    def tized(x, t=1):
        return forditas.tized(x, t, en)

    m = {}
    if a["meredekseg"] is not None:
        irany = (L("csökken", "decreases") if a["meredekseg"] < 0 else L("nő", "increases"))
        m["trend"] = L(f"A trend szerint az esetszám havonta átlagosan "
                       f"{tized(abs(a['meredekseg']))} esettel {irany}.",
                       f"According to the trend, the number of cases {irany} by "
                       f"{tized(abs(a['meredekseg']))} per month on average.")
    for kulcs in ["hibatípus", "ok"]:
        adat = a[f"pareto_{kulcs}"]
        if len(adat):
            nev = a["cimkek"].get(kulcs, kulcs).lower()
            n80, osszes = a[f"pareto_{kulcs}_n80"], a[f"pareto_{kulcs}_osszes"]
            m[f"pareto_{kulcs}"] = L(
                f"A leggyakoribb: {adat.index[0]} ({pct(adat.iloc[0] / adat.sum(), 0)}). "
                f"Az esetek 80%-át az első {n80} {nev} adja (összesen {osszes} különböző).",
                f"Most frequent: {adat.index[0]} ({pct(adat.iloc[0] / adat.sum(), 0)}). "
                f"The top {n80} ({nev}) account for 80% of cases ({osszes} different in total).")
    if "top3_cikkszám" in a:
        m["top3"] = L(f"Az első 3 cikk a becsült költség {pct(a['top3_cikkszám_resz'], 0)}-át "
                      f"adja; a 80%-hoz {a['top3_cikkszám_n80']} cikk kell "
                      f"(összesen {a['top3_cikkszám_osszes']}).",
                      f"The top 3 items account for {pct(a['top3_cikkszám_resz'], 0)} of the "
                      f"estimated cost; {a['top3_cikkszám_n80']} items are needed to reach 80% "
                      f"(out of {a['top3_cikkszám_osszes']}).")
        if "top3_munkaszám" in a:
            m["top3"] += L(f" Munkaszámoknál az első 3 részesedése "
                           f"{pct(a['top3_munkaszám_resz'], 0)}, a 80%-hoz "
                           f"{a['top3_munkaszám_n80']} munkaszám kell.",
                           f" For job numbers the top 3 share is "
                           f"{pct(a['top3_munkaszám_resz'], 0)}; {a['top3_munkaszám_n80']} "
                           "job numbers are needed to reach 80%.")
    if "par_db" in a:
        m["par"] = L(f"{a['par_db']} olyan cikk–hiba pár van, amely legalább kétszer "
                     "előfordult – ezeknél a korábbi intézkedés hatékonyságát érdemes "
                     "felülvizsgálni.",
                     f"{a['par_db']} item–defect pairs occurred at least twice – the "
                     "effectiveness of earlier actions should be reviewed for these.")
    belul = L("Minden hónap a szabályozási határokon belül van.",
              "All months are within the control limits.")
    if "pk" in a:
        kivul = a["kivul"]
        m["p"] = (L(f"Átlagos hibaarány: {pct(a['p_atlag'], 2)}. ",
                    f"Average defect rate: {pct(a['p_atlag'], 2)}. ")
                  + (L("Határon kívüli hónap(ok): ", "Month(s) out of limits: ")
                     + ", ".join(kivul) + "." if kivul else belul)
                  + f" (σz = {tized(a['sigma_z'], 2)})"
                  + (f" {a['p_nevezo']}" if "p_nevezo" in a else ""))
    if "pkv" in a:
        kivul = a["pv_kivul"]
        m["pv"] = (L(f"Átlagos végleges veszteség (selejt + másodosztály): "
                     f"{pct(a['pv_atlag'], 2)}. ",
                     f"Average final loss (scrap + 2nd class): {pct(a['pv_atlag'], 2)}. ")
                   + (L("Határon kívüli hónap(ok): ", "Month(s) out of limits: ")
                      + ", ".join(kivul) + "." if kivul else belul)
                   + f" (σz = {tized(a['pv_sigma'], 2)}) "
                   + L("A javítható hibák itt nem számítanak.",
                       "Repairable defects are not counted here."))
    return m


def p_cim(a: dict, veszteseg: bool = False) -> str:
    """A p-kártya oldal címe a kártya típusával."""
    L = forditas.nyelv(a["en"])
    tipus = "Laney p′" if a.get("laney", True) else L("Klasszikus p", "Classic p")
    mit = (L("végleges veszteség (selejt + másodosztály)", "final loss (scrap + 2nd class)")
           if veszteseg else L("havi hibaarány", "monthly defect rate"))
    return L(f"{tipus}-kártya – {mit}", f"{tipus} chart – {mit}")


def _kpi(a: dict) -> list:
    """A kulcsszámok listája: (felirat, érték)."""
    en = a["en"]
    L = forditas.nyelv(en)
    kpi = [(L("Nemmegfelelőségek", "Nonconformities"), forditas.szam(a["esetek"], en)),
           (L("Becsült összköltség", "Estimated total cost"), forditas.ft(a["koltseg"], en)),
           (L("Hibás darab", "Defective parts"), forditas.szam(a["hibas_db"], en)),
           (L("Nyitott / folyamatban", "Open / in progress"), forditas.szam(a["nyitott"], en)),
           (L("Átlagköltség / eset", "Average cost / case"), forditas.ft(a["atlag"], en))]
    if "pk" in a:
        kpi.append((L("Átlagos hibaarány (p)", "Average defect rate (p)"),
                    forditas.pct(a["p_atlag"], 2, en)))
    if "pkv" in a:
        kpi.append((L("Végleges veszteség (selejt + 2. o.)", "Final loss (scrap + 2nd class)"),
                    forditas.pct(a["pv_atlag"], 2, en)))
    return kpi


def _szovegek(a: dict) -> dict:
    """A beszámoló állandó feliratai a választott nyelven."""
    L = forditas.nyelv(a["en"])
    return {
        "cim": L("Minőségügyi beszámoló", "Quality report"),
        "alcim": L("Nemmegfelelőségek elemzése", "Nonconformity analysis"),
        "idoszak": L("Időszak", "Period"), "forras": L("Adatforrás", "Data source"),
        "forras_rovid": L("Forrás", "Source"), "szurok": L("Szűrők", "Filters"),
        "keszult": L("Készült", "Created"),
        "kulcsszamok": L("Kulcsszámok", "Key figures"),
        "trend": L("Havi trend – esetszám és becsült költség",
                   "Monthly trend – cases and estimated cost"),
        "pareto": L("Pareto – {} szerint", "Pareto – by {}"),
        "top3": L("Első 3 – becsült költség szerint", "Top 3 – by estimated cost"),
        "cikkek": L("Cikkek", "Items"), "munkaszamok": L("Munkaszámok", "Job numbers"),
        "par": L("Ugyanaz a cikk, ugyanazzal a hibával (top 10)",
                 "Same item, same defect (top 10)"),
        "nincs_p": L("Nincs gyártott mennyiség adat, ezért a p-kártya nem számolható.",
                     "No production quantity data, so the p chart cannot be calculated."),
        "intezkedesek": L("Intézkedések", "Actions"),
        "intezkedes_oszlopok": [L("Probléma", "Problem"), L("Intézkedés", "Action"),
                                L("Felelős", "Owner"), L("Határidő", "Deadline"),
                                L("Státusz", "Status")],
    }


def _lablec(a: dict, sz_: dict) -> str:
    return (f"{sz_['idoszak']}: {a['idoszak']}  |  {sz_['forras_rovid']}: {a['forras']}  |  "
            f"{sz_['keszult']}: {a['keszult']}")


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


def _uj_dia(prs, cim, a, sz_):
    dia = prs.slides.add_slide(prs.slide_layouts[6])  # üres elrendezés
    _szoveg(dia, cim, 0.5, 0.3, 12.3, 0.8, meret=26, felkover=True)
    _szoveg(dia, _lablec(a, sz_), 0.5, 7.0, 12.3, 0.4, meret=9, szin=SZURKE)
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
    en = a["en"]
    m = megallapitasok(a)
    sz_ = _szovegek(a)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)  # 16:9

    # 1. Címlap
    dia = prs.slides.add_slide(prs.slide_layouts[6])
    _szoveg(dia, sz_["cim"], 0.8, 2.0, 11.7, 1.0, meret=40, felkover=True)
    _szoveg(dia, sz_["alcim"], 0.8, 3.0, 11.7, 0.6, meret=22, szin=SZURKE)
    _szoveg(dia, [f"{sz_['idoszak']}: {a['idoszak']}", f"{sz_['forras']}: {a['forras']}",
                  f"{sz_['szurok']}: " + "; ".join(a["szurok"]),
                  f"{sz_['keszult']}: {a['keszult']}"],
            0.8, 4.2, 11.7, 2.0, meret=14)

    # 2. Kulcsszámok
    dia = _uj_dia(prs, sz_["kulcsszamok"], a, sz_)
    for i, (nev, ertek) in enumerate(_kpi(a)):  # 4 doboz egy sorban
        bal, fent = 0.6 + (i % 4) * 3.1, 1.6 + (i // 4) * 2.4
        _szoveg(dia, ertek, bal, fent, 3.0, 0.9, meret=24, felkover=True, szin=KEK)
        _szoveg(dia, nev, bal, fent + 0.9, 3.0, 0.8, meret=13, szin=SZURKE)

    # 3–5. Grafikonok
    grafikonok = []
    if a["havi"] is not None:
        grafikonok.append((sz_["trend"], _png(rajz_trend, a["havi"], en), m.get("trend", "")))
    for kulcs in ["hibatípus", "ok"]:
        if len(a[f"pareto_{kulcs}"]) >= 2:
            grafikonok.append((sz_["pareto"].format(a["cimkek"].get(kulcs, kulcs).lower()),
                               _png(rajz_pareto, a[f"pareto_{kulcs}"], en),
                               m.get(f"pareto_{kulcs}", "")))
    for cim, kep, szoveg in grafikonok:
        dia = _uj_dia(prs, cim, a, sz_)
        dia.shapes.add_picture(kep, Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, szoveg, 0.6, 6.4, 12.1, 0.6, meret=13)

    # 6. Első 3 – becsült költség
    if "top3_cikkszám" in a:
        dia = _uj_dia(prs, sz_["top3"], a, sz_)
        _szoveg(dia, sz_["cikkek"], 0.6, 1.2, 6, 0.4, meret=16, felkover=True)
        _pptx_tablazat(dia, a["top3_cikkszám"], 0.6, 1.65, 12.1)
        if "top3_munkaszám" in a:
            _szoveg(dia, sz_["munkaszamok"], 0.6, 3.3, 6, 0.4, meret=16, felkover=True)
            _pptx_tablazat(dia, a["top3_munkaszám"], 0.6, 3.75, 12.1)
        _szoveg(dia, m.get("top3", ""), 0.6, 5.6, 12.1, 1.0, meret=13)

    # 7. Ismétlődő cikk–hiba párok
    if a.get("par") is not None and len(a["par"]):
        dia = _uj_dia(prs, sz_["par"], a, sz_)
        _pptx_tablazat(dia, a["par"], 0.6, 1.2, 12.1, sor_mag=0.4, meret=10)
        _szoveg(dia, m.get("par", ""), 0.6, 6.2, 12.1, 0.7, meret=13)

    # 8. p-szabályozókártya
    dia = _uj_dia(prs, p_cim(a), a, sz_)
    if "pk" in a:
        dia.shapes.add_picture(_png(rajz_pkartya, a["pk"], a["p_atlag"], a["laney"], en),
                               Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, m["p"], 0.6, 6.4, 12.1, 0.6, meret=13)
    else:
        _szoveg(dia, sz_["nincs_p"], 0.6, 3.0, 12.1, 1.0, meret=16)

    # 8b. Végleges veszteség kártya (csak valós adatnál, ahol van "Következmény")
    if "pkv" in a:
        dia = _uj_dia(prs, p_cim(a, veszteseg=True), a, sz_)
        dia.shapes.add_picture(_png(rajz_pkartya, a["pkv"], a["pv_atlag"], a["laney"], en),
                               Inches(0.88), Inches(1.15), height=Inches(5.05))
        _szoveg(dia, m["pv"], 0.6, 6.4, 12.1, 0.6, meret=13)

    # 9. Intézkedések (kézzel kitöltendő)
    dia = _uj_dia(prs, sz_["intezkedesek"], a, sz_)
    ures = pd.DataFrame([[""] * 5] * 6, columns=sz_["intezkedes_oszlopok"])
    _pptx_tablazat(dia, ures, 0.6, 1.3, 12.1, sor_mag=0.6, meret=12)

    kimenet = io.BytesIO()
    prs.save(kimenet)
    return kimenet.getvalue()


# --- 5) PDF (ugyanaz a tartalom, fekvő A4 oldalakon) ----------------------------------
def _pdf_oldal(pdf, cim, a, sz_, rajz=None, szoveg="", tablazatok=(), cimlap=False):
    fig = plt.figure(figsize=(11.69, 8.27))  # fekvő A4
    if cimlap:
        fig.text(0.07, 0.62, sz_["cim"], fontsize=32, weight="bold")
        fig.text(0.07, 0.55, sz_["alcim"], fontsize=18, color=SZURKE)
        sorok = [f"{sz_['idoszak']}: {a['idoszak']}", f"{sz_['forras']}: {a['forras']}",
                 f"{sz_['szurok']}: " + "; ".join(a["szurok"]),
                 f"{sz_['keszult']}: {a['keszult']}"]
        for i, sor in enumerate(sorok):
            fig.text(0.07, 0.42 - i * 0.05, rovid(sor, 140), fontsize=12)
    else:
        fig.text(0.05, 0.93, cim, fontsize=20, weight="bold")
        fig.text(0.05, 0.03, _lablec(a, sz_), fontsize=8, color=SZURKE)
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
    en = a["en"]
    m = megallapitasok(a)
    sz_ = _szovegek(a)
    kimenet = io.BytesIO()
    with PdfPages(kimenet) as pdf:
        _pdf_oldal(pdf, "", a, sz_, cimlap=True)

        # Kulcsszámok
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.text(0.05, 0.93, sz_["kulcsszamok"], fontsize=20, weight="bold")
        for i, (nev, ertek) in enumerate(_kpi(a)):  # 4 doboz egy sorban
            x, y = 0.05 + (i % 4) * 0.235, 0.68 - (i // 4) * 0.3
            fig.text(x, y, ertek, fontsize=20, weight="bold", color=KEK)
            fig.text(x, y - 0.03, nev.replace(" (", "\n("), fontsize=12, color=SZURKE, va="top")
        fig.text(0.05, 0.03, _lablec(a, sz_), fontsize=8, color=SZURKE)
        pdf.savefig(fig)
        plt.close(fig)

        if a["havi"] is not None:
            _pdf_oldal(pdf, sz_["trend"], a, sz_, rajz=(rajz_trend, a["havi"], en),
                       szoveg=m.get("trend", ""))
        for kulcs in ["hibatípus", "ok"]:
            if len(a[f"pareto_{kulcs}"]) >= 2:
                _pdf_oldal(pdf, sz_["pareto"].format(a["cimkek"].get(kulcs, kulcs).lower()),
                           a, sz_, rajz=(rajz_pareto, a[f"pareto_{kulcs}"], en),
                           szoveg=m.get(f"pareto_{kulcs}", ""))
        if "top3_cikkszám" in a:
            tablak = [(sz_["cikkek"], a["top3_cikkszám"], 0.62, 0.2)]
            if "top3_munkaszám" in a:
                tablak.append((sz_["munkaszamok"], a["top3_munkaszám"], 0.3, 0.2))
            _pdf_oldal(pdf, sz_["top3"], a, sz_, tablazatok=tablak, szoveg=m.get("top3", ""))
        if a.get("par") is not None and len(a["par"]):
            _pdf_oldal(pdf, sz_["par"], a, sz_,
                       tablazatok=[("", a["par"], 0.2, 0.68)], szoveg=m.get("par", ""))
        if "pk" in a:
            _pdf_oldal(pdf, p_cim(a), a, sz_,
                       rajz=(rajz_pkartya, a["pk"], a["p_atlag"], a["laney"], en), szoveg=m["p"])
        if "pkv" in a:
            _pdf_oldal(pdf, p_cim(a, veszteseg=True), a, sz_,
                       rajz=(rajz_pkartya, a["pkv"], a["pv_atlag"], a["laney"], en),
                       szoveg=m["pv"])
        ures = pd.DataFrame([[""] * 5] * 6, columns=sz_["intezkedes_oszlopok"])
        _pdf_oldal(pdf, sz_["intezkedesek"], a, sz_, tablazatok=[("", ures, 0.3, 0.55)])
    return kimenet.getvalue()
