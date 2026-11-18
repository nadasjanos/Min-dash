# -----------------------------------------------------------------------------
# szamitasok.py
# EZ A FÁJL: a dashboard és a beszámoló (PowerPoint / PDF) KÖZÖS számításai.
# Azért van külön fájlban, hogy a képernyőn és a beszámolóban biztosan
# ugyanazok a számok jelenjenek meg. Grafikont nem rajzol, csak számol.
# -----------------------------------------------------------------------------

import numpy as np
import pandas as pd

# Rangsorolási lehetőségek az Ismétlődések fülön: felirat -> belső oszlopnév
RANGSOR = {"Esetek száma": "esetek", "Becsült költség": "koltseg", "Hibás db": "hibas_db"}


def havi_osszesites(df: pd.DataFrame) -> pd.DataFrame:
    """Havonta: esetszám (darab), költség, és a darabszám lineáris trendje."""
    havi = df.groupby("hónap").agg(darab=("hibatípus", "size"),
                                   koltseg=("költség", "sum")).reset_index()
    x = np.arange(len(havi))
    # Egyszerű lineáris trendvonal (legkisebb négyzetek módszere)
    havi["trend"] = (np.polyval(np.polyfit(x, havi["darab"], 1), x)
                     if len(havi) > 1 else havi["darab"])
    return havi


def pareto_adat(df: pd.DataFrame, bontas: str, mertek: str = "Darabszám",
                max_oszlop: int = 15) -> pd.Series:
    """Pareto-sorrend (csökkenő). 15-nél több kategóriánál a többi 'Egyéb'-be kerül."""
    if mertek == "Darabszám":
        adat = df[bontas].value_counts()
    else:
        adat = df.groupby(bontas)["költség"].sum().sort_values(ascending=False)
    if len(adat) > max_oszlop:  # a legnagyobbak külön, a többi összevonva a végén
        egyeb = adat.iloc[max_oszlop:].sum()
        tobbi = len(adat) - max_oszlop
        adat = adat.iloc[:max_oszlop].copy()
        adat[f"Egyéb ({tobbi} db)"] = egyeb
    return adat


def hany_adja_a_80_szazalekot(ertek: pd.Series) -> int:
    """Csökkenő sorrendben hány tétel adja együtt az összeg 80%-át."""
    ertek = ertek.sort_values(ascending=False)
    return int((ertek.cumsum() / ertek.sum() < 0.8).sum()) + 1


def p_kartya(nm: pd.DataFrame, mennyiseg: pd.DataFrame, tol, ig, laney: bool):
    """p-szabályozókártya számítása a [tol, ig] hónapokra.
    Visszaadja: (táblázat havonta, p̄ átlag, σz Laney-tényező)."""
    nm_havi = nm.assign(hónap=nm["dátum"].dt.to_period("M").dt.to_timestamp())
    # Hibás darabok havonta (valós adatnál az "Nmf db" oszlop összege)
    hibas_havi = nm_havi.groupby("hónap")["nmf_db"].sum().rename("hibás_db")
    pk = mennyiseg.set_index("hónap").join(hibas_havi).fillna({"hibás_db": 0})
    pk = pk[(pk.index >= tol) & (pk.index <= ig)].copy()  # csak a kiválasztott hónapok

    # p = hibás / gyártott;  p̄ = összes hibás / összes gyártott
    pk["p"] = pk["hibás_db"] / pk["gyártott_db"]
    p_atlag = pk["hibás_db"].sum() / pk["gyártott_db"].sum()
    # Klasszikus szórás havonta: √(p̄(1−p̄)/n)  (n havonta más, ezért "lépcsős")
    szoras = np.sqrt(p_atlag * (1 - p_atlag) / pk["gyártott_db"])

    # Laney-korrekció (σz): mennyivel nagyobb a hónapok közötti valódi ingadozás,
    # mint amit a klasszikus modell vár.
    #   z = (p − p̄) / szórás  ->  σz = (szomszédos z-k átlagos eltérése) / 1,128
    # Ha σz = 1, a két kártya ugyanaz; ha σz > 1, a Laney-határ ennyiszer szélesebb.
    z = (pk["p"] - p_atlag) / szoras
    sigma_z = z.diff().abs().mean() / 1.128 if len(pk) > 1 else 1.0
    szorzo = sigma_z if laney else 1.0

    pk["FSH"] = p_atlag + 3 * szoras * szorzo                 # felső szabályozási határ
    pk["ASH"] = (p_atlag - 3 * szoras * szorzo).clip(lower=0)  # alsó (nem lehet negatív)
    pk["kívül"] = (pk["p"] > pk["FSH"]) | (pk["p"] < pk["ASH"])
    return pk, p_atlag, sigma_z


def hibak_felsorolva(s: pd.Series) -> str:
    """Pl. 'Kerülethiba (3), Méreteltérés (1)' – gyakoriság szerint csökkenő sorrendben."""
    return ", ".join(f"{nev} ({db})" for nev, db in s.value_counts().items())


def ismetlodok(df: pd.DataFrame, kulcsok: list[str], min_db: int,
               rendezes: str = "esetek") -> pd.DataFrame:
    """Azokat a kulcsokat (pl. munkaszám) adja vissza, amelyek legalább min_db-szer
    szerepelnek, a 'rendezes' oszlop szerint rangsorolva (esetek / koltseg / hibas_db)."""
    df = df[(df[kulcsok] != "Nincs megadva").all(axis=1)]  # üres azonosítót nem számolunk
    csoport = df.groupby(kulcsok)
    tabla = csoport.agg(
        esetek=("dátum", "size"),
        hibas_db=("nmf_db", "sum"),
        koltseg=("költség", "sum"),
        elso=("dátum", "min"),
        utolso=("dátum", "max"),
    )
    if "hibatípus" not in kulcsok:
        tabla["hibák"] = csoport["hibatípus"].agg(hibak_felsorolva)
    if "munkaszám" not in kulcsok and "munkaszám" in df.columns:
        tabla["munkaszámok"] = csoport["munkaszám"].agg(lambda s: ", ".join(sorted(s.unique())))
    if "cikkszám" not in kulcsok and "cikkszám" in df.columns:
        tabla["cikkszám"] = csoport["cikkszám"].agg(lambda s: ", ".join(sorted(s.unique())))
    masodlagos = "koltseg" if rendezes != "koltseg" else "esetek"  # egyenlőségnél ez dönt
    tabla = (tabla[tabla["esetek"] >= min_db]
             .sort_values([rendezes, masodlagos], ascending=False).reset_index())
    tabla.insert(0, "#", range(1, len(tabla) + 1))  # rangsor
    tabla["elso"] = tabla["elso"].dt.strftime("%Y.%m.%d")
    tabla["utolso"] = tabla["utolso"].dt.strftime("%Y.%m.%d")
    return tabla.rename(columns={"esetek": "esetek száma", "hibas_db": "hibás db",
                                 "koltseg": "költség (Ft)", "elso": "első", "utolso": "utolsó"})


def rangsor_ertek(df: pd.DataFrame, kulcs: str, rendezes: str) -> pd.Series:
    """Tételenkénti érték (esetszám / költség / hibás db), csökkenő sorrendben.
    Itt MINDEN tétel számít (az egyszer szereplők is) – ebből jön a Top 3 részesedés."""
    df = df[df[kulcs] != "Nincs megadva"]
    if rendezes == "esetek":
        ertek = df.groupby(kulcs).size()
    else:
        ertek = df.groupby(kulcs)["költség" if rendezes == "koltseg" else "nmf_db"].sum()
    return ertek.sort_values(ascending=False)
