# -----------------------------------------------------------------------------
# forditas.py
# EZ A FÁJL: a magyar / angol nyelv kezelése.
#   - nyelv(en): egy kis függvényt ad, amivel a kódban mindkét nyelvű szöveget
#     egymás mellé írjuk:  L("Szűrők", "Filters")
#   - szám-, pénz- és százalékformázás az adott nyelv szokása szerint
#     (magyar: 1 234,5 Ft  |  angol: 1,234.5 HUF)
#   - a kitalált (szintetikus) adat értékeinek és a státuszoknak a fordítása
# A valós Excel értékeit (vevőnevek, hibaleírások) NEM fordítjuk – azok a te adataid.
# -----------------------------------------------------------------------------

import pandas as pd


def nyelv(en: bool):
    """Visszaad egy L(magyar, angol) függvényt, ami a választott nyelvű szöveget adja."""
    def L(hu: str, angol: str) -> str:
        return angol if en else hu
    return L


# --- Számformázás ------------------------------------------------------------
def szam(x, en: bool = False) -> str:
    """Egész szám ezres tagolással: magyarul 1 234, angolul 1,234."""
    s = f"{x:,.0f}"
    return s if en else s.replace(",", " ")


def ft(x, en: bool = False) -> str:
    """Pénzösszeg: magyarul '1 234 Ft', angolul '1,234 HUF'."""
    return f"{szam(x, en)} {'HUF' if en else 'Ft'}"


def pct(x, tizedes: int = 1, en: bool = False) -> str:
    """Százalék: magyarul tizedesvesszővel (2,5%), angolul ponttal (2.5%)."""
    s = f"{x:.{tizedes}%}"
    return s if en else s.replace(".", ",")


def tized(x, tizedes: int = 1, en: bool = False) -> str:
    """Tizedes tört: magyarul 1,25, angolul 1.25."""
    s = f"{x:.{tizedes}f}"
    return s if en else s.replace(".", ",")


# --- Adatértékek fordítása (csak angol nézetben) ------------------------------
# A státuszokat maga a program állítja elő, ezért valós adatnál is fordítjuk.
STATUSZOK = {
    "Lezárt": "Closed",
    "Folyamatban": "In progress",
    "Nyitott": "Open",
    "Nyitott – oktatás": "Open – training",
    "Nyitott – beszállító": "Open – supplier",
    "Nyitott – technológia": "Open – technology",
    "Nyitott – intézkedés szükséges": "Open – action required",
    "Nincs megadva": "Not specified",
}

# A kitalált (szintetikus) adat értékei (generate_data.py)
SZINTETIKUS = {
    "hibatípus": {
        "Méreteltérés": "Dimensional deviation", "Felületi karc": "Surface scratch",
        "Hiányzó alkatrész": "Missing part", "Hegesztési hiba": "Welding defect",
        "Rossz címke": "Wrong label", "Szennyeződés": "Contamination",
        "Deformáció": "Deformation", "Egyéb": "Other",
    },
    "ok": {
        "Ember": "Man", "Gép": "Machine", "Módszer": "Method", "Anyag": "Material",
        "Mérés": "Measurement", "Környezet": "Environment",
    },
    "termékcsalád": {"A család": "Family A", "B család": "Family B", "C család": "Family C"},
}


def adat_forditas(df: pd.DataFrame, szintetikus: bool) -> pd.DataFrame:
    """Angol nézethez lefordítja a státuszokat és (próbaadatnál) a kategóriákat."""
    df = df.copy()
    df["státusz"] = df["státusz"].map(lambda s: STATUSZOK.get(s, s))
    if szintetikus:
        for oszlop, szotar in SZINTETIKUS.items():
            if oszlop in df.columns:
                df[oszlop] = df[oszlop].map(lambda s, d=szotar: d.get(s, s))
    return df


def nyitott_e(statusz: pd.Series) -> pd.Series:
    """Igaz, ha az ügy nyitott vagy folyamatban van (mindkét nyelven)."""
    return (statusz.str.startswith(("Nyitott", "Open"))
            | statusz.isin(["Folyamatban", "In progress"]))


# Táblázatok oszlopnevei angol nézetben (a számítások magyar oszlopneveiből)
OSZLOPNEVEK_EN = {
    "esetek száma": "cases", "hibás db": "defective qty", "költség (Ft)": "cost (HUF)",
    "első": "first", "utolsó": "last", "hibák": "defects", "munkaszámok": "job numbers",
    "cikkszám": "item no.", "munkaszám": "job no.", "hibatípus": "defect type",
    "hibás_db": "defective qty", "gyártott_db": "produced qty", "FSH": "UCL", "ASH": "LCL",
    "kívül": "out of limits", "bevételezés_db": "received qty",
    "selejt_2o_db": "scrap + 2nd class qty",
    "dátum": "date", "AI eszköz": "AI tool", "feladat / AI kimenet": "task / AI output",
    "ellenőrzött szempontok": "checked aspects", "eredmény": "result", "megjegyzés": "note",
}


def oszlopok(df: pd.DataFrame, en: bool) -> pd.DataFrame:
    """Megjelenítéshez: angol nézetben lefordítja a táblázat oszlopneveit."""
    return df.rename(columns=OSZLOPNEVEK_EN) if en else df
