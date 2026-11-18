# -----------------------------------------------------------------------------
# generate_data.py
# EZ A FÁJL: szintetikus (kitalált, de valósághű) nemmegfelelőségi adatot
# állít elő a data/synthetic/ mappába. Valós adatot NEM használ és NEM küld
# sehova. A fix "seed" (véletlenszám-kezdőérték) miatt minden futtatás
# pontosan ugyanazt az adatot adja.
#
# Futtatás:  python generate_data.py
#
# Két fájlt hoz létre:
#   1) nemmegfelelosegek.csv  – egy sor = egy nemmegfelelőség
#      oszlopok: dátum, hibatípus, ok, termékcsalád, költség, státusz,
#                cikkszám, munkaszám
#   2) gyartott_mennyiseg.csv – havonta hány darabot gyártottunk/ellenőriztünk
#      (erre a p-szabályozókártyához van szükség: hibaarány = hibás / összes)
# -----------------------------------------------------------------------------

from pathlib import Path

import numpy as np
import pandas as pd

# --- Beállítások (ezeket nyugodtan átírhatod) --------------------------------
SEED = 42                      # fix kezdőérték -> mindig ugyanaz az eredmény
START_HONAP = "2024-11-01"     # az első hónap
HONAPOK_SZAMA = 24             # 24 hónap adat
KEZDO_HIBAARANY = 0.030        # kb. 3,0% hibaarány az elején...
VEGSO_HIBAARANY = 0.022        # ...ami enyhén csökken kb. 2,2%-ra a végére
KIMENETI_MAPPA = Path(__file__).parent / "data" / "synthetic"

# Hibatípusok és relatív gyakoriságuk (a Pareto-hatás miatt néhány domináns)
HIBATIPUSOK = {
    "Méreteltérés": 0.30,
    "Felületi karc": 0.22,
    "Hiányzó alkatrész": 0.14,
    "Hegesztési hiba": 0.10,
    "Rossz címke": 0.08,
    "Szennyeződés": 0.07,
    "Deformáció": 0.05,
    "Egyéb": 0.04,
}

# Hibatípusonkénti átlagos költség (Ft / eset) – nagyságrendileg
ATLAGKOLTSEG = {
    "Méreteltérés": 18000,
    "Felületi karc": 6000,
    "Hiányzó alkatrész": 12000,
    "Hegesztési hiba": 35000,
    "Rossz címke": 3000,
    "Szennyeződés": 5000,
    "Deformáció": 25000,
    "Egyéb": 8000,
}

# Gyökérok kategóriák (a klasszikus 6M / Ishikawa szerint)
OKOK = {
    "Ember": 0.25,
    "Gép": 0.22,
    "Módszer": 0.18,
    "Anyag": 0.17,
    "Mérés": 0.10,
    "Környezet": 0.08,
}

TERMEKCSALADOK = {"A család": 0.45, "B család": 0.35, "C család": 0.20}


def general():
    """Létrehozza a két CSV fájlt a data/synthetic/ mappában."""
    rng = np.random.default_rng(SEED)
    honapok = pd.date_range(START_HONAP, periods=HONAPOK_SZAMA, freq="MS")

    # A hibaarány lineárisan csökken az idővel (enyhén csökkenő trend)
    hibaaranyok = np.linspace(KEZDO_HIBAARANY, VEGSO_HIBAARANY, HONAPOK_SZAMA)

    utolso_nap = honapok[-1] + pd.offsets.MonthEnd(0)
    nm_sorok = []        # nemmegfelelőségek
    mennyiseg_sorok = [] # havi gyártott mennyiség

    for i, honap in enumerate(honapok):
        # Havi gyártott/ellenőrzött darabszám: 4500–6500 között ingadozik
        gyartott = int(rng.integers(4500, 6500))
        # Hibás darabok száma: binomiális eloszlás (minden darab p eséllyel hibás)
        hibas = int(rng.binomial(gyartott, hibaaranyok[i]))
        mennyiseg_sorok.append({"hónap": honap.date(), "gyártott_db": gyartott})

        napok_a_honapban = honap.days_in_month
        for _ in range(hibas):
            datum = honap + pd.Timedelta(days=int(rng.integers(0, napok_a_honapban)))
            hibatipus = rng.choice(list(HIBATIPUSOK), p=list(HIBATIPUSOK.values()))
            ok = rng.choice(list(OKOK), p=list(OKOK.values()))
            csalad = rng.choice(list(TERMEKCSALADOK), p=list(TERMEKCSALADOK.values()))
            # Költség: lognormális eloszlás (sok kicsi, néha egy nagy tétel)
            koltseg = ATLAGKOLTSEG[hibatipus] * rng.lognormal(mean=0, sigma=0.5)

            # Státusz: a régi esetek jellemzően már lezártak, a frissek nyitottak
            kor_napban = (utolso_nap - datum).days
            if kor_napban > 90:
                statusz = rng.choice(["Lezárt", "Folyamatban"], p=[0.97, 0.03])
            elif kor_napban > 30:
                statusz = rng.choice(["Lezárt", "Folyamatban", "Nyitott"], p=[0.7, 0.2, 0.1])
            else:
                statusz = rng.choice(["Lezárt", "Folyamatban", "Nyitott"], p=[0.2, 0.3, 0.5])

            nm_sorok.append({
                "dátum": datum.date(),
                "hibatípus": hibatipus,
                "ok": ok,
                "termékcsalád": csalad,
                "költség": round(koltseg, -2),  # 100 Ft-ra kerekítve
                "státusz": statusz,
            })

    nm = pd.DataFrame(nm_sorok)

    # Cikkszám és munkaszám (az "Ismétlődések" fülhöz). Külön véletlenszám-
    # generátort használunk, így a fenti oszlopok értékei nem változnak.
    rng2 = np.random.default_rng(SEED + 1)
    cikkek = [f"C{rng2.integers(10000, 99999)}-{i:02d}" for i in range(60)]
    cikk_suly = rng2.pareto(1.5, len(cikkek)) + 0.1  # néhány cikk sokszor hibás
    nm["cikkszám"] = rng2.choice(cikkek, size=len(nm), p=cikk_suly / cikk_suly.sum())
    # Munkaszám: havonta kb. 120 gyártási rendelés közül; egy rendelés többször is hibázhat
    nm["munkaszám"] = [f"MS{pd.Timestamp(d):%y%m}-{rng2.integers(1, 120):04d}" for d in nm["dátum"]]

    nm = nm.sort_values("dátum").reset_index(drop=True)
    mennyiseg = pd.DataFrame(mennyiseg_sorok)

    KIMENETI_MAPPA.mkdir(parents=True, exist_ok=True)
    # utf-8-sig: így az Excel is helyesen mutatja az ékezeteket
    nm.to_csv(KIMENETI_MAPPA / "nemmegfelelosegek.csv", index=False, encoding="utf-8-sig")
    mennyiseg.to_csv(KIMENETI_MAPPA / "gyartott_mennyiseg.csv", index=False, encoding="utf-8-sig")

    print(f"Kész! {len(nm)} nemmegfelelőség, {len(mennyiseg)} hónap -> {KIMENETI_MAPPA}")


# Ez a rész csak akkor fut le, ha közvetlenül indítod a fájlt (python generate_data.py)
if __name__ == "__main__":
    general()
