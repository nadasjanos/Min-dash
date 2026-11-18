# -----------------------------------------------------------------------------
# minta_keszit.py
# EZ A FÁJL: kitöltési mintát készít a minta/ mappába azoknak, akik a saját
# adataikkal próbálnák ki a dashboardot. A mintában KITALÁLT adat van
# (a szintetikus adatból, fix seed), valós adatot nem használ.
#
# Futtatás:  python minta_keszit.py   (előtte: python generate_data.py)
#
# Két fájlt hoz létre:
#   1) minta/NMF_minta.xlsx – pontosan olyan oszlopnevekkel, amilyeneket a
#      valos_adat.py vár. A saját adatod ilyen szerkezetben másold be.
#   2) minta/gyartott_mennyiseg.csv – havi gyártott darabszám a p-kártyához.
# -----------------------------------------------------------------------------

from pathlib import Path

import numpy as np
import pandas as pd

import valos_adat

SEED = 7
HONAPOK = 12          # az utolsó 12 hónap a szintetikus adatból
MAPPA = Path(__file__).parent
FORRAS = MAPPA / "data" / "synthetic"
CEL = MAPPA / "minta"

STATUSZ_FEJLEC = ("n=nincs teendő f=folyamatban o=oktatás szükséges b=beszállítói "
                  "intézkedés t=technológia módosítás i=egyéb intézkedés")
KOVETKEZMENYEK = {"Egyéb javítás": 0.45, "Felület javítás": 0.2, "Teljes selejt": 0.25,
                  "Másodosztály": 0.1}
MUVELETEK = ["Szabás", "Hajlítás", "Hegesztés", "Festés", "Szerelés", "Csomagolás"]
GEPEK = ["Lézer 1", "Lézer 2", "Élhajlító", "Hegesztő robot", "Festősor", "Kézi"]


def keszit():
    rng = np.random.default_rng(SEED)
    nm = pd.read_csv(FORRAS / "nemmegfelelosegek.csv", encoding="utf-8-sig", parse_dates=["dátum"])
    menny = pd.read_csv(FORRAS / "gyartott_mennyiseg.csv", encoding="utf-8-sig", parse_dates=["hónap"])
    honapok = sorted(menny["hónap"])[-HONAPOK:]
    nm = nm[nm["dátum"] >= honapok[0]].reset_index(drop=True)
    menny = menny[menny["hónap"].isin(honapok)]

    # Státusz -> intézkedési kód (n/f/o/b/t/i), ahogy a valós táblában
    kod = nm["státusz"].map({"Lezárt": "n", "Folyamatban": "f"})
    nyitott = kod.isna()
    kod[nyitott] = rng.choice(list("obti"), size=int(nyitott.sum()))

    o = valos_adat.OSZLOPOK
    v = valos_adat.OPCIONALIS_OSZLOPOK
    excel = pd.DataFrame({
        o["dátum"]: nm["dátum"].dt.date,
        v["munkaszám"]: nm["munkaszám"],
        v["cikkszám"]: nm["cikkszám"],
        o["termékcsalád"]: nm["termékcsalád"].str.replace("család", "Vevő", regex=False),
        o["hibatípus"]: nm["hibatípus"],
        o["ok"]: nm["ok"],
        v["keletkezés_helye"]: rng.choice(MUVELETEK, size=len(nm)),
        v["gép"]: rng.choice(GEPEK, size=len(nm)),
        o["nmf_db"]: rng.integers(1, 4, size=len(nm)),
        v["következmény"]: rng.choice(list(KOVETKEZMENYEK), size=len(nm),
                                      p=list(KOVETKEZMENYEK.values())),
        o["költség"]: nm["költség"],
        STATUSZ_FEJLEC: kod,
    })
    assert STATUSZ_FEJLEC.startswith(o["státusz"])  # a program így ismeri fel

    CEL.mkdir(exist_ok=True)
    excel.to_excel(CEL / "NMF_minta.xlsx", index=False)
    menny.assign(hónap=menny["hónap"].dt.date).to_csv(
        CEL / "gyartott_mennyiseg.csv", index=False, encoding="utf-8-sig")
    print(f"Kész! {len(excel)} sor, {len(menny)} hónap -> {CEL}")


if __name__ == "__main__":
    keszit()
