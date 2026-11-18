# -----------------------------------------------------------------------------
# github_csomag.py
# EZ A FÁJL: összemásolja egy külön mappába (github_csomag/) CSAK azokat a
# fájlokat, amelyek nyugodtan felkerülhetnek a GitHubra: kód, leírás,
# szintetikus adat, kitöltési minta. Valós adat (data/real/), AI-napló,
# .venv és minden más kimarad. Semmit nem tölt fel – csak helyben másol.
#
# Futtatás:  python github_csomag.py
# Utána a github_csomag/ mappa TARTALMÁT töltsd fel a GitHubra.
# -----------------------------------------------------------------------------

import shutil
from pathlib import Path

MAPPA = Path(__file__).parent
CEL = MAPPA / "github_csomag"

# Engedélyezett fájlok – csak ezek kerülnek a csomagba (minden más kimarad)
FAJLOK = [
    "app.py", "szamitasok.py", "valos_adat.py", "beszamolo.py",
    "generate_data.py", "minta_keszit.py", "github_csomag.py",
    "inditas.bat", "requirements.txt", ".gitignore",
    "README.md", "LEIRAS.md",
    ".streamlit/config.toml",
    "data/synthetic/nemmegfelelosegek.csv", "data/synthetic/gyartott_mennyiseg.csv",
    "minta/NMF_minta.xlsx", "minta/gyartott_mennyiseg.csv",
]
# Üres mappák, amelyeknek létezniük kell (a .gitkeep miatt a GitHub is megtartja)
URES_MAPPAK = ["data/real"]


def csomagol():
    if CEL.exists():
        shutil.rmtree(CEL)  # mindig tisztán, elölről készül
    for nev in FAJLOK:
        forras = MAPPA / nev
        if not forras.exists():
            print(f"FIGYELEM: hiányzik, kimarad: {nev}")
            continue
        cel = CEL / nev
        cel.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(forras, cel)
    for nev in URES_MAPPAK:
        (CEL / nev).mkdir(parents=True, exist_ok=True)
        (CEL / nev / ".gitkeep").write_text(
            "Ide jön a saját NMF Excel-táblád. Ez a mappa nem kerül fel a GitHubra.\n",
            encoding="utf-8")

    # Biztonsági ellenőrzés: a csomagban nem lehet semmi a data/real/-ből
    tiltott = [p for p in (CEL / "data" / "real").iterdir() if p.name != ".gitkeep"]
    assert not tiltott, f"Valós adat került a csomagba: {tiltott}"

    print(f"Kész: {CEL}")
    for p in sorted(CEL.rglob("*")):
        if p.is_file():
            print("  ", p.relative_to(CEL).as_posix())


if __name__ == "__main__":
    csomagol()
