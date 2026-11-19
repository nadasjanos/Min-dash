<!--
README_EN.md
EZ A FÁJL: a README.md angol változata (külföldi olvasóknak).
-->

# Quality dashboard (Streamlit)

A simple web dashboard for analysing nonconformities (NCs). It runs on your own computer.

> 🌐 **Try it without installing (synthetic data):** https://minosegugyi-dashboard.streamlit.app/?lang=en
> To use your own data, download the program and run it locally. See [Try it with your own data](#try-it-with-your-own-data).
>
> 🇭🇺 Magyar leírás: [README.md](README.md)

**Language:** choose **Magyar** or **English** at the top of the sidebar. The PowerPoint/PDF report is created in the selected language too. Categories and statuses of the synthetic data are translated. Values from your own Excel file (customer names, defect descriptions) are shown as they are. Filters are cleared when you switch language.

## What does it do?

| Tab | What it shows |
|---|---|
| 📈 **Monthly trend** | Number and cost of nonconformities per month, with a trend line. |
| 📊 **Pareto chart** | Ranking by a chosen breakdown (defect type, cause, customer, operation, machine, consequence), by count or cost, with cumulative % and an 80% line. Above 15 categories the rest is grouped into "Other". Tip: filter to one defect type and break it down by cause. |
| 🔁 **Repeat issues** | Ranked by *number of cases*, *estimated cost* or *defective qty*. A **🏆 Top 3** summary for items and job numbers shows their share of the total and how many items make up 80%. Below: a chart of the 15 most frequent items coloured by defect type, plus tables of *same item, same defect*, *repeat items* and *repeat job numbers*. A slider sets the minimum number of occurrences. |
| 🎯 **p control chart** | Monthly defect rate (defective qty / produced qty) with control limits; months outside the limits are marked with a red X. Choose **Laney p′** (default) or **classic p** (p̄ ± 3σ). With real data you can also choose what counts as defective: **all nonconformities** or **only scrap + 2nd class** (final loss). With large volumes (thousands per month) the classic limits are very narrow and almost every month falls outside. Laney p′ adjusts the limits to the real month-to-month variation (σz), so it flags only the truly unusual months. |
| 📝 **AI verification log** | A form to record what you checked manually in an AI-generated output. Saved locally to `data/ai_naplo.csv`. |

**Sidebar filters:** month (range slider), customer, defect type, and (real data only) operation, machine and consequence. An empty filter means **All**. *Clear filters* resets everything. Filters apply to the trend, the Pareto chart and the key figures. Only the month filter applies to the p chart, because the produced quantity is known only as a monthly total.

### Report (PowerPoint + PDF)

At the bottom of the sidebar, **📄 Create report** builds a report with the **current filters**. Then download it as **PowerPoint (.pptx)**, which you can edit, or as **PDF**.

Contents (9–10 slides):
- title page and key figures
- monthly trend
- Pareto charts by defect type and by cause
- top 3 items and job numbers
- repeating item–defect pairs
- p chart, plus a final-loss p chart for real data
- an empty action table

Each chart has automatic, factual findings. Add your own assessment in PowerPoint.

The report is created on your computer and nothing is sent anywhere. With real data it contains customer names and costs, so check who you send it to.

## Data and privacy

- By default the app uses **synthetic (made-up) data** from `data/synthetic/`.
- **Your own data** goes into the `data/real/` folder. The app reads it **only** when you switch on *"Use real data"* in the sidebar.
- `data/real/` and the AI log are listed in `.gitignore`, so Git never uploads them.
- The program uploads nothing to the internet:
  - Streamlit usage statistics are turned off (`.streamlit/config.toml`).
  - `inditas.bat` starts the app on `localhost` only, so it is reachable only from your own computer.

## Try it with your own data

1. **Download:** on the GitHub page click the green **Code** button → **Download ZIP**, and unzip it.
2. **First start with synthetic data:** double-click `inditas.bat` (Windows). It installs everything it needs and opens the dashboard in your browser.
3. **Your NC table:** open `minta/NMF_minta.xlsx` (the template). Your Excel file needs these **column names, exactly as written** (they are in Hungarian):

   | Excel column | Required? | Meaning |
   |---|---|---|
   | `Dátum` | yes | date of the nonconformity |
   | `Nem megfelelőség jellege` | yes | defect type (e.g. dimensional deviation) |
   | `Nem megfelelőség oka` | yes | cause (e.g. machine, operator) |
   | `Vevő` | yes | customer or product family |
   | `Becsült költségvonzat Ft` | yes | estimated cost |
   | `n=nincs teendő …` | yes | action code (see below) |
   | `Nmf db` | yes | defective quantity |
   | `Munkaszám`, `Cikkszám` | no | job number, item number (for Repeat issues) |
   | `Keletkezés helye (Művelet)`, `NMF keletkezés helye (gép)`, `Következmény` | no | operation, machine, consequence. The consequence values `Teljes selejt` (scrap) and `Másodosztály` (2nd class) are used by the final-loss p chart. |

   **Action codes:**

   | Code | Status in the app |
   |---|---|
   | `n` | Closed |
   | `f` | In progress |
   | `o` | Open – training |
   | `b` | Open – supplier |
   | `t` | Open – technology |
   | `i` | Open – action needed |
   | empty | Not specified |

   If your columns have different names, you can rename them in Excel. Or you can change the names in the `OSZLOPOK` dictionary at the top of `valos_adat.py`. The app tells you if a column is missing.
4. **Produced quantity (for the p chart):** copy the layout of `minta/gyartott_mennyiseg.csv` (columns `hónap` = first day of the month, `gyártott_db` = units produced), and save it under the same name in `data/real/`. Without it, everything except the p chart still works.
5. Copy your Excel file into `data/real/`, restart the program, and switch on **"Use real data"** in the sidebar.

Text values that differ only in spacing, upper/lower case, accents or punctuation are merged into one name, using the most frequent spelling. Your Excel file is not changed.

## Manual installation and start

1. Install Python 3.10 or newer: https://www.python.org/downloads/
2. Open PowerShell in the project folder.
3. Create a virtual environment and install the packages:

```bash
python -m venv .venv
```

```bash
.venv\Scripts\Activate.ps1
```

```bash
pip install -r requirements.txt
```

4. Create the synthetic data (once) and start the dashboard:

```bash
python generate_data.py
```

```bash
streamlit run app.py --server.address localhost
```

The browser opens `http://localhost:8501`. To stop, press **Ctrl+C** in the terminal, or close the `inditas.bat` window.

## Files

| File | Purpose |
|---|---|
| `inditas.bat` | Double-click launcher: installs what is needed and starts the app. |
| `app.py` | The dashboard. |
| `generate_data.py` | Creates synthetic data (24 months, fixed seed = 42). |
| `valos_adat.py` | Reads your Excel file and maps its columns. |
| `szamitasok.py` | Shared calculations (trend, Pareto, p chart, repeats), used by both the screen and the report. |
| `beszamolo.py` | Builds the PowerPoint and PDF report. |
| `forditas.py` | Hungarian / English: number and currency formats, translations. |
| `minta_keszit.py`, `minta/` | Template files (`NMF_minta.xlsx`, `gyartott_mennyiseg.csv`). |
| `github_csomag.py` | Builds an upload-ready folder without real data. |
| `requirements.txt` | Required Python packages. |

The code comments are in Hungarian.
