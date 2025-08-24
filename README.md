# Portfolio Tracker — Professional Charts (Excel + Python)

<img width="512" height="269" alt="images:charts_preview2" src="https://github.com/user-attachments/assets/a0d0be59-bd48-49c9-9845-82b79d9ac7b9" />

<img width="513" height="271" alt="images:charts_preview1" src="https://github.com/user-attachments/assets/ec580f97-37c3-4ed7-8e67-01bbf6367fd9" />

Automated a portfolio analytics workbook using Python (`openpyxl`) and Excel.
Generates two professional charts on a dedicated **Charts** sheet:

1. **Cumulative Portfolio Growth** (line)
2. **Average Asset Returns** (bar)

Includes a clean repo with script and MIT license for easy reuse.

---

## Quick Start

1. **Download** the Excel and script from this repo's release or the files directly.
2. Install Python 3.9+ and dependencies:
   ```bash
   pip install openpyxl
   ```
3. Put `chart_updater.py` in the same folder as `Portfolio_Tracker_Professional.xlsx`.
4. Run:
   ```bash
   python chart_updater.py
   ```
   The script will (re)create the **Charts** sheet and populate both charts.
   If your workbook lacks data, it auto-fills example data so the visuals are never empty.

> Tip: Add a screenshot of the Charts sheet at `images/charts_preview.png` and reference it here.

---

## Files

- `Portfolio_Tracker_Professional.xlsx` — Excel workbook with example data and charts.
- `chart_updater.py` — Python script that regenerates charts with safe fallbacks.
- `LICENSE` — MIT License.
- `README.md` — This page.

---

## How it works

- **openpyxl** creates a **LineChart** for cumulative portfolio value over time from the `Data` sheet.
- A **BarChart** displays average returns by asset from the `Returns` sheet.
- The script is idempotent: it removes/replaces charts on each run to keep things tidy.

---

## Why this is useful

- Portfolio snapshots for interviews or performance reviews
- Rapid visuals without opening Excel's chart UI
- Clean baseline you can expand: add risk metrics, sector splits, etc.

---

## License

MIT — do whatever you like, attribution appreciated.
