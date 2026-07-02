# Vietnamnet LSTK scraper

Small Python app for scraping Vietnamnet daily bank-deposit interest-rate
articles and exporting the tables to Excel.

## Setup

Use the local virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Run

Open the PyQt6 menu:

```powershell
.\.venv\Scripts\python.exe pyqt6_menu.py
```

Or on Windows, just use:

```powershell
run_scraper.bat
```

Install everything on a new PC with:

```powershell
install_requirements.bat
```

Scrape the tag archive safely up to 500 pages:

```powershell
.\.venv\Scripts\python.exe app.py --pages 500 --output outputs\lstk_vietnamnet.xlsx
```

Useful options:

```powershell
.\.venv\Scripts\python.exe app.py --menu
.\.venv\Scripts\python.exe app.py --pages 5 --max-articles 20
.\.venv\Scripts\python.exe app.py --date-from 2026-06-01 --date-to 2026-06-26
.\.venv\Scripts\python.exe app.py --refresh
```

Legacy free-text search for older articles before Vietnamnet table coverage:

```powershell
.\.venv\Scripts\python.exe app.py --legacy --legacy-from-year 2015 --legacy-to-year 2023
```

The PyQt6 menu also has a checkbox for older free-text news. This mode is best-effort:
it searches the configured news domains, reads article text, and extracts bank/term/rate
pairs into the same Excel layout with source URLs for review.

Legacy sources currently include VietnamFinance, Timo, InfoNet/Vietnamnet, Thanh Nien,
Thanh Tra, VnEconomy, The Saigon Times, Tin nhanh chung khoan, Bao Dau Tu, Dan Tri,
CafeF, Lao Dong, Tuoi Tre, VietnamPlus, and Thoi Bao Tai Chinh Viet Nam. Articles that
only say a broad range such as "from 4.2% to 6.5%" without a specific term are logged
but skipped, so the workbook does not contain guessed rates.

The workbook contains:

- `Summary`: run totals and source info
- `All Rates`: one row per bank per article date
- `Latest Table`: the newest scraped date in the same layout as the article
- `Articles`: article-level status and error details
