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

The workbook contains:

- `Summary`: run totals and source info
- `All Rates`: one row per bank per article date
- `Latest Table`: the newest scraped date in the same layout as the article
- `Articles`: article-level status and error details
