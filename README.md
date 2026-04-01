# ASX NZ Electricity Futures Scraper

Automated daily scraper for ASX Energy NZ electricity futures settle prices.

## What it does

Scrapes [asxenergy.com.au/futures_nz](https://www.asxenergy.com.au/futures_nz) every weekday at ~7pm NZT and appends rows to `asx_nz_futures.xlsx`.

## Excel output schema

| Column | Description |
|---|---|
| Execution Date | Date the scrape ran (NZT) |
| Node | `Otahuhu` or `Benmore` |
| Time Period | e.g. `Base Month – Apr 2026`, `Base Quarter – Jun 2026` |
| Settle Price ($/MWh) | Settlement price from the Settle column |

One row per contract per node per day. Base Month and Base Quarter contracts are captured automatically — as months/quarters roll off and new ones appear, they will be picked up without any code changes.

## Schedule

GitHub Actions runs the workflow on weekdays at **07:00 UTC**, which corresponds to:
- **7:00 PM NZST** (UTC+12, April–September)
- **8:00 PM NZDT** (UTC+13, October–March)

To change the schedule, edit `.github/workflows/scrape.yml`.

## Running manually

```bash
pip install -r requirements.txt
python scrape.py
```

Requires Chrome/Chromium and ChromeDriver installed.

## How it works

The ASX Energy page is JavaScript-rendered, so a headless Chrome browser (Selenium) is used to load the full page before parsing. The scraper:

1. Loads the page and waits for tables to render
2. Identifies sections by heading (e.g. "Otahuhu Base Month")
3. Extracts contract name and Settle price from each relevant table
4. Appends new rows to the Excel file
5. Commits and pushes the updated file via GitHub Actions

## Troubleshooting

- **No records scraped**: The page structure may have changed. Inspect the page and update the heading/table detection logic in `scrape.py`.
- **Push conflicts**: The workflow uses `fetch-depth: 1` and only commits when changes exist, minimising conflicts.
- **DST gap**: During the NZ DST transition weekends, the scrape may run at 8pm instead of 7pm. This is cosmetic only.
- 
