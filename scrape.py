"""
ASX Energy NZ Futures Scraper
Scrapes Base Month and Base Quarter settle prices for Otahuhu and Benmore
and appends a row per contract per day to an Excel file.
"""

import sys
import time
import re
from datetime import datetime, date
from pathlib import Path
from zoneinfo import ZoneInfo

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

URL = "https://www.asxenergy.com.au/futures_nz"
EXCEL_FILE = Path(__file__).parent / "asx_nz_futures.xlsx"

TARGET_SECTIONS = {"Base Month", "Base Quarter"}

NODE_MAP = {
    "Otahuhu": "OTA2201",
    "Benmore": "BEN2201",
}


def get_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    try:
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)
    except Exception:
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=opts
            )
        except Exception:
            driver = webdriver.Chrome(options=opts)
    return driver


def clean_heading(text):
    """Strip trailing junk chars appended by the site's JS (e.g. 'Base MonthED' -> 'Base Month')."""
    return re.sub(r'[A-Z]{1,3}$', '', text.strip()).strip()


def scrape() -> list[dict]:
    """
    Returns a list of dicts: {node, period_type, time_period, price}

    Page structure (confirmed from live DOM):
      <h2>Otahuhu</h2>
        <h3>Base MonthED</h3>
        <table>...</table>   columns: Contract | Bid Size | Bid | Ask | Ask Size | High | Low | Last | +/- | Vol | OpenInt | OpenInt +/- | Settle
        <h3>Base QuarterEA</h3>
        <table>...</table>
      <h2>Benmore</h2>
        ...
    """
    driver = get_driver()
    records = []

    try:
        driver.get(URL)

        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h2"))
            )
        except TimeoutException:
            print("Timed out waiting for page to load", file=sys.stderr)
            return records

        time.sleep(5)

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(driver.page_source, "html.parser")

        current_node = None
        current_section = None

        for elem in soup.find_all(["h2", "h3", "table"]):

            if elem.name == "h2":
                text = elem.get_text(strip=True)
                if "Otahuhu" in text:
                    current_node = "Otahuhu"
                elif "Benmore" in text:
                    current_node = "Benmore"
                else:
                    current_node = None
                current_section = None
                continue

            if elem.name == "h3":
                if current_node is None:
                    current_section = None
                    continue
                heading = clean_heading(elem.get_text(strip=True))
                current_section = heading if heading in TARGET_SECTIONS else None
                continue

            if elem.name == "table":
                if current_node is None or current_section is None:
                    continue

                rows = elem.find_all("tr")
                if len(rows) < 2:
                    continue

                header_cells = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]

                contract_idx = 0
                settle_idx = len(header_cells) - 1
                for i, h in enumerate(header_cells):
                    if "contract" in h.lower():
                        contract_idx = i
                    if "settle" in h.lower():
                        settle_idx = i

                for row in rows[1:]:
                    cells = row.find_all(["td", "th"])
                    if len(cells) <= max(contract_idx, settle_idx):
                        continue

                    contract = cells[contract_idx].get_text(strip=True)
                    settle = cells[settle_idx].get_text(strip=True)

                    if not contract or not settle:
                        continue
                    # Skip rows with no 4-digit year — not a contract row
                    if not re.search(r"\d{4}", contract):
                        continue

                    settle_clean = settle.replace(",", "").strip()
                    price = None
                    if settle_clean not in ("-", "", "N/A", "n/a"):
                        try:
                            price = float(settle_clean)
                        except ValueError:
                            pass

                    records.append({
                        "node": NODE_MAP.get(current_node, current_node),
                        "period_type": current_section,
                        "time_period": contract,
                        "price": price,
                    })

                current_section = None

    finally:
        driver.quit()

    return records


def append_to_excel(records: list[dict], execution_date: datetime):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    data_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_even = PatternFill("solid", start_color="EBF3FB")
    fill_odd = PatternFill("solid", start_color="FFFFFF")

    exec_date = execution_date.date()

    # Remove any existing rows for today to prevent duplicates
    rows_to_delete = []
    for row in ws.iter_rows(min_row=3):
        cell_value = row[0].value
        if cell_value is None:
            continue
        # Normalise to date — Excel may store as datetime or date
        if isinstance(cell_value, datetime):
            row_date = cell_value.date()
        elif isinstance(cell_value, date):
            row_date = cell_value
        else:
            continue
        if row_date == exec_date:
            rows_to_delete.append(row[0].row)

    for row_num in reversed(rows_to_delete):
        ws.delete_rows(row_num)

    if rows_to_delete:
        print(f"Removed {len(rows_to_delete)} duplicate rows for {exec_date}")

    last_row = ws.max_row
    insert_row = 3 if last_row < 3 else last_row + 1

    for record in records:
        fill = fill_even if (insert_row % 2 == 0) else fill_odd
        cells_data = [
            exec_date,
            record["node"],
            record["period_type"],
            record["time_period"],
            record["price"],
        ]

        for col, value in enumerate(cells_data, 1):
            cell = ws.cell(row=insert_row, column=col, value=value)
            cell.font = data_font
            cell.border = border
            cell.fill = fill
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if col == 1:
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col == 5 and value is not None:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")

        insert_row += 1

    wb.save(EXCEL_FILE)
    print(f"Appended {len(records)} records for {exec_date}")


def main():
    nzt = ZoneInfo("Pacific/Auckland")
    execution_dt = datetime.now(tz=nzt)
    print(f"Scraping at {execution_dt.strftime('%Y-%m-%d %H:%M:%S %Z')}")

    records = scrape()

    if not records:
        print("No records scraped — check page structure or network access.", file=sys.stderr)
        sys.exit(1)

    print(f"Scraped {len(records)} records:")
    for r in records:
        print(f"  {r['node']:10} | {r['period_type']:14} | {r['time_period']:12} | {r['price']}")

    append_to_excel(records, execution_dt)


if __name__ == "__main__":
    main()
