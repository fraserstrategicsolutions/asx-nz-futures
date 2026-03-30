"""
ASX Energy NZ Futures Scraper
Scrapes Base Month and Base Quarter settle prices for Otahuhu and Benmore
and appends a row per contract per day to an Excel file.
"""

import sys
import time
import re
from datetime import datetime
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

# Sections we want to capture (matched against table headings on the page)
TARGET_SECTIONS = {"Base Month", "Base Quarter"}

# Node labels as they appear in section headings
NODE_OTAHUHU = "Otahuhu"
NODE_BENMORE = "Benmore"


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
        # GitHub Actions / system chromedriver
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)
    except Exception:
        # Fall back to webdriver-manager if available
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=opts
            )
        except Exception:
            driver = webdriver.Chrome(options=opts)
    return driver


def scrape() -> list[dict]:
    """
    Returns a list of dicts:
      {node, section_type, time_period, settle_price}
    """
    driver = get_driver()
    records = []

    try:
        driver.get(URL)

        # Wait for at least one table to appear
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table"))
            )
        except TimeoutException:
            print("Timed out waiting for tables to load", file=sys.stderr)
            return records

        # Additional wait for JS to fully render all tables
        time.sleep(5)

        page_source = driver.page_source

        # Parse with BeautifulSoup for reliability
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(page_source, "html.parser")

        # The page structure: section headings precede each table.
        # Heading text contains both node and product type, e.g.:
        #   "Otahuhu Base Month" / "Benmore Base Quarter"
        # We walk all headings (h2/h3/h4/strong/div with relevant text)
        # and find the next table sibling.

        current_node = None
        current_section = None

        # Collect all relevant elements in document order
        # Strategy: find all <h2>/<h3>/<h4> and <table> tags, process in order
        elements = soup.find_all(["h2", "h3", "h4", "h5", "table", "div"])

        for elem in elements:
            text = elem.get_text(strip=True)

            # Detect section headings
            for node in [NODE_OTAHUHU, NODE_BENMORE]:
                for section in TARGET_SECTIONS:
                    pattern = re.compile(
                        rf"{re.escape(node)}.*{re.escape(section)}|"
                        rf"{re.escape(section)}.*{re.escape(node)}",
                        re.IGNORECASE
                    )
                    if pattern.search(text) and len(text) < 100:
                        current_node = node
                        current_section = section
                        break

            # Parse tables when we have a valid context
            if elem.name == "table" and current_node and current_section:
                rows = elem.find_all("tr")
                if not rows:
                    continue

                # Find header row to locate "Contract" and "Settle" columns
                header_row = rows[0]
                headers = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])]

                contract_idx = None
                settle_idx = None

                for i, h in enumerate(headers):
                    if re.search(r"contract", h, re.IGNORECASE):
                        contract_idx = i
                    if re.search(r"settle", h, re.IGNORECASE):
                        settle_idx = i

                if contract_idx is None or settle_idx is None:
                    # Try alternate column detection — some pages use positional layout
                    # Typically: Contract | Open | High | Low | Settle | ...
                    if len(headers) >= 5:
                        contract_idx = 0
                        settle_idx = 4
                    else:
                        continue

                for row in rows[1:]:
                    cells = row.find_all(["td", "th"])
                    if len(cells) <= max(contract_idx, settle_idx):
                        continue

                    contract = cells[contract_idx].get_text(strip=True)
                    settle = cells[settle_idx].get_text(strip=True)

                    # Skip empty or header-repeat rows
                    if not contract or not settle:
                        continue
                    if re.search(r"contract|settle", contract, re.IGNORECASE):
                        continue

                    # Clean settle price — remove commas, handle dashes (no trade)
                    settle_clean = settle.replace(",", "").strip()
                    if settle_clean in ("-", "", "N/A", "n/a"):
                        settle_price = None
                    else:
                        try:
                            settle_price = float(settle_clean)
                        except ValueError:
                            settle_price = None

                    records.append({
                        "node": current_node,
                        "section_type": current_section,
                        "time_period": f"{current_section} – {contract}",
                        "settle_price": settle_price,
                    })

                # Reset so we don't re-use heading for a second table
                current_node = None
                current_section = None

    finally:
        driver.quit()

    return records


def append_to_excel(records: list[dict], execution_date: datetime):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    # Style for data rows
    data_font = Font(name="Arial", size=10)
    thin = openpyxl.styles.Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Alternate row fill
    fill_even = PatternFill("solid", start_color="EBF3FB")
    fill_odd = PatternFill("solid", start_color="FFFFFF")

    # Find the next empty row (skip header + note rows)
    # Row 1 = header, row 2 = note. Data starts at row 3.
    # Find actual last row with data
    last_row = ws.max_row
    # Determine next insert row
    if last_row < 3:
        insert_row = 3
    else:
        insert_row = last_row + 1

    date_str = execution_date.strftime("%Y-%m-%d")
    exec_date = execution_date.date()

    for record in records:
        row_idx = insert_row
        fill = fill_even if (row_idx % 2 == 0) else fill_odd

        cells_data = [
            exec_date,
            record["node"],
            record["time_period"],
            record["settle_price"],
        ]

        for col, value in enumerate(cells_data, 1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = data_font
            cell.border = border
            cell.fill = fill
            cell.alignment = Alignment(horizontal="left", vertical="center")

            # Format date column
            if col == 1:
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            # Format price column
            if col == 4 and value is not None:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")

        insert_row += 1

    wb.save(EXCEL_FILE)
    print(f"Appended {len(records)} records for {date_str}")


def main():
    nzt = ZoneInfo("Pacific/Auckland")
    execution_dt = datetime.now(tz=nzt)

    print(f"Scraping at {execution_dt.strftime('%Y-%m-%d %H:%M:%S %Z')}")

    records = scrape()

    if not records:
        print("No records scraped — check page structure or network access.", file=sys.stderr)
        sys.exit(1)

    print(f"Scraped {len(records)} records")
    for r in records:
        print(f"  {r['node']:12} | {r['time_period']:35} | {r['settle_price']}")

    append_to_excel(records, execution_dt)


if __name__ == "__main__":
    main()
