import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import time
import pytz

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

DATASET_URL = "https://www.asxenergy.com.au/futures_nz/dataset"
EXCEL_FILE = "asx_nz_futures.xlsx"
LOCATIONS = ["Otahuhu", "Benmore"]

# Product types we want to capture (partial match, case-insensitive)
PRODUCT_TYPES = ["Base Month", "Base Quarter"]

# Per-contract fields captured from the table
CONTRACT_FIELDS = ["Bid Size", "Bid", "Ask", "Ask Size", "High", "Low", "Last", "+/-", "Vol"]


def fetch_html_with_selenium():
    """Use headless Chrome to fully render the JS-driven page."""
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(DATASET_URL)
        # Wait until at least one data table row appears
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )
        # Extra pause to let all tables finish rendering
        time.sleep(3)
        html = driver.page_source
    finally:
        driver.quit()

    return html


def parse_data(html):
    """
    Parse the rendered HTML and return:
      last_update (str)
      data: dict  {location: {product_name: [ {Contract, Bid Size, ...}, ... ] }}
    """
    soup = BeautifulSoup(html, "html.parser")

    # --- Last update timestamp ---
    last_div = soup.find("div", id="last")
    last_update = ""
    if last_div:
        date_str = last_div.get("data-date", "")
        time_str = last_div.get("data-time", "")
        last_update = f"{date_str} {time_str}".strip()

    results = {}

    for section in soup.find_all("div", class_="market-dataset"):
        h2 = section.find("h2")
        if not h2:
            continue
        location_name = h2.get_text(strip=True)
        if location_name not in LOCATIONS:
            continue

        location_data = {}

        for product_section in section.find_all("div", class_="dataset"):
            # Product name from the nearest h3
            parent = product_section.parent
            h3 = parent.find("h3") if parent else None
            if not h3:
                continue
            raw_name = h3.get_text(separator=" ", strip=True)
            # Strip trailing button text / "ED" suffix
            product_name = raw_name.replace("ED", "").split("\n")[0].strip()

            # Only capture the product types we care about
            if not any(pt.lower() in product_name.lower() for pt in PRODUCT_TYPES):
                continue

            table = product_section.find("table")
            if not table:
                continue

            # Parse column headers from thead
            thead = table.find("thead")
            if thead:
                header_cells = [th.get_text(strip=True) for th in thead.find_all("th")]
            else:
                header_cells = CONTRACT_FIELDS  # fallback

            rows = []
            tbody = table.find("tbody")
            if tbody:
                for tr in tbody.find_all("tr"):
                    cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                    if cells and any(c for c in cells):
                        row_dict = {}
                        for i, val in enumerate(cells):
                            col_name = header_cells[i] if i < len(header_cells) else f"Col{i}"
                            row_dict[col_name] = val
                        rows.append(row_dict)

            if rows:
                location_data[product_name] = rows

        results[location_name] = location_data

    return last_update, results


def get_nzt_timestamp():
    nzt = pytz.timezone("Pacific/Auckland")
    return datetime.now(nzt).strftime("%Y-%m-%d %H:%M:%S NZT")


def build_flat_columns(data):
    """
    Build the ordered list of column headers for the flat, one-row-per-day format.
    Structure: Scraped At | ASX Last Update | <Location>_<Product>_<Contract>_<Field> ...
    """
    cols = ["Scraped At", "ASX Last Update"]
    for location in LOCATIONS:
        if location not in data:
            continue
        for product_name, rows in data[location].items():
            for row_dict in rows:
                contract = row_dict.get("Contract", "")
                for field in CONTRACT_FIELDS:
                    col = f"{location} | {product_name} | {contract} | {field}"
                    if col not in cols:
                        cols.append(col)
    return cols


def write_to_excel(last_update, data):
    scraped_at = get_nzt_timestamp()

    # ------------------------------------------------------------------
    # Build the flat data dict for this scrape: column_name -> value
    # ------------------------------------------------------------------
    flat_row = {"Scraped At": scraped_at, "ASX Last Update": last_update}

    for location in LOCATIONS:
        if location not in data:
            continue
        for product_name, rows in data[location].items():
            for row_dict in rows:
                contract = row_dict.get("Contract", "")
                for field in CONTRACT_FIELDS:
                    col = f"{location} | {product_name} | {contract} | {field}"
                    flat_row[col] = row_dict.get(field, "")

    # ------------------------------------------------------------------
    # Load or create workbook with a single "Data" sheet
    # ------------------------------------------------------------------
    sheet_name = "Data"

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

    # ------------------------------------------------------------------
    # Determine existing headers (row 1) or create them
    # ------------------------------------------------------------------
    # Read existing headers from row 1
    existing_headers = []
    for cell in ws[1]:
        if cell.value is not None:
            existing_headers.append(cell.value)

    # Build the full column list (existing + any new contracts from today)
    new_cols = build_flat_columns(data)
    all_headers = list(existing_headers)  # preserve existing order
    for col in new_cols:
        if col not in all_headers:
            all_headers.append(col)

    # Write/update header row if this is a new sheet or new columns appeared
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    if not existing_headers:
        # Brand new sheet — write headers
        for col_idx, col_name in enumerate(all_headers, 1):
            cell = ws.cell(1, col_idx, col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.row_dimensions[1].height = 60
    elif len(all_headers) > len(existing_headers):
        # New contracts appeared — extend headers
        for col_idx in range(len(existing_headers) + 1, len(all_headers) + 1):
            cell = ws.cell(1, col_idx, all_headers[col_idx - 1])
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # ------------------------------------------------------------------
    # Append the data row
    # ------------------------------------------------------------------
    next_row = ws.max_row + 1
    for col_idx, col_name in enumerate(all_headers, 1):
        ws.cell(next_row, col_idx, flat_row.get(col_name, ""))

    # Freeze top row and set sensible column widths
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 26  # Scraped At
    ws.column_dimensions["B"].width = 34  # ASX Last Update
    for col_idx in range(3, len(all_headers) + 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 14

    wb.save(EXCEL_FILE)
    print(f"Saved to {EXCEL_FILE} — {scraped_at} | ASX Update: {last_update}")
    print(f"Total columns: {len(all_headers)} | Row written: {next_row}")


def main():
    print("Fetching ASX Energy NZ futures data (headless Chrome)...")
    html = fetch_html_with_selenium()
    last_update, data = parse_data(html)
    print(f"ASX Last Update: {last_update}")
    for loc, products in data.items():
        for prod, rows in products.items():
            print(f"  {loc} / {prod}: {len(rows)} contracts")
    write_to_excel(last_update, data)


if __name__ == "__main__":
    main()
