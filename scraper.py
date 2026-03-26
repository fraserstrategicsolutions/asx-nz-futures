import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import pytz

DATASET_URL = "https://www.asxenergy.com.au/futures_nz/dataset"
EXCEL_FILE = "asx_nz_futures.xlsx"
LOCATIONS = ["Otahuhu", "Benmore"]
COLUMNS = ["Contract", "Bid Size", "Bid", "Ask", "Ask Size", "High", "Low", "Last", "+/-", "Vol"]

def fetch_data():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    resp = requests.get(DATASET_URL, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.text

def parse_data(html):
    soup = BeautifulSoup(html, "html.parser")

    # Extract last update timestamp
    last_div = soup.find("div", id="last")
    last_update = ""
    if last_div:
        date_str = last_div.get("data-date", "")
        time_str = last_div.get("data-time", "")
        last_update = f"{date_str} {time_str}".strip()

    results = {}

    # Find all market-dataset sections
    for section in soup.find_all("div", class_="market-dataset"):
        h2 = section.find("h2")
        if not h2:
            continue
        location_name = h2.get_text(strip=True)
        if location_name not in LOCATIONS:
            continue

        location_data = {}

        # Each product type (Base Month, Base Quarter, etc.)
        for product_section in section.find_all("div", class_="dataset"):
            product_table = product_section.find("table")
            if not product_table:
                continue

            # Get the product name from the h3 above this dataset div
            parent = product_section.parent
            h3 = parent.find("h3") if parent else None
            product_name = h3.get_text(strip=True).replace("ED", "").strip() if h3 else "Unknown"
            # Clean up button text that gets included
            product_name = product_name.split("\n")[0].strip()

            rows = []
            tbody = product_table.find("tbody")
            if tbody:
                for tr in tbody.find_all("tr"):
                    cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                    if cells and any(c for c in cells):
                        rows.append(cells)

            if rows:
                location_data[product_name] = rows

        results[location_name] = location_data

    return last_update, results

def get_nzt_timestamp():
    nzt = pytz.timezone("Pacific/Auckland")
    return datetime.now(nzt).strftime("%Y-%m-%d %H:%M:%S NZT")

def write_to_excel(last_update, data):
    scraped_at = get_nzt_timestamp()

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill_blue = PatternFill("solid", fgColor="1F4E79")   # Otahuhu - dark blue
    header_fill_green = PatternFill("solid", fgColor="375623")  # Benmore - dark green
    meta_fill = PatternFill("solid", fgColor="D9E1F2")

    for location in LOCATIONS:
        if location not in data:
            continue

        sheet_name = location
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
            # Set column widths
            ws.column_dimensions["A"].width = 20
            for col in "BCDEFGHIJ":
                ws.column_dimensions[col].width = 12

        fill = header_fill_blue if location == "Otahuhu" else header_fill_green

        # Find next empty row (leave a gap between daily snapshots)
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            start_row = 1
        else:
            start_row = ws.max_row + 2  # blank row separator between days

        row = start_row

        # Meta row: scraped timestamp + last update
        ws.cell(row, 1, f"Scraped: {scraped_at}").font = Font(italic=True)
        ws.cell(row, 2, f"ASX Last Update: {last_update}").font = Font(italic=True)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=len(COLUMNS))
        for col in range(1, len(COLUMNS) + 1):
            ws.cell(row, col).fill = meta_fill
        row += 1

        for product_name, rows in data[location].items():
            # Product header
            ws.cell(row, 1, product_name)
            ws.cell(row, 1).font = Font(bold=True, color="FFFFFF")
            ws.cell(row, 1).fill = fill
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(COLUMNS))
            row += 1

            # Column headers
            for col_idx, col_name in enumerate(COLUMNS, 1):
                cell = ws.cell(row, col_idx, col_name)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center")
            row += 1

            # Data rows
            for data_row in rows:
                for col_idx, value in enumerate(data_row[:len(COLUMNS)], 1):
                    ws.cell(row, col_idx, value)
                row += 1

            row += 1  # gap between product types

    wb.save(EXCEL_FILE)
    print(f"Saved to {EXCEL_FILE} — {scraped_at} | ASX Update: {last_update}")

def main():
    print("Fetching ASX Energy NZ futures data...")
    html = fetch_data()
    last_update, data = parse_data(html)
    print(f"ASX Last Update: {last_update}")
    for loc, products in data.items():
        print(f"  {loc}: {list(products.keys())}")
    write_to_excel(last_update, data)

if __name__ == "__main__":
    main()
