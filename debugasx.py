"""
Temporary diagnostic script - dumps the rendered page structure so we can
see exactly what HTML ASX serves and fix the parser accordingly.
Run once, check the Actions log, then remove this file.
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

DATASET_URL = "https://www.asxenergy.com.au/futures_nz/dataset"

def fetch_html():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.get(DATASET_URL)
        # Wait for ANY table row, or fall back after 30s
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
            )
            print("TABLE ROWS FOUND - JS rendered successfully")
        except Exception:
            print("WARNING: No table rows found after 30s - page may not be rendering data")
        time.sleep(5)
        html = driver.page_source
    finally:
        driver.quit()
    return html

html = fetch_html()
soup = BeautifulSoup(html, "html.parser")

print("\n========== ALL CLASSES ON DIVs (unique) ==========")
classes = set()
for tag in soup.find_all(True):
    for c in tag.get("class", []):
        classes.add(f"{tag.name}.{c}")
for c in sorted(classes):
    print(c)

print("\n========== ALL H2 / H3 TEXT ==========")
for tag in soup.find_all(["h2", "h3"]):
    print(f"  <{tag.name}>: {repr(tag.get_text(strip=True)[:100])}")

print("\n========== #last div ==========")
last = soup.find(id="last")
print(last)

print("\n========== TABLES FOUND ==========")
tables = soup.find_all("table")
print(f"Total tables: {len(tables)}")
for i, t in enumerate(tables[:6]):
    print(f"\n--- Table {i} ---")
    # Print classes on this table and its parents
    print(f"  table classes: {t.get('class')}")
    parent = t.parent
    while parent and parent.name != "body":
        print(f"  parent <{parent.name}> classes={parent.get('class')} id={parent.get('id')}")
        parent = parent.parent
    # Print first 4 rows
    for row in t.find_all("tr")[:4]:
        print("  ROW:", [c.get_text(strip=True) for c in row.find_all(["th", "td"])])

print("\n========== FIRST 3000 CHARS OF BODY ==========")
body = soup.find("body")
if body:
    print(body.get_text(separator="\n", strip=True)[:3000])
