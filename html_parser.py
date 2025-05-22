# html_parser.py

import re
import requests
from bs4 import BeautifulSoup

def fetch_html(source: str) -> str:
    """Fetch HTML from URL or local file path."""
    if source.startswith(("http://", "https://")):
        r = requests.get(source, timeout=30)
        r.raise_for_status()
        return r.text
    else:
        with open(source, "r", encoding="utf-8") as f:
            return f.read()

def parse_claim_html(source: str) -> list[dict]:
    """
    Parses the SIRS page and returns a list of dicts:
      { "sep": "<SEP_NUMBER>", "receipt": "<PRESCRIPTION_ID>" }
    """
    html = fetch_html(source)
    soup = BeautifulSoup(html, "html.parser")

    # 1) Find the container and table
    container = soup.find("div", id="dv_content")
    if not container:
        raise ValueError("Could not find <div id='dv_content'> in HTML.")  
    table = container.find("table", class_="tblcontrast")
    if not table:
        raise ValueError("Could not find <table class='tblcontrast'> in HTML.")

    records = []
    # 2) Walk each data row
    for row in table.find("tbody").find_all("tr"):
        cols = row.find_all("td")
        if len(cols) < 4:
            continue

        # 3) SEP is in the 4th column (zero-based index 3)
        sep_val = cols[3].get_text(strip=True)

        # 4) Find the Print Resep button and grab its onclick argument
        btn = row.find("input", {"value": "Print Resep"})
        if not btn or "onclick" not in btn.attrs:
            continue
        onclick = btn["onclick"]

        m = re.search(r'print_prescription\("([^"]+)"', onclick)
        if not m:
            continue
        receipt_val = m.group(1)

        records.append({"sep": sep_val, "receipt": receipt_val})

    if not records:
        raise ValueError("No SEP/receipt pairs found. Parser may need adjustment.")

    return records
