# test_playwright_sirs_cdp.py

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from config import SIRS_APP_URL

def main():
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx     = browser.contexts[0] if browser.contexts else browser.new_context()
        page    = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(SIRS_APP_URL, timeout=10000)

        # set filters
        page.select_option("select#urut",        label="Nama")
        page.select_option("select#jenis_rawat", label="Rawat Jalan")
        page.select_option("select#tanggal",     label="1")
        page.select_option("select#bulan",       label="Mei")
        page.select_option("select#tahun",       label="2025")
        page.click("input[value='Tampilkan']")

        locator = page.locator("div#dv_content table.tblcontrast tbody tr")
        try:
            locator.first.wait_for(timeout=5000)
        except PWTimeoutError:
            print("❌ No rows appeared.")
            browser.close()
            return

        count = locator.count()
        print(f"✅ Found {count} row{'s' if count != 1 else ''} in summary table.")
        browser.close()

if __name__ == "__main__":
    main()
