import asyncio
import gspread
import re
from datetime import datetime
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError
from config import SERVICE_ACCOUNT_PATH

SIRS_URL = "http://10.67.2.229/sirs/index.php?XP_xrptoolrun_xrptools=3&run=y&rp_id=17"
SHEET_NAME = "temp daftar obat"

# === Your predefined doctor list (Penulis Resep IDs) ===
DOCTOR_IDS = ["23","595","476","1914","1922","580","721","1986","142","277","3140","647","14","56","1987","1487","225","1577","1967","1488","13","3072","621","3364","2926","2466","140","3123","2060","1599","3449","1919","2436","1237","2662","2426","1920","1364","3365","2916","469","266","1943","553","1590","2259"]  # example subset


# === DATE RANGE HANDLER ==================================================
async def set_date_range(page, auto_submit: bool = True):
    """
    Set the start/end datetime range for the SIRS report.
    Handles hidden <input id='s_1'> and <input id='s_2'> fields,
    also updates visible spans (qdttm binding).
    """
    dt_format = "%Y-%m-%d %H:%M:%S"

    def prompt_dt(label):
        while True:
            s = input(f"Enter {label} datetime ({dt_format}): ").strip()
            try:
                return datetime.strptime(s, dt_format), s
            except ValueError:
                print("❌ Invalid format. Please use 'YYYY-MM-DD HH:MM:SS'.")

    dttm_from, dttm_from_str = prompt_dt("start")
    dttm_to, dttm_to_str = prompt_dt("end")

    if dttm_to <= dttm_from:
        raise ValueError("⚠️ End datetime must be greater than start datetime.")

    print(f"🕒 Using filter range: {dttm_from_str} → {dttm_to_str}")

    # Wait for required inputs
    try:
        await page.wait_for_selector("#s_1", timeout=15000)
        await page.wait_for_selector("#s_2", timeout=15000)
    except PWTimeoutError:
        print("❌ Could not find date input fields (#s_1 / #s_2). Check if you're on report page.")
        return False

    # Update both hidden inputs and visible spans
    await page.evaluate(
        """(from_, to_) => {
            const s1 = document.querySelector('#s_1');
            const s2 = document.querySelector('#s_2');
            const span1 = document.querySelector("span[onclick*='qdttm(\"s_1\"']");
            const span2 = document.querySelector("span[onclick*='qdttm(\"s_2\"']");
            if (!s1 || !s2) return false;

            s1.value = from_;
            s2.value = to_;

            const formatDisplay = (iso) => {
                const [date, time] = iso.split(' ');
                const [y, m, d] = date.split('-');
                const months = [
                    'Januari','Februari','Maret','April','Mei','Juni',
                    'Juli','Agustus','September','Oktober','November','Desember'
                ];
                const month = months[parseInt(m, 10) - 1];
                const [hh, mm] = time.split(':');
                return `${parseInt(d)} ${month} ${y} ${hh}:${mm}`;
            };

            if (span1) span1.textContent = formatDisplay(from_);
            if (span2) span2.textContent = formatDisplay(to_);

            // Trigger both JS and DOM events
            [s1, s2].forEach(el => {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            });

            if (typeof qdttm === 'function') {
                try { qdttm('s_1', span1, new Event('change')); } catch(e) {}
                try { qdttm('s_2', span2, new Event('change')); } catch(e) {}
            }

            return true;
        }""",
        dttm_from.strftime(dt_format),
        dttm_to.strftime(dt_format),
    )

    # Confirm
    vals = await page.evaluate("""() => ({
        from: document.querySelector('#s_1')?.value,
        to: document.querySelector('#s_2')?.value,
        span1: document.querySelector("span[onclick*='qdttm(\"s_1\"']")?.textContent,
        span2: document.querySelector("span[onclick*='qdttm(\"s_2\"']")?.textContent
    })""")

    print(f"✅ Date inputs updated: {vals['from']} → {vals['to']}")
    print(f"✅ Span display updated: {vals['span1']} → {vals['span2']}")

    if auto_submit:
        # Click “Kirim” or “Tampilkan”
        try:
            btn = page.locator("input[type='button'][value='Tampilkan'], input[type='button'][value='Kirim']")
            await btn.first.click(timeout=5000)
            print("▶️ Clicked 'Tampilkan' button.")
        except Exception:
            print("⚠️ Could not click 'Tampilkan' automatically, please verify manually.")
            return True

        try:
            await page.wait_for_selector("#dv_process_start", state="attached", timeout=15000)
            print("⏳ Process started...")
            await page.wait_for_selector("#dv_process_start", state="detached", timeout=120000)
            print("✅ Process finished, table visible.")
        except PWTimeoutError:
            print("⚠️ No process indicator detected — continuing anyway.")

    return True


# === PLAYWRIGHT SESSION ATTACH ===========================================
async def attach_browser(cdp_endpoint="http://127.0.0.1:9222"):
    """Attach to existing Chrome session."""
    p = await async_playwright().start()
    browser = await p.chromium.connect_over_cdp(cdp_endpoint)
    context = browser.contexts[0] if browser.contexts else await browser.new_context()
    page = context.pages[0] if context.pages else await context.new_page()
    print("✅ Attached to Chrome.")
    return p, browser, page


# === GOOGLE SHEETS =======================================================
def open_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=scopes)
    client = gspread.authorize(creds)
    ss = client.open_by_url(
        "https://docs.google.com/spreadsheets/d/1MdEQrxNS6kuHkwks8Fgg6q29HxJ3qx2br-DPBpGecn4"
    )
    return ss.worksheet(SHEET_NAME)


# === TABLE EXTRACTION ====================================================
async def extract_table_rows(page):
    """Extract rows from <table class='qresult'>."""
    try:
        await page.wait_for_selector("table.qresult tbody tr", timeout=15000)
    except PWTimeoutError:
        print("⚠️ No result table detected.")
        return []

    rows = page.locator("table.qresult tbody tr")
    count = await rows.count()
    data = []
    for i in range(count):
        tds = await rows.nth(i).locator("td").all_inner_texts()
        data.append(tds)
    print(f"📊 Extracted {len(data)} rows.")
    return data


# === MAIN RUNNER =========================================================
async def run_extraction():
    p, browser, page = await attach_browser()
    ws = open_sheet()

    await page.goto(SIRS_URL)
    await page.wait_for_selector("#rpf", timeout=15000)
    print("✅ Report filter form ready.")

    # await set_date_range(page)
    print("⏳ Waiting 5 seconds before prompting...")
    await asyncio.sleep(5)
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, lambda: input("Press Enter to continue when ready..."))

    for doc_id in DOCTOR_IDS:
        print(f"\n👩‍⚕️ Processing doctor ID: {doc_id}")

        # select doctor
        await page.select_option("#s_8_", doc_id)

        # click submit
        await page.locator("input[value='Kirim']").click()

        # wait for table refresh
        try:
            await page.wait_for_selector("div#loading", state="attached", timeout=2000)
            await page.wait_for_selector("div#loading", state="detached", timeout=15000)
        except:
            pass  # tolerate SIRS variants without #loading div

        # extract data
        rows = await extract_table_rows(page)
        if not rows:
            print(f"⚠️ No data for doctor {doc_id}")
            continue

        labeled_rows = [[doc_id] + r for r in rows]
        ws.append_rows(labeled_rows, value_input_option="USER_ENTERED")
        print(f"✅ Uploaded {len(labeled_rows)} rows for doctor {doc_id} to sheet.")

        await asyncio.sleep(2.0)  # pacing to prevent quota throttling

    print("\n🏁 Extraction completed.")
    await browser.close()
    await p.stop()


if __name__ == "__main__":
    asyncio.run(run_extraction())
