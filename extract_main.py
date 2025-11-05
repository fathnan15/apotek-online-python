import asyncio
from sirs_runner import init_sirs_manual, get_claim_records, download_claims, set_playwright_context
from sheets_handler import get_worksheet, write_initial_sep_rows
from config import WORKSHEET_NAME
from playwright.async_api import async_playwright

async def main(start_day: int, end_day: int, bulan: str):
    async with async_playwright() as p:
        await set_playwright_context(p)  # sets _playwright, _browser, _page

        for day in range(start_day, end_day + 1):
            date_str = str(day)
            print(f"\n🔁 Processing date {date_str} {bulan} ...", flush=True)

            await init_sirs_manual(date=date_str, bulan=bulan)
            records = await get_claim_records()
            ws = get_worksheet(WORKSHEET_NAME)
            write_ok = write_initial_sep_rows(ws, records)  # sync
            download_ok = await download_claims()
            if download_ok and write_ok:
                print(f"✅ {date_str}: Wrote {len(records)} records into your sheet.", flush=True)
                print(f"✅ {date_str}: downloaded the claims.", flush=True)
                print("----------------------------------", flush=True)

if __name__ == "__main__":
    while True:
        start = int(input("Start date (DD): ").strip())
        end = int(input("End date (DD): ").strip())
        bulan = input("Bulan (e.g. September): ").strip()
        asyncio.run(main(start, end, bulan))
        again = input("Run again? (y/n): ").strip().lower()
        if again != "y":
            break
