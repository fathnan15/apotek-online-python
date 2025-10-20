import asyncio
from sirs_runner import init_sirs_manual, get_claim_records, download_claims, set_playwright_context
from sheets_handler import get_worksheet, write_initial_sep_rows
from config import WORKSHEET_NAME
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        await set_playwright_context(p)  # sets _playwright, _browser, _page
        await init_sirs_manual()
        records = await get_claim_records()
        ws = get_worksheet(WORKSHEET_NAME)
        write = write_initial_sep_rows(ws, records)  # <-- remove 'await'
        download = await download_claims()
        if download and write:
            print(f"✅ Wrote {len(records)} records into your sheet.")
            print(f"✅ downloaded the claims.")
            print("----------------------------------")

if __name__ == "__main__":
    while True:
        asyncio.run(main())
        again = input("Run again? (y/n): ").strip().lower()
        if again != "y":
            break
