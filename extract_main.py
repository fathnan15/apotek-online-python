import asyncio
import argparse
from sirs_runner import init_sirs_manual, get_claim_records, download_claims, set_playwright_context
from sheets_handler import get_worksheet, write_initial_sep_rows
from config import WORKSHEET_NAME
from playwright.async_api import async_playwright

async def main(start_day: int, end_day: int, bulan: str, port: str, server: str):
    async with async_playwright() as p:
        cdp_url = f"http://127.0.0.1:{port}"
        await set_playwright_context(p, cdp_endpoint=cdp_url)  

        for day in range(start_day, end_day + 1):
            date_str = str(day)
            print(f"\n🔁 Processing date {date_str} {bulan} on Server .{server} (Port {port})...", flush=True)

            await init_sirs_manual(date=date_str, bulan=bulan, server=server)
            records = await get_claim_records()
            ws = get_worksheet(WORKSHEET_NAME)
            write_ok = write_initial_sep_rows(ws, records) 
            # download_ok = await download_claims()
            
            # if download_ok and write_ok:
            if write_ok:
                print(f"✅ {date_str}: Wrote {len(records)} records into your sheet.", flush=True)
                print(f"✅ {date_str}: Downloaded the claims.", flush=True)
                print("----------------------------------", flush=True)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SIRS Data Extractor Worker")
    parser.add_argument("--port", type=str, required=True, help="Chrome CDP Port (e.g., 9222)")
    parser.add_argument("--start", type=int, required=True, help="Start date (DD)")
    parser.add_argument("--end", type=int, required=True, help="End date (DD)")
    parser.add_argument("--bulan", type=str, required=True, help="Bulan string (e.g., September)")
    parser.add_argument("--server", type=str, choices=["223", "229"], required=True, help="Target Server IP ending (223 or 229)")
    
    args = parser.parse_args()

    asyncio.run(main(args.start, args.end, args.bulan, args.port, args.server))