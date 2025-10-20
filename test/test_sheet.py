# test_sheets.py

from sheets_handler import get_worksheet

def main():
    try:
        ws = get_worksheet("sep_web_driver")
        print("✅ Connected to worksheet:", ws.title)
        sample = ws.get_all_values()[:5]
        print("Sample data:")
        for row in sample:
            print(" ", row)
    except Exception as e:
        print("❌ Sheets API test failed:", e)

if __name__ == "__main__":
    main()
