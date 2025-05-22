# config.py

# 1) SIRS page (for extraction)
SIRS_APP_URL         = "http://10.67.2.223/sirs/index.php?XP_ehrdocumentfarmasi_menu=0"

# 2) Apotek BPJS entry form
APOTEK_URL           = "https://apotek.bpjs-kesehatan.go.id/apotek/RspMsk1.aspx"

# 3) Google Sheets settings
SHEET_URL            = "https://docs.google.com/spreadsheets/d/1f-quvC9jSRnTvjMKFUbsge0XES4gQ3vSxbE44TUGG4o"
WORKSHEET_NAME       = "sep_web_driver"
KEMO_SHEET_NAME      = "kemo_recp_num"
SERVICE_ACCOUNT_PATH = "./keys/sep-sync-bot-abcdef123456.json"

# 4) sep_web_driver headers (columns A→F)
SEP_SHEET_HEADERS    = [
    "dttm",          # A: timestamp
    "sep_num",       # B: SEP number
    "receipt_num",   # C: prescription number
    "receipt_type",  # D: “Obat Kronis Blm Stabil” or “Obat Kemoterapi”
    "status",        # E: normal / error
    "note"           # F: alert message or “-”
]
