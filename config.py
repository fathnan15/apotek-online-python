# config.py

# — SIRS extraction settings —  
SIRS_APP_URL = "http://10.67.2.223/sirs/index.php?XP_ehrdocumentfarmasi_menu=0"  
SIRS_SELECTORS = {
    "row":                 "div#dv_content table.tblcontrast tbody tr",
    "sep_cell_index":      3,
    "print_button":        "input[value='Print Resep']",
}

# — Apotek submission settings —  
APOTEK_URL = "https://apotek.bpjs-kesehatan.go.id/apotek/RspMsk1.aspx"  
APOTEK_SELECTORS = {
    "sep_input":"#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TxtREFASALSJP_I",               # SEP number field  
    "cari_button":"#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnCariSEP_CD",    # Cari button  
    "no_kartu_input": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNOKAPST_I",           # No Kartu field  
    "receipt_type_input":  "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_cboJnsObat_I",          # Jenis Resep dropdown  
    "receipt_input":       "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNoResep_I",           # No Resep field  
    "simpan_button":       "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnSimpan_CD", # Simpan button  
    "reset_button":        "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnReset_CD",  # Reset button  
}

# — Google Sheets settings —  
SHEET_URL            = "https://docs.google.com/spreadsheets/d/1f-quvC9jSRnTvjMKFUbsge0XES4gQ3vSxbE44TUGG4o"  
WORKSHEET_NAME       = "sep_web_driver"  
SERVICE_ACCOUNT_PATH = "./keys/sep-sync-bot.json"  

# — Column headers for sep_web_driver (A→G) —  
SEP_SHEET_HEADERS = [  
    "sep_dttm",      # A: timestamp
    "mrn",           # B: medical record number
    "sep_num",       # B: SEP number
    "receipt_num",   # C: prescription number
    "receipt_type",  # D: “Obat Kronis Blm Stabil” or “Obat Kemoterapi”
    "updated_dttm",  # E: timestamp of last update
    "status",        # F: normal / error
    "note"           # G: alert message or “-”
]
