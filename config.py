# config.py

# — SIRS extraction settings —  
# — SIRS extraction settings —  
SIRS_APP_URLS = {
    "223": "http://10.67.2.223/sirs/index.php?XP_ehrdocumentfarmasi_menu=0",
    "229": "http://10.67.2.229/sirs/index.php?XP_ehrdocumentfarmasi_menu=0"
} 
SIRS_SELECTORS = {
    "row":                 "div#dv_content table.tblcontrast tbody tr",
    "sep_cell_index":      3,
    "print_button":        "input[value='Print Resep']",
}

# — Apotek submission settings —  
APOTEK_URL = "https://apotek.bpjs-kesehatan.go.id/apotek/RspMsk1.aspx"  
APOTEK_SELECTORS = {
    "sep_input":           "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TxtREFASALSJP_I",
    "cari_button":         "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnCariSEP_CD",
    "no_kartu_input":      "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNOKAPST_I",
    "receipt_type_input":  "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_cboJnsObat_I",
    "receipt_input":       "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNoResep_I",
    "simpan_button":       "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnSimpan_CD",
    "reset_button":        "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnReset_CD",
    
    # NEW SELECTORS FOR REVISION
    "date_receipt_input":  "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_dtpTGLRSP_I",
    "date_service_input":  "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_dtpTGLPELRSP_I",
    "iteration_status":     "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_cboIterasi_I"
}

# — Google Sheets settings —  
SHEET_URL            = "https://docs.google.com/spreadsheets/d/1f-quvC9jSRnTvjMKFUbsge0XES4gQ3vSxbE44TUGG4o"  
WORKSHEET_NAME       = "sep_web_driver"
REVISE_WORKSHEET_NAME = "revise_input" # New Sheet Name
SERVICE_ACCOUNT_PATH = "./keys/sep-sync-bot.json"  

# — Column headers for sep_web_driver (A→G) —  
SEP_SHEET_HEADERS = [  
    "sep_dttm",      # A: timestamp
    "mrn",           # B: medical record number
    "sep_num",       # C: SEP number
    "presc_id",      # D: prescription id (for internal use, not in sheet)
    "receipt_num",   # E: prescription number
    "receipt_type",  # F: “Obat Kronis Blm Stabil” or “Obat Kemoterapi”
    "updated_dttm",  # G: timestamp of last update
    "status",        # H: normal / error
    "note"           # I: alert message or “-”
]