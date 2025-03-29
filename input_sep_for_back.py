from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from openpyxl import load_workbook
import time

sep_no = '0179R0270325V010902'
receipt_no = '58517'
date = '14/03/2025'

# Connect to existing Chrome session
# google-chrome --remote-debugging-port=9222 --user-data-dir="$HOME/.chrome_debug"
options = webdriver.ChromeOptions()
options.debugger_address = "localhost:9222"
driver = webdriver.Chrome(options=options)

print(driver.current_url)

wait = WebDriverWait(driver, 10)
failed_list = []


try:
        # ✅ Input SEP
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TxtREFASALSJP_I").send_keys(sep_no)

        # ✅ Click "Cari"
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnCariSEP_CD").click()

        # ✅ Wait until "Nomor Kartu" field is filled
        for _ in range(3):
            try:
                wait.until(lambda d: d.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNOKAPST_I").get_attribute("value").strip() != "")
                break
            except StaleElementReferenceException:
                time.sleep(1)
                continue

        # ✅ Input Jenis Resep
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_cboJnsObat_I").send_keys("Obat Kronis Blm Stabil")

        # ✅ Input No Resep
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_txtNoResep_I").send_keys(receipt_no)

                # ✅ Input No Resep
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_dtpTGLRSP_I").send_keys(date)

        # ✅ Click "Simpan"
        driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnSimpan_CD").click()

        # ✅ Wait for Alert
        try:
            alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert_text = alert.text.strip()
            alert.accept()  # Click "OK"

            if "Simpan Berhasil" in alert_text:
                print(f" {sep_no} - Success!")
            else:
                # ✅ Write error message to column E
                print(f" {sep_no} - Error: {alert_text}")

                # ✅ Click Reset before continuing
                driver.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnReset_CD").click()
                # ✅ Wait until "No SEP" field is empty
                for _ in range(3):
                    try:
                        wait.until(lambda d: d.find_element(By.ID, "ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TxtREFASALSJP_I").get_attribute("value").strip() == "")
                        break
                    except StaleElementReferenceException:
                        time.sleep(1)
                        continue
                print(f"{sep_no}  - Reset Done.")

        except TimeoutException:
            print(f"{sep_no}  - No alert found")
    
            print(f"{sep_no}  - Process completed.")
    
except TimeoutException:
    failed_list.append(sep_no)


print("Process Completed")
# print("Failed Entries:", failed_list)
