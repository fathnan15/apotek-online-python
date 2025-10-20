# utils.py

from playwright.sync_api import Page

def reset_form(page: Page):
    """
    Clicks the Reset button on SIRS and waits
    until the SEP input is empty again.
    """
    page.click("#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_BtnReset_CD")
    page.wait_for_function(
        "() => document.querySelector('#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TxtREFASALSJP_I')"
        ".value.trim() === ''",
        timeout=3000
    )
