import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://my.greenprofi.de/")
    page.goto("https://my.greenprofi.de/cockpit")
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    
    page.get_by_role("textbox", name="E-Mail").click()
    page.get_by_role("textbox", name="E-Mail").fill("")
    page.get_by_role("textbox", name="Passwort").click()
    page.get_by_role("textbox", name="Passwort").fill("")
    page.locator("[id=\"__BVID__136___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()

   
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.get_by_role("button", name="Ergebnisse pro Seite:").click()
    page.get_by_role("menuitem", name="50").click()
    page.locator("span:nth-child(5) > a").first.click()
    page.get_by_text("Leistungsumfang", exact=True).click()
    page.get_by_text("Baubeginn/Ende").click()
    page.locator("label").filter(has_text=re.compile(r"^50$")).click()
    with page.expect_download() as download_info:
        page.get_by_role("link", name="Exportieren").click()
    download = download_info.value

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
