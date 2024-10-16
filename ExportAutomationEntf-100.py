import os
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright, browser_type: str) -> None:
   
    if os.name == 'nt':
        default_download_path = os.path.expanduser("~\\Downloads")
    else:
        default_download_path = os.path.expanduser("~/Downloads")

    
    if browser_type == "chrome":
        browser = playwright.chromium.launch(headless=False) 
    elif browser_type == "edge":
        browser = playwright.chromium.launch(headless=False, channel="msedge")  
    else:
        raise ValueError("Unsupported browser type. Use 'chrome' or 'edge'.")
    
    
    context = browser.new_context(
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )                              
    
    
    page = context.new_page()
    
    
    page.goto("https://my.greenprofi.de/cockpit")
    page.wait_for_load_state("networkidle")
    
    
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    page.wait_for_load_state("networkidle")

    
    
    page.get_by_role("textbox", name="E-Mail").fill("b.schmidt@ib-bauabrechnung.de")
    page.get_by_role("textbox", name="Passwort").fill("bydfuc-gowmyW-vexbo9")
    page.locator("[id=\"__BVID__81___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()
    page.wait_for_load_state("load")
    
    
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.wait_for_load_state("load")
    
    
    page.get_by_role("button", name="Ergebnisse pro Seite:").click()
    page.get_by_role("menuitem", name="100").click()
    page.wait_for_load_state("load")
    
    
    page.get_by_role("cell", name="Entf.").click()
    page.wait_for_load_state("load")
    page.locator("span:nth-child(5) > a").first.click()
    
    
    page.get_by_text("Leistungsumfang", exact=True).click()
    page.get_by_text("Baubeginn/Ende").click()
    page.locator("label").filter(has_text="100").click()
    
   
    with page.expect_download() as download_info:

        page.evaluate("""
            () => {
                const exportLink = Array.from(document.querySelectorAll('a')).find(el => el.textContent.trim() === 'Exportieren');
                if (exportLink) {
                    exportLink.click();
                } else {
                    throw new Error('Export link not found');
                }
            }
        """)
   
    
    
    download = download_info.value
    print(f"Download started: {download.suggested_filename}")

    final_path = os.path.join(default_download_path,download.suggested_filename)
    download.save_as(final_path)
    print(f"Download completed:{final_path}")
  
    
    
    if download.suggested_filename.endswith('.xls') or download.suggested_filename.endswith('.xlsx'):
        print(f"Successfully downloaded Excel file at: {final_path}")
    else:
        print(f"Downloaded file is not an Excel file: {final_path}")  


    
    
    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright, browser_type="edge")  
