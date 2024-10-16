import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright, expect
import json
import logging


# Set up logging
logging.basicConfig(filename='greenprofi_extraction.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# SharePoint configuration
SHAREPOINT_SITE_URL = "https://ingenschmidt.sharepoint.com"
SHAREPOINT_FOLDER_URL = "https://ingenschmidt.sharepoint.com/sites/Dateiserver/Freigegebene%20Dokumente/Forms/AllItems.aspx?e=5%3Ad39c7e60281640508fb3754902a91cb1&sharingv2=true&fromShare=true&at=9&CID=bd5051bf%2D1d8b%2D4b3f%2Dae2a%2D6508b3b3e715&FolderCTID=0x0120001744D1600461AA45985A863F14BB27A6&id=%2Fsites%2FDateiserver%2FFreigegebene%20Dokumente%2F04%5FMarketing%2FAkquise%2FGreenProfi%20Extracted%20Data%20Test&viewid=07a9ee37%2Dd793%2D45c0%2Da6b9%2D2a98315dbf3d"
USERNAME = "r.pancho@ib-bauabrechnung.de"
PASSWORD = "xvGR%XLh%#Eg"

def get_last_extraction_date():
    # Get the last extraction date from the JSON file
    print("Getting the last extraction date from the JSON file...")
    try:
        with open('last_extraction.json', 'r') as f:
            data = json.load(f)
            return datetime.datetime.strptime(data['last_extraction'], '%Y-%m-%d').date()
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        logging.info("No previous extraction date found. This appears to be the first run.")
        print("No previous extraction date found. This appears to be the first run.")
        return None

def save_last_extraction_date(date):
    # Save the current extraction date to the JSON file
    print(f"Saving the last extraction date: {date}...")
    with open('last_extraction.json', 'w') as f:
        json.dump({'last_extraction': date.strftime('%Y-%m-%d')}, f)
    logging.info(f"Saved last extraction date: {date}")
    print(f"Saved last extraction date: {date}")

def get_last_downloaded_project():
    # Get the last downloaded project from the JSON file
    print("Getting the last downloaded project from the JSON file...")
    if not os.path.exists('last_downloaded_project.json'):
        print("No previous project found. Creating new file 'last_downloaded_project.json'...")
        with open('last_downloaded_project.json', 'w') as f:
            json.dump({'last_project': None}, f)
        return None
    try:
        with open('last_downloaded_project.json', 'r') as f:
            data = json.load(f)
            return data.get('last_project', None)
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        logging.info("No previous project found. This appears to be the first run.")
        print("No previous project found. This appears to be the first run.")
        return None

def save_last_downloaded_project(project_id):
    # Save the last downloaded project to the JSON file
    print(f"Saving the last downloaded project: {project_id}...")
    with open('last_downloaded_project.json', 'w') as f:
        json.dump({'last_project': project_id}, f)
    logging.info(f"Saved last downloaded project: {project_id}")
    print(f"Saved last downloaded project: {project_id}")
          
def upload_to_sharepoint(page, local_file_path):
    try:
        # Navigate to the SharePoint document library
        print("Navigating to SharePoint document library...")
        page.goto(SHAREPOINT_FOLDER_URL)
        time.sleep(30)

        # Click on the upload button and directly set the file path for upload
        print("Clicking on upload button...")
        upload_button = page.locator('button[data-automationid="uploadCommand"][aria-label="Upload"]')
        upload_button.click()
        page.click('text="Files"')

        print(f"Uploading file from: {local_file_path}")
        page.set_input_files('input[type="file"]', local_file_path)

        # Wait for the file to finish uploading
        print("Waiting for file to finish uploading...")
        page.wait_for_timeout(15000)  # Adjust this as needed based on network speed and file size

        logging.info(f"File uploaded successfully to SharePoint: {local_file_path}")
        print(f"File uploaded successfully to SharePoint: {local_file_path}")
        return True
    except Exception as e:
        logging.error(f"Error uploading file to SharePoint: {str(e)}")
        print(f"Error uploading file to SharePoint: {str(e)}")
        return False

def run(playwright: Playwright, browser_type: str) -> None:
    logging.info("Starting the script...")
    print("Starting the script...")

    # Set up download path
    print("Setting up download path...")
    if os.name == 'nt':
        default_download_path = os.path.expanduser("~\\Downloads")
    else:
        default_download_path = os.path.expanduser("~//Downloads")
    
    custom_download_folder = os.path.join(default_download_path, "greenprofi_downloads")
    os.makedirs(custom_download_folder, exist_ok=True)
    print(f"Custom download folder created: {custom_download_folder}")

    # Launch the browser
    print(f"Launching {browser_type} browser...")
    if browser_type == "chrome":
        browser = playwright.chromium.launch(headless=False) 
    elif browser_type == "edge":
        browser = playwright.chromium.launch(headless=False, channel="msedge")  
    else:
        raise ValueError("Unsupported browser type. Use 'chrome' or 'edge'.")
    
    # Create a new browser context
    print("Creating new context...")
    context = browser.new_context(
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )                              
    
    # Create a new page
    print("Creating new page...")
    page = context.new_page()
    
    # Navigate to GreenProfi website and download the file
    print("Navigating to greenprofi.de...")
    page.goto("https://my.greenprofi.de/cockpit")
    page.wait_for_load_state("networkidle")
    
    # Accept cookies
    print("Accepting cookies...")
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    page.wait_for_load_state("networkidle")

    # Log in to GreenProfi
    print("Logging in to GreenProfi...")
    page.get_by_role("textbox", name="E-Mail").fill("b.schmidt@ib-bauabrechnung.de")
    page.get_by_role("textbox", name="Passwort").fill("bydfuc-gowmyW-vexbo9")
    page.locator("[id=\"__BVID__81___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()
    page.wait_for_load_state("load")
    
    last_project_id = get_last_downloaded_project()
    print(f"Last downloaded project: {last_project_id}")

    # Navigate to Submissionsergebnisse and download the Excel file
    print("Navigating to Submissionsergebnisse...")
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.wait_for_load_state("load")
    
    print("Setting results per page to 100...")
    page.get_by_role("button", name="Ergebnisse pro Seite:").click()
    page.get_by_role("menuitem", name="100").click()
    page.wait_for_load_state("load")
    
    page.locator("span:nth-child(5) > a").first.click()
    
    print("Adjusting filter settings...")
    page.get_by_text("Leistungsumfang", exact=True).click()
    page.get_by_text("Baubeginn/Ende").click()
    page.locator("label").filter(has_text="100").click()
    
    print("Initiating download...")
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

    final_path = os.path.join(custom_download_folder, download.suggested_filename)
    print(f"File path to be uploaded: {final_path}")

    # Check if file exists before attempting upload
    if not os.path.exists(final_path):
        print(f"Error: File not found at {final_path}")
    else:
        download.save_as(final_path)
        print(f"Download completed: {final_path}")

         # Save last downloaded project
        save_last_downloaded_project(download.suggested_filename)

        # Upload the file to SharePoint
        print("Navigating to SharePoint site...")
        page.goto(SHAREPOINT_SITE_URL)
        page.wait_for_load_state("networkidle")

        # Log in to SharePoint
        print("Logging in to SharePoint...")
        page.fill('input[type="email"]', USERNAME)
        page.click('input[type="submit"]')
        page.fill('input[type="password"]', PASSWORD)
        page.click('input[type="submit"]')
        time.sleep(5)  # Adjust as needed for login process
        try:
            page.click('input[value="Yes"]')  # Stay signed in prompt
        except:
            pass

        # Uploading file to SharePoint
        print(f"Uploading file to SharePoint: {final_path}...")
        if upload_to_sharepoint(page, final_path):
            save_last_extraction_date(datetime.datetime.now().date())
        else:
            print("Failed to upload to SharePoint. Last extraction date not updated.")
            logging.warning("Failed to upload to SharePoint. Last extraction date not updated.")

    # Close the browser
    print("Closing browser...")
    context.close()
    browser.close()

    logging.info("Script execution finished.")
    print("Script execution finished.")

print("Initializing Playwright...")
with sync_playwright() as playwright:
    run(playwright, browser_type="edge")
print("Script execution finished. Check the log file for detailed information.")