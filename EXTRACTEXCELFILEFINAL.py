import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright, expect
import json
import logging
import streamlit as st

# Set up logging
logging.basicConfig(filename='greenprofi_extraction.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Load secrets from Streamlit
SHAREPOINT_SITE_URL = st.secrets["SHAREPOINT_SITE_URL"]
SHAREPOINT_FOLDER_URL = st.secrets["SHAREPOINT_FOLDER_URL"]
USERNAME = st.secrets["USERNAME"]
PASSWORD = st.secrets["PASSWORD"]
GREENPROFI_EMAIL = st.secrets["GREENPROFI_EMAIL"]
GREENPROFI_PASSWORD = st.secrets["GREENPROFI_PASSWORD"]

def get_last_extraction_date():
    # Get the last extraction date from the JSON file
    try:
        with open('last_extraction.json', 'r') as f:
            data = json.load(f)
            return datetime.datetime.strptime(data['last_extraction'], '%Y-%m-%d').date()
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        logging.info("No previous extraction date found. This appears to be the first run.")
        return None

def save_last_extraction_date(date):
    # Save the current extraction date to the JSON file
    with open('last_extraction.json', 'w') as f:
        json.dump({'last_extraction': date.strftime('%Y-%m-%d')}, f)
    logging.info(f"Saved last extraction date: {date}")

def get_last_downloaded_project():
    # Get the last downloaded project from the JSON file
    if not os.path.exists('last_downloaded_project.json'):
        with open('last_downloaded_project.json', 'w') as f:
            json.dump({'last_project': None}, f)
        return None
    try:
        with open('last_downloaded_project.json', 'r') as f:
            data = json.load(f)
            return data.get('last_project', None)
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        logging.info("No previous project found. This appears to be the first run.")
        return None

def save_last_downloaded_project(project_id):
    # Save the last downloaded project to the JSON file
    with open('last_downloaded_project.json', 'w') as f:
        json.dump({'last_project': project_id}, f)
    logging.info(f"Saved last downloaded project: {project_id}")

def upload_to_sharepoint(page, local_file_path):
    try:
        # Navigate to the SharePoint document library
        page.goto(SHAREPOINT_FOLDER_URL)
        time.sleep(30)

        # Click on the upload button and directly set the file path for upload
        upload_button = page.locator('button[data-automationid="uploadCommand"][aria-label="Upload"]')
        upload_button.click()
        page.click('text="Files"')

        page.set_input_files('input[type="file"]', local_file_path)

        # Wait for the file to finish uploading
        page.wait_for_timeout(15000)  # Adjust this as needed based on network speed and file size

        logging.info(f"File uploaded successfully to SharePoint: {local_file_path}")
        return True
    except Exception as e:
        logging.error(f"Error uploading file to SharePoint: {str(e)}")
        return False

def run(playwright: Playwright, browser_type: str) -> None:
    logging.info("Starting the script...")

    # Set up download path
    default_download_path = os.path.expanduser("~//Downloads")
    custom_download_folder = os.path.join(default_download_path, "greenprofi_downloads")
    os.makedirs(custom_download_folder, exist_ok=True)

    # Launch the browser
    if browser_type == "chrome":
        browser = playwright.chromium.launch(headless=True) 
    elif browser_type == "edge":
        browser = playwright.chromium.launch(headless=False, channel="msedge")  
    else:
        raise ValueError("Unsupported browser type. Use 'chrome' or 'edge'.")
    
    # Create a new browser context
    context = browser.new_context(
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )                              
    
    # Create a new page
    page = context.new_page()
    
    # Navigate to GreenProfi website and download the file
    page.goto("https://my.greenprofi.de/cockpit")
    page.wait_for_load_state("networkidle")
    
    # Accept cookies
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    page.wait_for_load_state("networkidle")

    # Log in to GreenProfi
    page.get_by_role("textbox", name="E-Mail").fill(GREENPROFI_EMAIL)
    page.get_by_role("textbox", name="Passwort").fill(GREENPROFI_PASSWORD)
    page.locator("[id=\"__BVID__81___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()
    page.wait_for_load_state("load")
    
    last_project_id = get_last_downloaded_project()

    # Navigate to Submissionsergebnisse and download the Excel file
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.wait_for_load_state("load")
    
    page.get_by_role("button", name="Ergebnisse pro Seite:").click()
    page.get_by_role("menuitem", name="100").click()
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
    final_path = os.path.join(custom_download_folder, download.suggested_filename)

    # Check if file exists before attempting upload
    if not os.path.exists(final_path):
        logging.error(f"Error: File not found at {final_path}")
    else:
        download.save_as(final_path)

        # Save last downloaded project
        save_last_downloaded_project(download.suggested_filename)

        # Upload the file to SharePoint
        page.goto(SHAREPOINT_SITE_URL)
        page.wait_for_load_state("networkidle")

        # Log in to SharePoint
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
        if upload_to_sharepoint(page, final_path):
            save_last_extraction_date(datetime.datetime.now().date())
        else:
            logging.warning("Failed to upload to SharePoint. Last extraction date not updated.")

    # Close the browser
    context.close()
    browser.close()
    logging.info("Script execution finished.")

with sync_playwright() as playwright:
    run(playwright, browser_type="chrome")
