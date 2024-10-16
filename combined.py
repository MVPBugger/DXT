import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright
import json
import logging
import streamlit as st
import threading

# Set up logging
logging.basicConfig(filename='greenprofi_extraction.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# SharePoint and GreenProfi configuration
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL", "https://ingenschmidt.sharepoint.com")
SHAREPOINT_FOLDER_URL = os.getenv("SHAREPOINT_FOLDER_URL", "https://ingenschmidt.sharepoint.com/sites/Dateiserver/Freigegebene%20Dokumente/Forms/AllItems.aspx?e=5%3Ad39c7e60281640508fb3754902a91cb1&sharingv2=true&fromShare=true&at=9&CID=bd5051bf%2D1d8b%2D4b3f%2Dae2a%2D6508b3b3e715&FolderCTID=0x0120001744D1600461AA45985A863F14BB27A6&id=%2Fsites%2FDateiserver%2FFreigegebene%20Dokumente%2F04%5FMarketing%2FAkquise%2FGreenProfi%20Extracted%20Data%20Test&viewid=07a9ee37%2Dd793%2D45c0%2Da6b9%2D2a98315dbf3d")
GREENPROFI_URL = "https://my.greenprofi.de/cockpit"

# Streamlit interface for credentials


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
        page.wait_for_timeout(15000)  # Adjust as needed based on network speed and file size

        logging.info(f"File uploaded successfully to SharePoint: {local_file_path}")
        return True
    except Exception as e:
        logging.error(f"Error uploading file to SharePoint: {str(e)}")
        return False

def run(playwright: Playwright, browser_type: str, greenprofi_username: str, greenprofi_password: str, sharepoint_username: str, sharepoint_password: str) -> None:
    logging.info("Starting the script...")
    
    # Set up download path
    if os.name == 'nt':
        default_download_path = os.path.expanduser("~\\Downloads")
    else:
        default_download_path = os.path.expanduser("~//Downloads")
    
    custom_download_folder = os.path.join(default_download_path, "greenprofi_downloads")
    os.makedirs(custom_download_folder, exist_ok=True)

    # Launch the browser
    if browser_type == "chrome":
        browser = playwright.chromium.launch(headless=False) 
    elif browser_type == "edge":
        browser = playwright.chromium.launch(headless=False, channel="msedge")  
    else:
        raise ValueError("Unsupported browser type. Use 'chrome' or 'edge'.")
    
    # Create a new browser context
    context = browser.new_context(accept_downloads=True)
    
    # Create a new page
    page = context.new_page()
    
    # Navigate to GreenProfi website and download the file
    page.goto(GREENPROFI_URL)
    page.wait_for_load_state("networkidle")
    
    # Log in to GreenProfi
    page.get_by_role("textbox", name="E-Mail").fill(greenprofi_username)
    page.get_by_role("textbox", name="Passwort").fill(greenprofi_password)
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
        page.fill('input[type="email"]', sharepoint_username)
        page.click('input[type="submit"]')
        page.fill('input[type="password"]', sharepoint_password)
        page.click('input[type="submit"]')
        time.sleep(5)
        try:
            page.click('input[value="Yes"]')  # Stay signed in prompt
        except:
            pass

        if upload_to_sharepoint(page, final_path):
            save_last_extraction_date(datetime.datetime.now().date())
        else:
            logging.warning("Failed to upload to SharePoint. Last extraction date not updated.")

    context.close()
    browser.close()
    logging.info("Script execution finished.")

# Streamlit interface
st.title("GreenProfi Data Extraction Tool")
st.write("Use this app to extract data and upload it to SharePoint.")

# Input fields for credentials
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.sidebar.header("Login")
    username = st.sidebar.text_input("Username", type="default")
    password = st.sidebar.text_input("Password", type="password")
    # Verify against application credentials
    app_username = st.secrets["applogin"]["username"]
    app_password = st.secrets["applogin"]["password"]
    if st.sidebar.button("Login"):
        if username == app_username and password == app_password:
            st.session_state.authenticated = True
            st.sidebar.success("Login successful! Please proceed with the extraction.")  # Avoid calling experimental_rerun directly after setting session state, use a workaround to prevent error
        else:
            st.sidebar.error("Invalid application username or password")
else:
    # Extract Data button
    if st.button("Extract Data"):
        with st.spinner("Running extraction..."):
            with sync_playwright() as playwright:
                try:
                    run(playwright, browser_type="edge", 
                        greenprofi_username=st.secrets["greenprofi"]["username"], 
                        greenprofi_password=st.secrets["greenprofi"]["password"], 
                        sharepoint_username=st.secrets["sharepoint"]["username"], 
                        sharepoint_password=st.secrets["sharepoint"]["password"])
                    st.success("Data extraction completed successfully.")
                except Exception as e:
                    logging.error(f"Error during extraction: {str(e)}")
                    st.error(f"Error occurred: {str(e)}")

    # Logout button
    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.sidebar.success("Logged out successfully.")  # Avoid calling experimental_rerun directly after setting session state, use a workaround to prevent error