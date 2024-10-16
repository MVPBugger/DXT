import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright, expect

def parse_euro_amount(amount_str):
    if not amount_str or amount_str.isspace():
        print("Warning: Empty or whitespace-only Euro amount encountered")
        return 0.0

    try:
        # Remove currency symbols and trim whitespace
        cleaned = amount_str.replace('€', '').replace(' ', '').strip()

        # Handle decimal and thousand separators
        if ',' in cleaned:
            # Split by the last comma to separate whole and decimal parts
            parts = cleaned.rsplit(',', 1)
            integer_part = parts[0]  # The part before the last comma
            decimal_part = parts[1] if len(parts) > 1 else ''  # The part after the last comma

            # Remove all dots (thousands separators) from the integer part
            integer_part = integer_part.replace('.', '')

            # Combine cleaned parts
            cleaned = f"{integer_part}.{decimal_part}"  # Combine whole and decimal parts

        else:
            # If no comma, just remove dots (thousands separators)
            cleaned = cleaned.replace('.', '')

        # Check for any invalid characters (only digits and one dot)
        if not cleaned.replace('.', '').isdigit():
            print(f"Warning: Invalid characters found in '{amount_str}' after cleaning.")
            return 0.0

        # Convert to float
        return float(cleaned)
    except ValueError as e:
        print(f"Failed to parse Euro amount: '{amount_str}' with error: {e}")
        return 0.0



def run(playwright: Playwright, browser_type: str) -> None:
    print("Starting the script...")

    # Set up download path
    print("Setting up download path...")
    if os.name == 'nt':
        default_download_path = os.path.expanduser("~\\Downloads")
    else:
        default_download_path = os.path.expanduser("~/Downloads")
    
    custom_download_folder = os.path.join(default_download_path, "greenprofi_downloads")
    os.makedirs(custom_download_folder, exist_ok=True)
    print(f"Custom download folder created: {custom_download_folder}")

    print(f"Launching {browser_type} browser...")
    if browser_type == "chrome":
        browser = playwright.chromium.launch(headless=False) 
    elif browser_type == "edge":
        browser = playwright.chromium.launch(headless=False, channel="msedge")  
    else:
        raise ValueError("Unsupported browser type. Use 'chrome' or 'edge'.")
    
    print("Creating new context...")
    context = browser.new_context(
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )                              
    
    print("Creating new page...")
    page = context.new_page()
    
    print("Navigating to greenprofi.de...")
    page.goto("https://my.greenprofi.de/cockpit")
    page.wait_for_load_state("networkidle")
    
    print("Accepting cookies...")
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    page.wait_for_load_state("networkidle")

    print("Logging in...")
    page.get_by_role("textbox", name="E-Mail").fill("b.schmidt@ib-bauabrechnung.de")
    page.get_by_role("textbox", name="Passwort").fill("bydfuc-gowmyW-vexbo9")
    page.locator("[id=\"__BVID__81___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()
    page.wait_for_load_state("load")
    
    print("Navigating to Submissionsergebnisse...")
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    print("Waiting table to be visible...")
    page.wait_for_selector('table tbody tr', state='visible')
    time.sleep(2)
    
   
    print("Extracting and evaluating projects...")
    projects = page.evaluate("""
        () => {
            const rows = Array.from(document.querySelectorAll('table tbody tr'));
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td'));
                if (cells.length < 9) {  // We now expect at least 9 cells
                    console.warn('Row has fewer than expected cells:', row.innerHTML);
                    return null;
                }
                const linkElement = cells[2].querySelector('a');
                const nameElement = linkElement ? linkElement.querySelector('span') : null;
                return {
                    submissionDate: cells[0]?.textContent?.trim() || '',
                    erfasst: cells[1]?.textContent?.trim() || '',
                    name: nameElement ? nameElement.textContent.trim() : '',
                    link: linkElement ? linkElement.href : '',
                    summe: cells[6]?.textContent?.trim() || '0 €',
                    entfernung: cells[7]?.textContent?.trim() || '0',
                };
            }).filter(project => project !== null);
        }
    """)
    print(f"Found {len(projects)} projects. Processing and filtering based on criteria...")

    current_date = datetime.datetime.now()
    seven_days_ago = current_date - datetime.timedelta(days=7)
    filtered_projects = []

    # Inside your run function, where you filter projects
    for project in projects:
        try:
            submission_date = datetime.datetime.strptime(project['submissionDate'], '%d.%m.%Y')
            project_sum = parse_euro_amount(project['summe'])
            project_distance = float(project['entfernung'].replace(' km', ''))
            
            if (submission_date >= seven_days_ago and
                project_distance <= 100 and
                project_sum >= 1500000):
                filtered_projects.append(project)
                time.sleep(5)
            
            print(f"Project: Submission Date={project['submissionDate']}, Erfasst={project['erfasst']}, Name={project['name'][:30]}..., Distance={project_distance}km, Sum={project_sum}€")
        except ValueError as e:
            print(f"Error processing project: {e}")
            print(f"Project data: {project}")



    print(f"Filtered down to {len(filtered_projects)} projects meeting criteria.")

    # Download Excel for filtered projects 
     
    if filtered_projects:
        print("Initiating Excel download for filtered projects...")      
        with page.expect_download() as download_info:
             
             page.locator("span:nth-child(5) > a").first.click()
             page.get_by_text("Leistungsumfang", exact=True).click()
             page.get_by_text("Baubeginn/Ende").click()
             page.locator("label").filter(has_text="25").click()
             page.get_by_role("link", name="Exportieren").click()

        download = download_info.value
        print(f"Download started: {download.suggested_filename}")
        time.sleep(5)
        
        final_path = os.path.join(custom_download_folder, download.suggested_filename)
        download.save_as(final_path)
        time.sleep(2)
        print(f"Download completed: {final_path}")
    
        if download.suggested_filename.endswith('.xls') or download.suggested_filename.endswith('.xlsx'):
            print(f"Successfully downloaded Excel file at: {final_path}")
        else:
            print(f"Downloaded file is not an Excel file: {final_path}")  

    else:
        print("No projects met the criteria for download.")

        time.sleep(3)

    print("Closing browser...")
    context.close()
    browser.close()

print("Initializing Playwright...")
with sync_playwright() as playwright:
    run(playwright, browser_type="edge")
print("Script execution finished.")