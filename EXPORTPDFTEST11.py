import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright

LAST_DATE_FILE = "last_extraction_date.txt"

def read_last_extraction_date():
    """Read the last extraction date from file."""
    if os.path.exists(LAST_DATE_FILE):
        with open(LAST_DATE_FILE, "r") as file:
            date_str = file.read().strip()
            return datetime.datetime.strptime(date_str, "%d.%m.%Y")
    return None

def write_last_extraction_date(last_date):
    """Write the last extraction date to file."""
    with open(LAST_DATE_FILE, "w") as file:
        file.write(last_date.strftime("%d.%m.%Y"))

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=True)
    context = browser.new_context()
    page = context.new_page()
    
    print("Navigating website...")
    page.goto("https://my.greenprofi.de/cockpit")
    page.wait_for_load_state("networkidle")
    page.get_by_role("button", name="Nur Session-Cookies freigeben").click()
    page.wait_for_load_state("networkidle")

    print("Logging in...")
    page.get_by_role("textbox", name="E-Mail").fill("b.schmidt@ib-bauabrechnung.de")
    page.get_by_role("textbox", name="Passwort").fill("bydfuc-gowmyW-vexbo9")
    page.locator("[id=\"__BVID__81___BV_modal_body_\"]").get_by_role("button", name="Einloggen").click()
    page.wait_for_load_state("load")

    print("Navigating results...")
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.wait_for_selector('table')
    page.wait_for_timeout(10000)

    print("Checking projects...")
    project_entries = page.evaluate('''() => {
        const rows = Array.from(document.querySelectorAll('table tbody tr'));
        return rows.map(row => {
            const cells = Array.from(row.querySelectorAll('td'));
            const dateText = cells[0].textContent.trim();                        
            const projectLink = cells[2] ? cells[2].querySelector('a') : null;
            const projectId = projectLink ? projectLink.href.split('/')[2] : '';
            const submissionLink = projectLink ? projectLink.href : '';
            return {
                id: projectId,
                submissionDate: dateText,
                name: projectLink ? projectLink.textContent.trim() : '',
                submissionLink: submissionLink
            };
        }).filter(entry => entry.submissionDate && entry.id);
    }''')
    print(f"Found {len(project_entries)} project entries.")

    current_date = datetime.datetime.now()

    last_extraction_date = read_last_extraction_date()
    if last_extraction_date:
        print(f"Last extraction date: {last_extraction_date.strftime('%d.%m.%Y')}")
    else:
        print("No previous extraction date found.")
        last_extraction_date = current_date - datetime.timedelta(days=7)

    print(f"Filtering projects from {last_extraction_date.strftime('%d.%m.%Y')} to {current_date.strftime('%d.%m.%Y')}")

    for entry in project_entries:
        entry['parsed_date'] = datetime.datetime.strptime(entry['submissionDate'], "%d.%m.%Y")

    project_entries.sort(key=lambda x: x['parsed_date'], reverse=True)

    print("Creating downloads directory...")
    os.makedirs('downloads', exist_ok=True)

    latest_processed_date = last_extraction_date

    for index, entry in enumerate(project_entries, 1):
        print(f"\nProcessing project {index}/{len(project_entries)}")

        if entry['parsed_date'] <= last_extraction_date:
            print(f"Skipping project {entry['id']} - submitted on {entry['submissionDate']} which is already processed.")
            continue

        latest_processed_date = max(latest_processed_date, entry['parsed_date'])

        print(f"Processing project: {entry['name']} (ID: {entry['id']}) - Submitted on: {entry['submissionDate']}")
        page.goto(entry['submissionLink'])
        page.wait_for_timeout(2000)

        print("Extracting 'Entfernung' value...")
        entfernung = page.evaluate('''() => {
            const entfernungCell = document.querySelector('dd.col-sm-6:nth-of-type(1) span');
            return entfernungCell ? parseFloat(entfernungCell.textContent.replace(' km', '').trim()) : Infinity;
        }''')
        print(f"Entfernung: {entfernung} km")

        print("Extracting 'Summe' value...")
        summe = page.evaluate(r'''() => {
            const summeElement = Array.from(document.querySelectorAll('#project_location dt')).find(el => el.textContent.trim() === 'Summe');
            const summeCell = summeElement?.nextElementSibling?.querySelector('div');
            if (!summeCell) return 'Not found';
            
            const summeText = summeCell.textContent.trim();
            
            const convertGermanNumber = (numberString) => {
                const cleanNumber = numberString.replace('€', '').replace(/\s/g, '').trim();
                const parts = cleanNumber.split(',');
                const integerPart = parts[0].replace(/\./g, '');
                const fractionalPart = parts[1] || '0';
                return parseFloat(integerPart + '.' + fractionalPart);
            };

            const summeValue = convertGermanNumber(summeText);
            
            return {
                original: summeText,
                value: summeValue
            };
        }''')

        if summe != 'Not found':
            print(f"Summe: {summe['original']} ({summe['value']:.2f} €)")

            # Apply adjustment for VAT
            adjusted_summe = summe['value'] / 1.19
            print(f"Adjusted Summe (after VAT adjustment): {adjusted_summe:.2f} €")

        # Check both criteria:
        criteria1 = entfernung <= 100 and adjusted_summe >= 1500000
        criteria2 = entfernung >= 100 and adjusted_summe >= 30000000

        if criteria1 or criteria2:
            print("Project meets criteria. Proceeding to download...")
            page.locator(".pt-0 > ul > li:nth-child(4) > .nav-link").first.click()
            page.wait_for_selector('a:text("Exportieren")')

            print("Initiating PDF download...")
            with page.expect_download() as download_info:
                page.get_by_role("link", name="Exportieren").click()

                download = download_info.value
                pdf_filename = download.suggested_filename
                pdf_path = os.path.join('downloads', pdf_filename)
                download.save_as(pdf_path)

                print(f"Downloaded: {pdf_path}")
        else:
            print("Project does not meet criteria. Skipping download.")

        print("Returning to main table...")
        page.go_back()
        page.wait_for_selector('table')

    print("Finished processing all projects.")

    if latest_processed_date > last_extraction_date:
        print(f"Updating last extraction date to {latest_processed_date.strftime('%d.%m.%Y')}")
        write_last_extraction_date(latest_processed_date)

    context.close()
    browser.close()
    print("Script completed.")

print("Initializing Playwright...")
with sync_playwright() as playwright:
    run(playwright)
print("Script execution finished.")
