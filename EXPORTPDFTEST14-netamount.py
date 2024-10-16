import os
import datetime
import time
from playwright.sync_api import Playwright, sync_playwright

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
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

    print("Navigating results")
    page.get_by_text("Submissionsergebnisse", exact=True).click()
    page.wait_for_selector('table')
    page.wait_for_timeout(10000)

    print("Checking projects")
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
    print(f"Found {len(project_entries)}project entries.")

     
    current_date = datetime.datetime.now()
    thirty_days_ago = current_date - datetime.timedelta(days=30)   
    print(f"Filtering projects from {thirty_days_ago.strftime('%Y-%m-%d')} to {current_date.strftime('%Y-%m-%d')}")

    for entry in project_entries:
        entry['parsed_date'] = datetime.datetime.strptime(entry['submissionDate'], "%d.%m.%Y")

    project_entries.sort(key=lambda x: x['parsed_date'], reverse=True)

    print ("Creating downloads directory...")
    os.makedirs('downloads', exist_ok=True)

    for index, entry in enumerate(project_entries, 1):
        print(f"\nProcessing project {index}/{len(project_entries)}")

        if entry['parsed_date'] < thirty_days_ago:
            print(f"Skipping project {entry['id']} - submitted on {entry['submissionDate']} which is too old")
            continue

        
        print(f"Processing project: {entry['name']} (ID: {entry['id']}) - Submitted on: {entry['submissionDate']}")
        page.goto(entry['submissionLink'])
        page.wait_for_timeout(2000)
       

        print ("Extracting 'Entfernung' value...")
        entfernung = page.evaluate('''() => {
            const entfernungCell = document.querySelector('dd.col-sm-6:nth-of-type(1) span');
            return entfernungCell ? parseFloat(entfernungCell.textContent.replace(' km', '').trim()) : Infinity;
        }''')
        print(f"Entfernung: {entfernung} km")

        print("Extracting submission date...")
        submission_date = page.evaluate('''() => {
            const submissionCell = document.querySelector('dt.col-sm-6:nth-of-type(3) + dd.col-sm-6 div');
            return submissionCell ? submissionCell.textContent.trim() : 'Not found';
        }''')
        print(f"Submission Date: {submission_date}")

        summe = page.evaluate(r'''() => {
            const summeElement = Array.from(document.querySelectorAll('#project_location dt')).find(el => el.textContent.trim() === 'Summe');
            const summeCell = summeElement?.nextElementSibling?.querySelector('div');
            if (!summeCell) return 'Not found';
            
            const summeText = summeCell.textContent.trim();
            
            // Function to convert German number format to float
            const convertGermanNumber = (numberString) => {
                const cleanNumber = numberString.replace('€', '').replace(/\s/g, '').trim(); // Using regex correctly
                const parts = cleanNumber.split(',');
                const integerPart = parts[0].replace(/\./g, '');
                const fractionalPart = parts[1] || '0';
                return parseFloat(integerPart + '.' + fractionalPart);
            };

            const summeValue = convertGermanNumber(summeText);
            
            return {
                original: summeText,
                value: summeValue,
                meetsThreshold: summeValue >= 1500000
            };
        }''')

        if summe != 'Not found':
            print(f"Summe: {summe['original']}")
            print(f"Summe (as float): {summe['value']:.2f}")

            adjusted_summe = summe ['value'] / 1.19
            print(f"Adjusted Summe (divided by 1.19): {adjusted_summe}:.2f")

            Net_amount = adjusted_summe >= 1500000

            print(f"Net amount (>= 1,500,000€): {'Yes' if Net_amount else 'No'}")
        else:
            print("Summe: Not found")

        page.wait_for_timeout(5000)
        
        if entfernung <= 100 and summe != 'Not found' and summe['value'] >= 1500000:
           
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

                print(f"Waiting for download to complete: {pdf_filename}")
                max_wait_time = 30  
                start_time = time.time()
                while not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
                    if time.time() - start_time > max_wait_time:
                        print(f"Timeout: Failed to download {pdf_filename}")
                        break
                    time.sleep(0.5)
                
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    print(f"Downloaded: {pdf_path}")
                else:
                    print(f"Failed to download: {pdf_filename}")

        else:
            print(f"Project does not meet criteria. Skipping download.")        

        print("Returning to main table...")
        page.go_back()
        page.wait_for_selector('table')

    print("Finished processing all projects.")
    print("Cleaning up...")
    context.close()
    browser.close()
    print("Script completed.")

print("Initializing Playwright...")
with sync_playwright() as playwright:
    run(playwright)
print("Script execution finished.")