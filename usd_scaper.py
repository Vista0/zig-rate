
import time
import fitz 
import pandas as pd 
import requests 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


months_to_scrape = [
   ("August", "2025"),
    ("July", "2025"),
    ("June", "2025"),
    ("November", "2024"),
    ("October", "2024"),
    ("September", "2024"),
    ("August", "2024"),
    ("July", "2024"),
    ("June", "2024"),
    ("May", "2024"),
    ("April", "2024")
]
all_collected_records = []

def extract_usd_midrate(pdf_url):
    try:
        # Download the PDF 
        response = requests.get(pdf_url)
        with open("temp.pdf", "wb") as f:
            f.write(response.content)
        # Open the PDF 
        with fitz.open("temp.pdf") as doc:
            for page in doc:
                text = page.get_text()
                lines = text.split("\n")
                usd_block_active = False
                potential_rates = []
                for line_content in lines:
                    stripped_line = line_content.strip()
                    
                    if stripped_line == "USD":
                        usd_block_active = True
                        potential_rates = [] 
                        continue
                        

                    if usd_block_active:
                        
                        if stripped_line and not stripped_line.replace('.', '', 1).isdigit():
                            if potential_rates:
                                return potential_rates[-1] 
                            usd_block_active = False 
                            continue 
                        
                        if stripped_line: 
                            try:
                                num = float(stripped_line)
                                potential_rates.append(num)
                            except ValueError:
                                
                                if potential_rates:
                                    return potential_rates[-1] 
                                
                                usd_block_active = False
                
                if usd_block_active and potential_rates:
                    return potential_rates[-1]
    except requests.exceptions.RequestException as req_e:
        print(f"Error downloading PDF from {pdf_url}: {req_e}")
    except Exception as e:
        print(f"Error extracting from PDF {pdf_url}: {e}")
    return None 


for month_name, year in months_to_scrape:
    full_month_string = f"{month_name} {year}"
    print(f"\n--- Starting extraction for {full_month_string} ---")

    
    options = Options()
    options.add_argument("--disable-gpu") 
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10) 

    
    print(f"Opening RBZ exchange rates page for {full_month_string}...")
    
    driver.get("https://www.rbz.co.zw/index.php/research/markets/exchange-rates")
    time.sleep(3) 

    
    try:
        print(f"Attempting to use the 'Title Filter' for '{full_month_string}'...")
        
        title_filter_input = wait.until(EC.element_to_be_clickable((By.ID, "filter-search")))
        title_filter_input.clear() 
        title_filter_input.send_keys(full_month_string)
        filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit" and text()="Filter"]')))
        filter_button.click()
        time.sleep(3) 
        print(f"Filter applied successfully for {full_month_string}.")
    except Exception as e:
        print(f"ERROR: Failed to interact with Title Filter or Filter button for {full_month_string}. Details: {e}")
        driver.quit() 
        continue 

    
    try:
        print(f"Clicking '{full_month_string}' link to open daily PDF page...")
        month_result_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, full_month_string)))
        month_result_link.click()
        time.sleep(3) 
        print("Navigated to daily PDF links page.")
    except Exception as e:
        print(f"ERROR: Failed to click the month result link for {full_month_string}. Details: {e}")
        driver.quit()
        continue 

    # --- 4. GET ALL PDF LINKS AND DATES 
    print(f"Extracting PDF links and dates for {full_month_string}...")
    pdf_links = []
    # Find all table rows (tr elements) within tables
    rows = driver.find_elements(By.XPATH, '//table//tr')

    
    print(f"DEBUG (Pre-loop): Found {len(rows)} potential table rows.")


    for row in rows:
        try:
            
            date_text = row.find_element(By.XPATH, './td[1]').text.strip()
            link = row.find_element(By.XPATH, './td[2]//a').get_attribute("href")
            if link and link.endswith(".pdf"):
                pdf_links.append((date_text, link))
        except Exception as e:
            continue
    driver.quit() 

    # --- 6. SCRAPE EACH PDF FOR USD RATE --- [17]
    print(f"Processing {len(pdf_links)} PDF documents for {full_month_string}...")
    
    # Map month names to numbers for YYYY-MM-DD format
    month_number_map = {
        "January": "01", "February": "02", "March": "03", "April": "04",
        "May": "05", "June": "06", "July": "07", "August": "08",
        "September": "09", "October": "10", "November": "11", "December": "12"
    }
    current_month_number = month_number_map.get(month_name)

    for date_text, link in pdf_links:
        print(f"📄 Processing {date_text}...")
        rate = extract_usd_midrate(link)
        if rate:
            if current_month_number:
                 all_collected_records.append({
                    "Date": f"{year}-{current_month_number}-{date_text.zfill(2)}",
                    "USD Mid Rate": rate
                })
            else:
                print(f"WARNING: Could not determine month number for {month_name}. Skipping record for {date_text}.")
    print(f"DEBUG: After {full_month_string}, total records = {len(all_collected_records)}")

    # --- 7. SAVE TO EXCEL FILE --- [18]
    df = pd.DataFrame( all_collected_records)
    if not df.empty:
        df.sort_values("Date", inplace=True)
        df.to_excel('all-rates.xlsx', index=False)
        print(f"\n✅ Done! Data saved to 'all-rates.xlsx'")
    else:
        print(f"\n⚠️ No data extracted for {full_month_string} — check if PDFs exist for the filtered month or if the format has changed.")
