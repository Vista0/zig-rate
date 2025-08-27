import time
import fitz # PyMuPDF for PDF extraction
import pandas as pd # For data manipulation and Excel export
import requests # For downloading PDFs
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- ------------------------------- ---
# --- 1. SETUP SELENIUM HEADLESS BROWSER ---
# --- ------------------------------- ---
options = Options()
# options.add_argument("--headless")  # Runs Chrome in headless mode (no visible browser window)
options.add_argument("--disable-gpu") # Recommended for headless mode
# options.add_argument("--window-size=1920,1080") # Optional: set a window size for consistent rendering

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10) # Set a 10-second timeout for element interactions

# --- -------------------------------- ---
# --- 2. NAVIGATE TO RBZ EXCHANGE RATES PAGE ---
# --- -------------------------------- ---
print("Opening RBZ exchange rates page...")
driver.get("https://www.rbz.co.zw/index.php/research/markets/exchange-rates") # Correct URL confirmed [8, 9]
time.sleep(3) # Give the page some time to load completely

# --- ------------------------------- ---
# --- 3. FILTER BY TITLE (e.g., "August 2025") ---
# --- ------------------------------- ---
# We are switching to the "Title Filter" method as it's more robust [1, 2].
# The "Title Filter" input field has the ID 'filter-search' [1, 2].
try:
    print("Attempting to use the 'Title Filter'...")
    # Wait for the Title Filter input field to be present and clickable
    title_filter_input = wait.until(EC.element_to_be_clickable((By.ID, "filter-search")))
    
    # Send the desired month and year to the filter input [2]
    # IMPORTANT: Update this string for different months/years [10]
    title_filter_input.send_keys("August 2025") 

    # Wait for and click the "Filter" button to apply the search.
    # The button is a <button> element, not an <input>, with text "Filter" [user's previous turn]
    filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit" and text()="Filter"]')))
    filter_button.click()
    time.sleep(3) # Wait for the results to load after filtering
    print("Filter applied successfully for August 2025.")

except Exception as e:
    print(f"ERROR: Failed to interact with Title Filter or Filter button. Details: {e}")
    driver.quit()
    exit() # Exit the script if filtering fails as subsequent steps depend on it

# --- NEW STEP: Click the filtered month link to access daily PDFs ---
try:
    print("Clicking 'August 2025' link to open daily PDF page...")
    # Adjust this XPath as needed based on how the "August 2025" link is structured
    # It could be a simple <a> tag or within a <td> or <li>
    month_result_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "August 2025")))
    # If LINK_TEXT doesn't work, you might need a more specific XPath, e.g.:
    # month_result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[text()="August 2025"]')))
    # Or if it's within a table cell:
    # month_result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//table//td/a[text()="August 2025"]')))

    month_result_link.click()
    time.sleep(3) # Wait for the new page with daily PDFs to load
    print("Navigated to daily PDF links page.")

except Exception as e:
    print(f"ERROR: Failed to click the month result link. Details: {e}")
    driver.quit()
    exit() # Exit if navigation fails

# --- -------------------------------- ---
# --- 4. GET ALL PDF LINKS AND DATES ---
# --- -------------------------------- ---
print("Extracting PDF links and dates...")
pdf_links = []
# Find all table rows
rows = driver.find_elements(By.XPATH, '//table//tr')
for row in rows:
    try:
        # Extract the date from the first column of the row
        date = row.find_element(By.XPATH, './td[1]').text.strip()
        # Extract the PDF link from the anchor tag in the second column
        link = row.find_element(By.XPATH, './td[2]//a').get_attribute("href")
        
        if link and link.endswith(".pdf"):
            pdf_links.append((date, link))
    except Exception as e:
        # Some rows might not contain valid date/link, so we skip them
        # print(f"Skipping row due to error: {e}") # Uncomment for debugging row issues
        continue
driver.quit() # Close the browser once all links are extracted

# --- ------------------------------------ ---
# --- 5. FUNCTION TO EXTRACT USD MID RATE FROM PDF ---
# --- ------------------------------------ ---
def extract_usd_midrate(pdf_url):
    try:
        # Download the PDF file
        response = requests.get(pdf_url)
        with open("temp.pdf", "wb") as f:
            f.write(response.content)
        print(f"🔍 DEBUG: Successfully downloaded content for {pdf_url} to temp.pdf")

        # Open the PDF with PyMuPDF (fitz) and extract text
         # Open the PDF with PyMuPDF (fitz) and extract text
        with fitz.open("temp.pdf") as doc:
            for page in doc:
                text = page.get_text()
                lines = text.split("\n")
                usd_block_active = False
                potential_rates = []

                for line_content in lines:
                    stripped_line = line_content.strip()

                    # Check if we are in the USD block of information
                    if stripped_line == "USD":
                        usd_block_active = True
                        potential_rates = [] # Reset for this currency block
                        continue # Move to the next line

                    if usd_block_active:
                        # If we find another currency or a header, assume the USD block has ended
                        if stripped_line and not stripped_line.replace('.', '', 1).isdigit():
                            # If we've collected numbers, the last one is the mid-rate
                            if potential_rates:
                                return potential_rates[-1]
                            usd_block_active = False # Reset if new non-numeric line starts a new block
                            continue # Continue searching for "USD" for other sections if any

                        # Try to convert subsequent non-empty lines to a float
                        if stripped_line: # Only process non-empty lines
                            try:
                                num = float(stripped_line)
                                potential_rates.append(num)
                            except ValueError:
                                # Not a number, could be an unexpected text or separator.
                                # If we have collected numbers, assume we've passed the data.
                                if potential_rates:
                                    return potential_rates[-1]
                                # If it's not a number and no rates collected, and not empty,
                                # it means this isn't part of the USD data block, so deactivate
                                usd_block_active = False

                # If the loop finishes and we were in a USD block and collected rates, return the last one
                if usd_block_active and potential_rates:
                    return potential_rates[-1]

    except requests.exceptions.RequestException as req_e:
        print(f"Error downloading PDF from {pdf_url}: {req_e}")
    except Exception as e:
        print(f"Error extracting from PDF {pdf_url}: {e}")
    return None # Return None if extraction fails

# --- ------------------------------- ---
# --- 6. SCRAPE EACH PDF FOR USD RATE ---
# --- ------------------------------- ---
print(f"Processing {len(pdf_links)} PDF documents...")
records = []
for date, link in pdf_links:
    print(f"📄 Processing {date}...")
    rate = extract_usd_midrate(link)
    if rate:
        # Ensure date format is YYYY-MM-DD for consistency and sorting [14, 16, 17]
        records.append({
            "Date": f"2025-08-{date.zfill(2)}", 
            "USD Mid Rate": rate
        })

# --- ------------------------------- ---
# --- 7. SAVE TO EXCEL FILE ---
# --- ------------------------------- ---
df = pd.DataFrame(records)

# Only sort and save if there's data to prevent KeyError [18, 19]
if not df.empty:
    df.sort_values("Date", inplace=True)
    # IMPORTANT: Update the output filename for different months/years [10]
    df.to_excel("usd_mid_rates_august_2025.xlsx", index=False)
    print("\n✅ Done! Data saved to 'usd_mid_rates_august_2025.xlsx'")
else:
    print("\n⚠️ No data extracted — check if PDFs exist for the filtered month or if the format has changed.")