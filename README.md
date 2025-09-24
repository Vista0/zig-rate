# 🏦 RBZ ZIG-Rate Scraper  

This project automates the extraction of ZiG-rate from the Reserve Bank of Zimbabwe (RBZ) website. It uses Selenium for automated browsing, PyMuPDF (fitz) for PDF parsing, and Pandas to store the results in a structured Excel file.

The script loops through a list of specified months, filters the RBZ exchange rates page, extracts links to daily exchange rate PDFs, parses each PDF to retrieve the USD mid-rate, and saves all data in chronological order to an Excel file (all-rates.xlsx).

---

## 🚀 Features  

- ✅ **Automated Web Navigation** – Filters RBZ pages by month & year  
- ✅ **PDF Parsing** – Extracts USD mid-rate values from downloaded files  
- ✅ **Structured Data Output** – Saves results to `all-rates.xlsx`  
- ✅ **Multi-Month Support** – Scrapes multiple months in one run  
- ✅ **Error Handling** – Skips missing files gracefully  

---

## 🛠️ Tech Stack  

- **Python 3**  
- **Selenium** – For automated browser control  
- **PyMuPDF (fitz)** – For extracting text from PDF files  
- **Pandas** – For storing and cleaning extracted data  
- **Requests** – For downloading PDF files  

---

## 📦 Installation  

### 1️⃣ Clone the Repository  

```bash
git clone https://github.com/Vista0/zig-rate.git
cd zig-rate
```
### 2️⃣ Install Dependencies

```bash
pip install pandas requests selenium PyMuPDF
```
 ### ⚠️ Chrome Browser + ChromeDriver:
  Make sure you have Google Chrome installed and that ChromeDriver matches your browser version.
  Add ChromeDriver to your system PATH so Selenium can launch Chrome automatically.

## 🚀 Usage

Run the scraper from your terminal:

```bash
python scraper.py
```
This will retrieve all USD mid-rates for the months defined in months_to_scrape and save them to an Excel file named:
```text
all-rates.xlsx
```
## ⚙️ Configuration

You can modify which months/years to scrape by editing the `months_to_scrape` list in the script:

```python
months_to_scrape = [
    ("September", "2025"),
    ("August", "2025"),
    ("July", "2025"),
    ("June", "2025"),
    ("November", "2024"),
    ("October", "2024")
]
```

## 📄 Sample Output

After running the script, you’ll get an Excel file with the following structure:

| Date       | USD Mid Rate |
|-----------|-------------|
| 2025-09-01 | 26.7413 |
| 2025-09-02 | 26.7357 |
| 2025-09-03 | 26.7631|
| ...       | ... |

The file `all-rates.xlsx` will be sorted by date in ascending order.

## 🛠 Code Overview

**1. Imports Dependencies**  
Loads Selenium, PyMuPDF, Pandas, and Requests.

**2. Scrapes RBZ Website**  
Opens the exchange rate page, applies month/year filter, and extracts PDF links.

**3. Downloads and Parses PDFs**  
Extracts the USD mid-rate from each daily PDF file.

**4. Stores Results**  
Saves the data to a sorted Excel file for easy analysis.

## 💡 Practical Use Cases

Some practical use cases specifically for **USD–ZiG exchange rate data**:

- **Financial modeling:**  
  Use the data to forecast future USD–ZiG exchange rates for budgeting and pricing models.

- **International trade:**  
  Businesses can time their transactions based on historical exchange rate patterns.

- **Portfolio diversification:**  
  Investors holding USD assets can analyze historical trends to manage currency risk.

- **Accounting and tax reporting:**  
  Companies with USD-denominated transactions can use accurate daily zig-rates for financial statements and compliance.

- **Travel planning:**  
  Individuals can check historical rates to estimate currency conversion costs and plan better for trips abroad.


---
## ⚠️ Known Issues & Limitations

While this scraper is very useful, it does have some limitations you should be aware of:

- **RBZ Website Dependency:**  
  If the RBZ website structure changes (HTML layout, table format, or file links), the scraper may stop working until updated.

- **ChromeDriver Setup Required:**  
  You must have Google Chrome and the correct ChromeDriver installed and available in your PATH.  
  This may be tricky for beginners to set up correctly.

- **Execution Time:**  
  Since the script downloads and parses multiple PDFs, scraping large date ranges can take several minutes.

- **No Offline Mode:**  
  The scraper requires an active internet connection to access the RBZ website and download PDFs.

- **Error Handling is Basic:**  
  While missing files are skipped gracefully, advanced logging (to a file) is not yet implemented.

- **Zig-Only:**  
  Currently, the scraper only extracts **Zig-rate**. Other currencies would require modifying the parsing logic.

---

## 🚀 Future Improvements 

Planned enhancements to make the project even better:

- **Multi-Currency Support:**  
  Extend parsing logic to extract rates for other currencies listed in RBZ PDFs.

- **Headless Mode:**  
  Run Selenium in headless mode by default to speed up execution and avoid opening a visible Chrome window.

- **Improved Error Handling & Logging:**  
  Write detailed logs to a file for debugging and auditing purposes.

- **Command-Line Arguments:**  
  Allow users to pass months/years as arguments instead of editing the script manually.

- **Database Integration:**  
  Optionally store results in a SQLite or PostgreSQL database for more complex analysis and reporting.

- **Visualization:**  
  Generate basic charts (e.g., exchange rate trends over time) directly after scraping.

- **Docker Support:**  
  Package the project in a Docker container to simplify setup for any environment.

---



## 👤 Author

**Michael Nhete**  
🔗 [GitHub Profile](https://github.com/Vista0)  
📧 tinashemike68@gmail.com

