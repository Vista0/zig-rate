# 🏦 RBZ ZIG Rate Scraper  

A Python automation tool that extracts **ZIG rate** from the **Reserve Bank of Zimbabwe (RBZ)** website, parses PDF documents, and saves all results into a clean, structured Excel file.  

This project eliminates the need for manual PDF downloads and rate extraction — perfect for **financial analysts**, **economists**, and **data scientists** who need reliable exchange rate data.  

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
cd rbz-usd-scraper

