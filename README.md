Description:

This project automates the extraction of USD mid-rates from the Reserve Bank of Zimbabwe (RBZ) website. It uses Selenium for automated browsing, PyMuPDF (fitz) for PDF parsing, and Pandas to store the results in a structured Excel file.
The script loops through a list of specified months, filters the RBZ exchange rates page, extracts links to daily exchange rate PDFs, parses each PDF to retrieve the USD mid-rate, and saves all data in chronological order to an Excel file (all-rates.xlsx).

Key Features:

✅ Fully automated browsing of RBZ exchange rates page
✅ Extracts all daily USD mid-rates for given months
✅ Handles PDF downloads and text parsing
✅ Saves clean, structured, and sorted data to Excel
✅ Error handling for missing months or failed downloads

Tech Stack:

Python 3
Selenium (Chrome WebDriver)
PyMuPDF (fitz)
Pandas
Requests
