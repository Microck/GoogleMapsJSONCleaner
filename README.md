# JSON Processor & Excel Email Scraper

This project contains two Python scripts to process business data:

1.  **`cleaner_appender.py`**:
    *   **Input**: New JSON file(s) (e.g., from Apify Google Maps Extractor).
    *   **Action**: Cleans data, separates it into "MainData" and "AdditionalInfo", and appends it to a master Excel file (`XLXS_Consolidated/Consolidated_Business_Data.xlsx`). Creates the master file if it doesn't exist.
    *   **Use Case**: Consolidating data from multiple JSON extractions into one structured Excel workbook.

2.  **`excel_email_scraper.py`**:
    *   **Input**: An existing Excel file (like the one produced by `cleaner_appender.py`).
    *   **Action**: Prompts for the sheet and columns. For rows with a website URL but no valid email, it attempts to scrape the website to find an email address. Updates the input Excel file directly.
    *   **Use Case**: Enriching an existing Excel sheet with email addresses by scraping websites.

## Prerequisites
Python 3.x. Install required libraries:
```bash
pip install pandas openpyxl requests beautifulsoup4
```

## Setup & Configuration
Before running, you can adjust settings at the top of each script:
- **`cleaner_appender.py`**:
    - `MANDATORY_MAIN_FIELDS`, `SECOND_SHEET_FIELDS`, `UNNECESSARY_FIELDS`: Define which data fields go where or are removed.
    - `DESIRED_MAIN_COLUMN_ORDER`: Sets column order in the "MainData" sheet.
    - `MASTER_EXCEL_FILENAME`: Name of the consolidated output file.
- **`excel_email_scraper.py`**:
    - `REQUEST_TIMEOUT`, `REQUEST_DELAY`: Control scraping politeness.
    - `COMMON_CONTACT_PATHS`: URLs paths to check for emails (e.g., `/contact`).

## How to Run

1.  **Consolidate JSON Data (Optional, but recommended first):**
    ```bash
    python cleaner_appender.py
    ```
    Follow prompts to input your JSON file paths. This creates/updates `XLXS_Consolidated/Consolidated_Business_Data.xlsx`.

2.  **Scrape Emails from Excel:**
    ```bash
    python excel_email_scraper.py
    ```
    Follow prompts to specify the Excel file (e.g., the one from step 1), sheet name, website column, and email column.

## Notes
- **Email Scraping (`excel_email_scraper.py`)**:
    - Not 100% reliable; success depends on website structure and anti-bot measures.
    - Bypasses SSL errors (`verify=False`), which is less secure but often necessary for sites with certificate issues.
    - Be respectful of website terms and `robots.txt`
- **Output**:
    - `cleaner_appender.py` outputs to `XLXS_Consolidated/Consolidated_Business_Data.xlsx`.
    - `excel_email_scraper.py` modifies the Excel file you provide as input.
- **Customization**: Both scripts have clearly marked configuration sections at the top for tailoring field handling and scraping behavior.
