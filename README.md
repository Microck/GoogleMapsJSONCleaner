# JSON Data Processor & Email Scraper

This repository contains two Python scripts for processing JSON data (primarily designed for output from sources like the Apify Google Maps Extractor) and converting it into structured Excel (XLSX) files:

1.  **`cleaner.py`**: Cleans, restructures, and organizes JSON data into a multi-sheet Excel file.
2.  **`emailscrapper.py`**: Includes all functionalities of `cleaner.py` **plus** an attempt to scrape websites listed in the data to find email addresses if they are not already present in the input JSON.

Choose the script that best fits your needs. If you only need to clean and restructure existing JSON data, `cleaner.py` is sufficient and faster. If you also want to attempt to find missing email addresses by scraping websites, use `emailscrapper.py`.

## Table of Contents
1.  [Features Common to Both Scripts](#features-common-to-both-scripts)
2.  [Specific Features of `emailscrapper.py`](#specific-features-of-emailscrapperpy)
3.  [Prerequisites](#prerequisites)
4.  [Configuration](#configuration)
    *   [Core Cleaning & Structuring (Both Scripts)](#core-cleaning--structuring-both-scripts)
    *   [Web Scraping Specific (`emailscrapper.py` only)](#web-scraping-specific-emailscrapperpy-only)
5.  [How to Run](#how-to-run)
6.  [Input](#input)
7.  [Output](#output)
8.  [Important Notes & Limitations for `emailscrapper.py`](#important-notes--limitations-for-emailscrapperpy)

---

## 1. Features Common to Both Scripts
- **Data Cleaning**: Removes predefined unnecessary fields from the JSON data.
- **Multi-Sheet Excel Output**:
    - Moves a specific set of fields to a separate "AdditionalInfo" sheet.
    - Ensures a defined set of mandatory fields are present in the primary "MainData" sheet.
- **Custom Column Ordering**: Allows custom ordering of columns in the "MainData" sheet for better readability.
- **Flexible Input**: Handles single or multiple JSON file inputs provided at runtime.
- **Organized Output**:
    - Automatically creates an `XLXS` subfolder (in the same directory as the script) for output files if it doesn't already exist.
    - Names output XLSX files based on their corresponding input JSON file names (e.g., `data.json` becomes `data.xlsx`).
- **Formatted Excel**: Auto-adjusts column widths in the generated Excel sheets to fit content.

---

## 2. Specific Features of `emailscrapper.py`
Includes all features of `cleaner.py`, plus:

- **Email Scraping**: If an "email" field is not present in the input JSON for an item but a "website" URL is available, the script will:
    - Attempt to visit the main website URL.
    - Attempt to visit a predefined list of common contact-related pages (e.g., `/contact`, `/about-us`, `/contacto`).
    - Extract the first email address found using regular expressions from the page content or `mailto:` links.
- **SSL Bypass**: Bypasses SSL certificate verification errors for websites by using `verify=False` with the `requests` library. This helps access sites with misconfigured SSL but means the connection's security is not fully verified.
- **Polite Scraping**: Includes a configurable delay between web requests to the same domain.

---

## 3. Prerequisites
Ensure you have Python installed on your system.

**For `cleaner.py`:**
- `pandas`
- `openpyxl`

**For `emailscrapper.py` (includes `cleaner.py` dependencies):**
- `pandas`
- `openpyxl`
- `requests`
- `beautifulsoup4`
- `urllib3` (typically installed as a dependency of `requests`)

You can install these dependencies using pip. For `emailscrapper.py`, run:
```bash
pip install pandas openpyxl requests beautifulsoup4
```
If you only intend to use `cleaner.py`, you can install fewer packages:
```bash
pip install pandas openpyxl
```

---

## 4. Configuration
Before running either script, you can (and likely should) customize its behavior by editing the configuration variables defined at the top of the respective `.py` file.

### Core Cleaning & Structuring (Both Scripts)
These settings are present in both `cleaner.py` and `emailscrapper.py`. Ensure they are consistent if you use both, or tailor them to each script's purpose.

- **`MANDATORY_MAIN_FIELDS`**: A Python `set` of field names that *must* be included in the "MainData" sheet.
    *For `emailscrapper.py`, ensure `"email"` is included here if you want scraped emails to appear.*
- **`SECOND_SHEET_FIELDS`**: A Python `set` of field names that will be moved to the "AdditionalInfo" sheet (and removed from "MainData" unless also in `MANDATORY_MAIN_FIELDS`).
- **`UNNECESSARY_FIELDS`**: A Python `set` of field names that will be completely removed from the output (unless they are in `MANDATORY_MAIN_FIELDS`).
- **`DESIRED_MAIN_COLUMN_ORDER`**: A Python `list` defining the preferred order of columns for the "MainData" sheet. The `url` field is specially handled to appear last if present.
    *For `emailscrapper.py`, include `"email"` in your desired position.*
- **`OUTPUT_SUBFOLDER`**: The name of the subfolder where output XLSX files will be saved (default: "XLXS").
- **`MAIN_SHEET_NAME`**: The name for the primary data sheet in the Excel file (default: "MainData").
- **`EXTRA_SHEET_NAME`**: The name for the sheet containing additional information (default: "AdditionalInfo").

### Web Scraping Specific (`emailscrapper.py` only)
These settings are only found in `emailscrapper.py`:

- **`REQUEST_TIMEOUT`**: Max time (seconds) to wait for a server response during scraping.
- **`REQUEST_DELAY`**: Pause (seconds) between requests to the same domain during scraping.
- **`COMMON_CONTACT_PATHS`**: List of URL paths (e.g., "/contact") to check for contact info on websites.
- **`EMAIL_REGEX`**: Regular expression used to identify email addresses.

---

## 5. How to Run
1.  Save the desired script (`cleaner.py` or `emailscrapper.py`) to a directory on your computer.
2.  Open a terminal or command prompt.
3.  Navigate to the directory where you saved the script.
4.  Run the script using the appropriate command:
    For the cleaner:
    ```bash
    python cleaner.py
    ```
    For the cleaner with email scraping:
    ```bash
    python emailscrapper.py
    ```
5.  The script will prompt you to: `Enter the paths to your JSON files, separated by commas (or a single path):`
    -   Provide the full or relative path(s) to your JSON file(s).
    -   If providing multiple paths, separate them with a comma.
    -   Paths enclosed in double quotes (e.g., `"D:\My Files\data.json"`) are also supported.

---

## 6. Input
- Comma-separated file paths to JSON files.
- Each JSON file is expected to contain either a single JSON object or a list of JSON objects.
- For `emailscrapper.py` to attempt email finding, items in the JSON should ideally have a "website" key with a valid URL.

---

## 7. Output
- For each input JSON file (e.g., `my_data.json`), a corresponding Excel file (e.g., `my_data.xlsx`) will be created in the `XLXS` subfolder (located in the same directory as the script).
- Each Excel file will contain two sheets:
    - **"MainData"**: Contains the primary, cleaned, and reordered data. If using `emailscrapper.py`, this sheet will include an "email" column (populated from the original JSON or via web scraping).
    - **"AdditionalInfo"**: Contains fields specified in the `SECOND_SHEET_FIELDS` configuration.

---

## 8. Important Notes & Limitations for `emailscrapper.py`
- **Scraping Reliability**: Email scraping is not a perfect science. Websites vary greatly in structure, may use JavaScript to load email addresses (which this script's basic scraping doesn't execute), or employ anti-scraping measures. Success rates will vary.
- **SSL Verification**: `verify=False` is used for broader site access during scraping but means the security of the connection to those websites is not cryptographically verified.
- **Ethical Considerations**: Be mindful of website `robots.txt` files and terms of service. This script is intended for small-scale, considerate use. Do not use it to overwhelm servers.
- **Performance**: Web scraping adds significant time to the processing, especially due to network requests, timeouts, and the `REQUEST_DELAY`. Be patient when processing large files or many websites.
- **Error Handling**: The script includes basic error handling for common network issues and parsing problems during scraping. However, complex site structures or robust anti-bot measures can still lead to failures for specific sites. Check the console output for error messages.
