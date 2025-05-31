import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
from pathlib import Path
from urllib.parse import urljoin
import urllib3 # To suppress InsecureRequestWarning
import traceback # For more detailed error messages

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Web Scraping Configuration ---
REQUEST_TIMEOUT = 15
REQUEST_DELAY = 1 # Seconds between requests to the same domain
COMMON_CONTACT_PATHS = ["/contact", "/contact-us", "/contacto", "/about", "/about-us", "/impressum", "/legal", "/aviso-legal"]
EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
# --- End Web Scraping Configuration ---

def find_emails_on_page(url):
    """Attempts to find emails on a single webpage URL."""
    emails_found = set()
    try:
        headers = {"User-Agent": USER_AGENT}
        response = requests.get(url, timeout=REQUEST_TIMEOUT, headers=headers, verify=False, allow_redirects=True)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")
        
        text_content = soup.get_text(separator=" ")
        found_in_text = re.findall(EMAIL_REGEX, text_content)
        for email in found_in_text:
            emails_found.add(email.lower())

        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if href.startswith("mailto:"):
                email_candidate = href.replace("mailto:", "").split("?")[0]
                if re.match(EMAIL_REGEX, email_candidate):
                    emails_found.add(email_candidate.lower())
        
        return list(emails_found) if emails_found else None
    except requests.exceptions.RequestException as e:
        print(f"  Error fetching {url}: {e}")
        return None
    except Exception as e:
        print(f"  Error parsing {url}: {e}")
        return None

def scrape_website_for_email(base_website_url):
    """Scrapes a base website URL and common contact pages for an email."""
    if not base_website_url or not isinstance(base_website_url, str) or not base_website_url.strip().startswith(("http://", "https://")):
        print(f"  Invalid or missing website URL: {base_website_url}")
        return None

    print(f"  Scraping {base_website_url} for email...")
    
    normalized_base_url = base_website_url.strip()
    if not normalized_base_url.endswith('/'):
        normalized_base_url += '/'

    urls_to_check = {normalized_base_url}
    for path in COMMON_CONTACT_PATHS:
        try:
            full_url = urljoin(normalized_base_url, path.lstrip('/'))
            urls_to_check.add(full_url)
        except Exception as e:
            print(f"    Error creating URL for path {path} with base {normalized_base_url}: {e}")
            continue

    for i, url_to_check in enumerate(list(urls_to_check)):
        if i > 0:
            time.sleep(REQUEST_DELAY)
        emails = find_emails_on_page(url_to_check)
        if emails:
            print(f"    Found email(s) on {url_to_check}: {emails[0]}")
            return emails[0]
    print(f"    No email found on {base_website_url} or common contact pages.")
    return None

def adjust_column_widths(writer, sheet_name, df):
    """Adjusts column widths for a given sheet in the ExcelWriter object."""
    if df.empty:
        return
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df): # Iterate through columns
        series = df[col]
        max_len = (
            max(
                (
                    series.astype(str).map(len).max(), # len of largest item
                    len(str(series.name)), # len of column name/header
                )
            )
            + 2
        ) # Adding a little extra space
        max_len = min(max_len, 70) # Cap max width
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=idx + 1).column_letter
        ].width = max_len

def main_excel_scraper():
    excel_file_path_input = input("Enter the path to your Excel file: ").strip()
    if excel_file_path_input.startswith('"') and excel_file_path_input.endswith('"'):
        excel_file_path_input = excel_file_path_input[1:-1]
    
    excel_file_path = Path(excel_file_path_input)

    if not excel_file_path.is_file() or excel_file_path.suffix.lower() not in ['.xlsx', '.xls']:
        print(f"Error: File not found or not a valid Excel file: {excel_file_path}")
        return

    # --- Step 1: Read all existing sheets into a dictionary of DataFrames ---
    all_sheets_data = {}
    target_sheet_name_from_user = None
    try:
        xls = pd.ExcelFile(excel_file_path)
        sheet_names_in_file = xls.sheet_names
        print("\nSheets found in the Excel file:")
        for s_name in sheet_names_in_file:
            print(f"- {s_name}")

        if "MainData" in sheet_names_in_file:
            target_sheet_name_from_user = "MainData"
            print("Defaulting to sheet 'MainData'.")
        else:
            target_sheet_name_from_user = input("Enter the name of the sheet to process: ").strip()

        if target_sheet_name_from_user not in sheet_names_in_file:
            print(f"Error: Sheet '{target_sheet_name_from_user}' not found in the Excel file.")
            return
            
        for s_name in sheet_names_in_file:
            all_sheets_data[s_name] = pd.read_excel(xls, sheet_name=s_name)
        
        df_to_modify = all_sheets_data[target_sheet_name_from_user]

    except Exception as e:
        print(f"Error reading Excel file or sheets: {e}")
        traceback.print_exc()
        return

    print(f"\nColumns in sheet '{target_sheet_name_from_user}':")
    for i, col in enumerate(df_to_modify.columns):
        print(f"{i+1}. {col}")

    while True:
        try:
            website_col_num = int(input("Enter the number of the column containing website URLs: "))
            website_col_name = df_to_modify.columns[website_col_num - 1]
            break
        except (ValueError, IndexError):
            print("Invalid input. Please enter a valid column number.")
    
    email_col_name = input(f"Enter the name for the email column (or existing one to update, e.g., 'email'): ").strip()

    if email_col_name not in df_to_modify.columns:
        df_to_modify[email_col_name] = pd.NA # Use pd.NA for missing values
        print(f"Created new column: '{email_col_name}'")

    print(f"\nStarting email scraping. This may take a while...")
    found_any_new_email = False

    for index, row in df_to_modify.iterrows():
        website_url = row.get(website_col_name)
        current_email = row.get(email_col_name)

        should_scrape = False
        if pd.notna(website_url) and isinstance(website_url, str) and website_url.strip():
            if pd.isna(current_email) or not (isinstance(current_email, str) and re.match(EMAIL_REGEX, current_email)):
                should_scrape = True
        
        if should_scrape:
            print(f"\nProcessing row {index + 2} (Website: {website_url})...") # Excel rows are 1-indexed, +1 for header
            scraped_email = scrape_website_for_email(website_url)
            if scraped_email:
                df_to_modify.loc[index, email_col_name] = scraped_email
                found_any_new_email = True
                print(f"  Updated row {index + 2} with email: {scraped_email}")
            else:
                print(f"  No email found for row {index + 2}.")
        elif pd.notna(website_url) and website_url.strip():
            print(f"Skipping row {index + 2}: Email '{current_email}' already present and looks valid, or no website URL.")
        elif not (pd.notna(website_url) and isinstance(website_url, str) and website_url.strip()):
             print(f"Skipping row {index + 2}: No valid website URL.")


    if found_any_new_email:
        try:
            # --- Step 2: Update the target sheet's DataFrame in our dictionary ---
            all_sheets_data[target_sheet_name_from_user] = df_to_modify

            # --- Step 3: Write all sheets back to the Excel file ---
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                for sheet_name_to_write, df_to_write in all_sheets_data.items():
                    df_to_write.to_excel(writer, sheet_name=sheet_name_to_write, index=False)
                    adjust_column_widths(writer, sheet_name_to_write, df_to_write) # Adjust widths for all sheets

            print(f"\nSuccessfully updated Excel file: {excel_file_path}")
        except Exception as e:
            print(f"\nError writing updated Excel file: {e}")
            traceback.print_exc()
            print("Consider saving the DataFrame to a new file manually if needed.")
    else:
        print("\nNo new emails were found or updated.")

if __name__ == "__main__":
    main_excel_scraper()