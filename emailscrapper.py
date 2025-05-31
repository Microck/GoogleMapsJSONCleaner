import json
import pandas as pd
import os
from pathlib import Path
import traceback
import requests
from bs4 import BeautifulSoup
import re
import time

# --- Previous Configuration (Keep as is or adjust) ---
MANDATORY_MAIN_FIELDS = {
    "imageUrl", "title", "totalScore", "reviewsCount", "street", "city",
    "state", "website", "phone", "categoryName", "url", "email" # Added email
}
SECOND_SHEET_FIELDS = {
    "claimThisBusiness", "permanentlyClosed", "temporarilyClosed",
    "openingHours", "additionalInfo", "countryCode",
}
UNNECESSARY_FIELDS = {
    "price", "neighborhood", "imageCategories", "scrapedAt", "googleFoodUrl",
    "hotelAds", "gasPrices", "searchPageUrl", "searchString", "language",
    "placeId", "cid", "fid", "kgmid", "imagesCount", "rank",
    "isAdvertisement", "phoneUnformatted", "reviewsDistribution",
    "peopleAlsoSearch", "placesTags", "reviewsTags",
}
DESIRED_MAIN_COLUMN_ORDER = [
    "title", "categoryName", "email", "totalScore", "reviewsCount", # Added email
    "street", "city", "state", "website", "phone", "imageUrl",
]
OUTPUT_SUBFOLDER = "XLXS"
MAIN_SHEET_NAME = "MainData"
EXTRA_SHEET_NAME = "AdditionalInfo"
# --- End Previous Configuration ---

# --- Web Scraping Configuration ---
REQUEST_TIMEOUT = 10  # seconds
REQUEST_DELAY = 1     # seconds between requests to the same domain
COMMON_CONTACT_PATHS = ["/contact", "/contact-us", "/about", "/about-us", "/impressum"]
EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
# --- End Web Scraping Configuration ---

def create_output_directory(folder_name):
    path = Path.cwd() / folder_name
    path.mkdir(parents=True, exist_ok=True)
    return path

def adjust_column_widths(writer, sheet_name, df):
    if df.empty: return
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df):
        series = df[col]
        max_len = (max((series.astype(str).map(len).max(), len(str(series.name)))) + 2)
        max_len = min(max_len, 70)
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=idx + 1).column_letter
        ].width = max_len

def find_emails_on_page(url):
    """Attempts to find emails on a single webpage URL."""
    emails_found = set()
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, timeout=REQUEST_TIMEOUT, headers=headers)
        response.raise_for_status() # Raise an exception for bad status codes
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Search in text
        text_content = soup.get_text(separator=" ")
        found_in_text = re.findall(EMAIL_REGEX, text_content)
        for email in found_in_text:
            emails_found.add(email.lower())

        # Search in mailto links
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if href.startswith("mailto:"):
                email = href.replace("mailto:", "").split("?")[0] # Remove params
                if re.match(EMAIL_REGEX, email):
                    emails_found.add(email.lower())
        
        return list(emails_found) if emails_found else None
    except requests.exceptions.RequestException as e:
        print(f"  Error fetching {url}: {e}")
        return None
    except Exception as e:
        print(f"  Error parsing {url}: {e}")
        return None

def scrape_website_for_email(base_website_url):
    """Scrapes a base website URL and common contact pages for an email."""
    if not base_website_url or not base_website_url.startswith(("http://", "https://")):
        return None

    print(f"  Scraping {base_website_url} for email...")
    
    # Normalize base URL (ensure it ends with / for proper joining)
    if not base_website_url.endswith('/'):
        base_website_url += '/'

    urls_to_check = [base_website_url]
    for path in COMMON_CONTACT_PATHS:
        # Properly join base URL and path
        try:
            from urllib.parse import urljoin
            full_url = urljoin(base_website_url, path.lstrip('/'))
            urls_to_check.append(full_url)
        except ImportError: # Should not happen in modern Python
            urls_to_check.append(base_website_url.rstrip('/') + path)


    for i, url_to_check in enumerate(list(set(urls_to_check))): # Check unique URLs
        if i > 0: # Add delay only for subsequent requests to the same domain
            time.sleep(REQUEST_DELAY)
        emails = find_emails_on_page(url_to_check)
        if emails:
            print(f"    Found email(s) on {url_to_check}: {emails[0]}")
            return emails[0] # Return the first one found
    print(f"    No email found on {base_website_url} or common contact pages.")
    return None


def process_json_file(
    json_file_path, mandatory_main_fields, unnecessary_fields,
    second_sheet_fields, desired_main_col_order, output_dir,
):
    all_main_data_rows = []
    all_extra_data_rows = []

    try:
        with open(json_file_path, "r", encoding="utf-8") as f:
            data_from_file = json.load(f)

        items_to_process = (
            [data_from_file] if isinstance(data_from_file, dict)
            else data_from_file if isinstance(data_from_file, list)
            else []
        )

        if not items_to_process:
            print(f"Warning: Content of {json_file_path.name} is not processable. Skipping.")
            return

        for item_idx, item in enumerate(items_to_process):
            if not isinstance(item, dict):
                print(f"Warning: Non-dictionary item at index {item_idx} in {json_file_path.name}. Skipping.")
                continue

            main_data_obj = {}
            extra_data_obj = {}

            # Initialize email field if it's expected
            if "email" in mandatory_main_fields:
                main_data_obj["email"] = item.get("email") # Get existing email if any

            # Scrape for email if not already present and website exists
            if not main_data_obj.get("email") and item.get("website"):
                scraped_email = scrape_website_for_email(item.get("website"))
                if scraped_email:
                    main_data_obj["email"] = scraped_email
            
            # Populate other fields
            for key, value in item.items():
                if key == "email" and "email" in main_data_obj and main_data_obj["email"]:
                    continue # Already handled/scraped
                if key in mandatory_main_fields:
                    main_data_obj[key] = value
                elif key in second_sheet_fields:
                    extra_data_obj[key] = value
                elif key not in unnecessary_fields:
                    if key not in main_data_obj:
                        main_data_obj[key] = value
            
            for m_key in mandatory_main_fields:
                if m_key not in main_data_obj:
                    main_data_obj[m_key] = item.get(m_key, None)

            if main_data_obj: all_main_data_rows.append(main_data_obj)
            if extra_data_obj: all_extra_data_rows.append(extra_data_obj)

        if not all_main_data_rows and not all_extra_data_rows:
            print(f"No processable data in {json_file_path.name}. No Excel file created.")
            return

        df_main = pd.DataFrame(all_main_data_rows)
        df_extra = pd.DataFrame(all_extra_data_rows)

        if not df_main.empty:
            current_main_cols = df_main.columns.tolist()
            ordered_cols = []
            for col in desired_main_col_order:
                if col in current_main_cols:
                    ordered_cols.append(col)
                    if col in current_main_cols: current_main_cols.remove(col) # Ensure removal
            
            url_col_present = "url" in current_main_cols
            if url_col_present: current_main_cols.remove("url")
            
            ordered_cols.extend(current_main_cols)
            
            if "url" in df_main.columns:
                if "url" not in ordered_cols: ordered_cols.append("url")
                elif ordered_cols[-1] != "url":
                    if "url" in ordered_cols: ordered_cols.remove("url") # Ensure removal before append
                    ordered_cols.append("url")
            
            df_main = df_main.reindex(columns=ordered_cols, fill_value=None)

        output_excel_filename = f"{json_file_path.stem}.xlsx"
        output_excel_path = output_dir / output_excel_filename

        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            if not df_main.empty:
                df_main.to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)
                adjust_column_widths(writer, MAIN_SHEET_NAME, df_main)
            elif not df_extra.empty:
                 pd.DataFrame().to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)

            if not df_extra.empty:
                df_extra.to_excel(writer, sheet_name=EXTRA_SHEET_NAME, index=False)
                adjust_column_widths(writer, EXTRA_SHEET_NAME, df_extra)
            elif not df_main.empty:
                 pd.DataFrame().to_excel(writer, sheet_name=EXTRA_SHEET_NAME, index=False)
        print(f"Successfully created Excel file: {output_excel_path}")

    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from {json_file_path.name}. Skipping.")
    except Exception as e:
        print(f"An unexpected error occurred while processing {json_file_path.name}: {e}")
        traceback.print_exc()

def main():
    output_directory_path = create_output_directory(OUTPUT_SUBFOLDER)
    print(f"Output will be saved to: {output_directory_path.resolve()}")

    json_file_paths_input = input(
        "Enter paths to JSON files, separated by commas (or a single path): "
    )
    if not json_file_paths_input:
        print("No file paths provided. Exiting.")
        return

    raw_paths = json_file_paths_input.split(",")
    json_file_paths = []
    for p_raw in raw_paths:
        p_stripped = p_raw.strip()
        p_cleaned = p_stripped[1:-1] if p_stripped.startswith('"') and p_stripped.endswith('"') else p_stripped
        if p_cleaned: json_file_paths.append(Path(p_cleaned))

    if not json_file_paths:
        print("No valid file paths extracted. Exiting.")
        return

    for file_path in json_file_paths:
        if not file_path.is_file():
            print(f"Warning: File not found at '{file_path}'. Skipping.")
            continue
        if file_path.suffix.lower() != ".json":
            print(f"Warning: File '{file_path}' is not a .json file. Skipping.")
            continue

        print(f"\nProcessing {file_path.name}...")
        process_json_file(
            file_path, MANDATORY_MAIN_FIELDS, UNNECESSARY_FIELDS,
            SECOND_SHEET_FIELDS, DESIRED_MAIN_COLUMN_ORDER, output_directory_path,
        )

if __name__ == "__main__":
    main()