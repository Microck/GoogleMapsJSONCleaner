import json
import pandas as pd
import os
from pathlib import Path
import traceback

# --- Configuration ---
MANDATORY_MAIN_FIELDS = {
    "imageUrl", "title", "totalScore", "reviewsCount", "street", "city",
    "state", "website", "phone", "categoryName", "url", "email" # Assuming email might be in JSON
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
    "title", "categoryName", "email", "totalScore", "reviewsCount",
    "street", "city", "state", "website", "phone", "imageUrl",
]

OUTPUT_SUBFOLDER = "XLXS_Consolidated" # Changed subfolder name
MASTER_EXCEL_FILENAME = "Consolidated_Business_Data.xlsx"
MAIN_SHEET_NAME = "MainData"
EXTRA_SHEET_NAME = "AdditionalInfo"
# --- End Configuration ---

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

def process_and_append_json(
    json_file_paths, # Now takes a list of paths
    mandatory_main_fields,
    unnecessary_fields,
    second_sheet_fields,
    desired_main_col_order,
    output_dir,
    master_excel_file_path,
):
    all_new_main_data_rows = []
    all_new_extra_data_rows = []

    for json_file_path in json_file_paths:
        print(f"\nProcessing {json_file_path.name}...")
        try:
            with open(json_file_path, "r", encoding="utf-8") as f:
                data_from_file = json.load(f)

            items_to_process = (
                [data_from_file] if isinstance(data_from_file, dict)
                else data_from_file if isinstance(data_from_file, list)
                else []
            )

            if not items_to_process:
                print(f"  Warning: Content of {json_file_path.name} is not processable. Skipping file.")
                continue

            for item_idx, item in enumerate(items_to_process):
                if not isinstance(item, dict):
                    print(f"  Warning: Non-dictionary item at index {item_idx} in {json_file_path.name}. Skipping item.")
                    continue

                main_data_obj = {}
                extra_data_obj = {}
                
                # Ensure email field is initialized if expected
                if "email" in mandatory_main_fields:
                    main_data_obj["email"] = item.get("email") # Get existing email if any

                for key, value in item.items():
                    if key == "email" and "email" in main_data_obj and main_data_obj["email"]: # Already handled
                        continue
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

                if main_data_obj: all_new_main_data_rows.append(main_data_obj)
                if extra_data_obj: all_new_extra_data_rows.append(extra_data_obj)
        
        except json.JSONDecodeError:
            print(f"  Error: Could not decode JSON from {json_file_path.name}. Skipping file.")
        except Exception as e:
            print(f"  An unexpected error occurred while processing {json_file_path.name}: {e}")
            traceback.print_exc()

    if not all_new_main_data_rows and not all_new_extra_data_rows:
        print("No new processable data found in any JSON file. Master Excel file not updated.")
        return

    df_new_main = pd.DataFrame(all_new_main_data_rows)
    df_new_extra = pd.DataFrame(all_new_extra_data_rows)

    # --- Read existing master file or create new DataFrames ---
    df_existing_main = pd.DataFrame()
    df_existing_extra = pd.DataFrame()

    if master_excel_file_path.is_file():
        print(f"\nReading existing master file: {master_excel_file_path}")
        try:
            xls = pd.ExcelFile(master_excel_file_path)
            if MAIN_SHEET_NAME in xls.sheet_names:
                df_existing_main = pd.read_excel(xls, sheet_name=MAIN_SHEET_NAME)
            if EXTRA_SHEET_NAME in xls.sheet_names:
                df_existing_extra = pd.read_excel(xls, sheet_name=EXTRA_SHEET_NAME)
        except Exception as e:
            print(f"  Warning: Could not read existing master Excel file. A new one will be created. Error: {e}")
    
    # --- Append new data ---
    df_combined_main = pd.concat([df_existing_main, df_new_main], ignore_index=True)
    df_combined_extra = pd.concat([df_existing_extra, df_new_extra], ignore_index=True)

    # --- Optional: Deduplication (example: based on 'title' and 'website' for main data) ---
    # You might want more sophisticated deduplication logic
    if not df_combined_main.empty and 'title' in df_combined_main.columns and 'website' in df_combined_main.columns:
        subset_cols = ['title', 'website'] # Define columns for identifying duplicates
         # Keep 'first' occurrence, you might want 'last' if new data should overwrite
        df_combined_main.drop_duplicates(subset=subset_cols, keep='first', inplace=True)
    
    # (Add deduplication for df_combined_extra if needed, based on relevant unique identifiers)


    # --- Reorder columns for MainData sheet ---
    if not df_combined_main.empty:
        current_main_cols = df_combined_main.columns.tolist()
        ordered_cols = []
        for col in desired_main_col_order:
            if col in current_main_cols:
                ordered_cols.append(col)
                if col in current_main_cols: current_main_cols.remove(col)
        
        url_col_present = "url" in current_main_cols
        if url_col_present: current_main_cols.remove("url")
        
        ordered_cols.extend(current_main_cols)
        
        if "url" in df_combined_main.columns:
            if "url" not in ordered_cols: ordered_cols.append("url")
            elif ordered_cols[-1] != "url":
                if "url" in ordered_cols: ordered_cols.remove("url")
                ordered_cols.append("url")
        
        df_combined_main = df_combined_main.reindex(columns=ordered_cols, fill_value=None)

    # --- Write to master Excel file ---
    try:
        with pd.ExcelWriter(master_excel_file_path, engine="openpyxl") as writer:
            if not df_combined_main.empty:
                df_combined_main.to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)
                adjust_column_widths(writer, MAIN_SHEET_NAME, df_combined_main)
            elif not df_combined_extra.empty: # Create empty main if extra has data
                 pd.DataFrame().to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)


            if not df_combined_extra.empty:
                df_combined_extra.to_excel(writer, sheet_name=EXTRA_SHEET_NAME, index=False)
                adjust_column_widths(writer, EXTRA_SHEET_NAME, df_combined_extra)
            elif not df_combined_main.empty: # Create empty extra if main has data
                 pd.DataFrame().to_excel(writer, sheet_name=EXTRA_SHEET_NAME, index=False)

        print(f"\nSuccessfully updated/created master Excel file: {master_excel_file_path}")
    except Exception as e:
        print(f"Error writing to master Excel file: {e}")
        traceback.print_exc()


def main_json_appender():
    output_directory_path = create_output_directory(OUTPUT_SUBFOLDER)
    master_excel_file = output_directory_path / MASTER_EXCEL_FILENAME
    print(f"Master Excel file will be: {master_excel_file.resolve()}")

    json_file_paths_input = input(
        "Enter paths to NEW JSON files to append, separated by commas (or a single path): "
    )
    if not json_file_paths_input:
        print("No file paths provided. Exiting.")
        return

    raw_paths = json_file_paths_input.split(",")
    json_file_paths_to_process = []
    for p_raw in raw_paths:
        p_stripped = p_raw.strip()
        p_cleaned = p_stripped[1:-1] if p_stripped.startswith('"') and p_stripped.endswith('"') else p_stripped
        if p_cleaned: json_file_paths_to_process.append(Path(p_cleaned))

    if not json_file_paths_to_process:
        print("No valid file paths extracted. Exiting.")
        return
    
    valid_json_files = []
    for file_path in json_file_paths_to_process:
        if not file_path.is_file():
            print(f"Warning: File not found at '{file_path}'. Skipping.")
            continue
        if file_path.suffix.lower() != ".json":
            print(f"Warning: File '{file_path}' is not a .json file. Skipping.")
            continue
        valid_json_files.append(file_path)

    if not valid_json_files:
        print("No valid JSON files to process. Exiting.")
        return

    process_and_append_json(
        valid_json_files,
        MANDATORY_MAIN_FIELDS,
        UNNECESSARY_FIELDS,
        SECOND_SHEET_FIELDS,
        DESIRED_MAIN_COLUMN_ORDER,
        output_directory_path, # Not strictly needed here but kept for consistency
        master_excel_file
    )

if __name__ == "__main__":
    # Decide which script to run or provide a way to choose
    # For now, I'll call the appender. You can change this.
    # main_excel_scraper() 
    main_json_appender()