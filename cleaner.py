import json
import pandas as pd
import os
from pathlib import Path
import traceback # Added for more detailed error logging if needed

# --- Configuration ---

# Fields that SHOULD be in the "MainData" sheet.
# Their presence is prioritized. Order will be managed separately.
MANDATORY_MAIN_FIELDS = {
    "imageUrl",
    "title",
    "totalScore",
    "reviewsCount",
    "street",
    "city",
    "state",
    "website",
    "phone",
    "categoryName",
    "url",
    # "email" # Add 'email' here if it's mandatory and present in your JSON
}

# Fields to be moved to the "AdditionalInfo" sheet.
# These will be REMOVED from the "MainData" sheet UNLESS they are also in MANDATORY_MAIN_FIELDS.
SECOND_SHEET_FIELDS = {
    "claimThisBusiness",
    "permanentlyClosed",
    "temporarilyClosed",
    "openingHours",
    "additionalInfo",
    "countryCode",
}

# Define the fields (keys) you want to REMOVE COMPLETELY from the data,
# UNLESS they are in MANDATORY_MAIN_FIELDS.
UNNECESSARY_FIELDS = {
    "price",
    "neighborhood",
    "imageCategories",
    "scrapedAt",
    "googleFoodUrl",
    "hotelAds",
    "gasPrices",
    "searchPageUrl",
    "searchString",
    "language",
    "placeId",
    "cid",
    "fid",
    "kgmid",
    "imagesCount",
    "rank",
    "isAdvertisement",
    "phoneUnformatted",
    "reviewsDistribution",
    "peopleAlsoSearch",
    "placesTags",
    "reviewsTags",
}

# Desired order for columns in the MainData sheet.
# 'url' will be handled to be last. Other non-specified columns will be appended before 'url'.
DESIRED_MAIN_COLUMN_ORDER = [
    "title",
    "categoryName",
    "totalScore",
    "reviewsCount",
    "street",
    "city",
    "state",
    "website",
    "phone",
    "imageUrl",
    # 'email', # If you have an email field and want it here
]


OUTPUT_SUBFOLDER = "XLXS"
MAIN_SHEET_NAME = "MainData"
EXTRA_SHEET_NAME = "AdditionalInfo"
# --- End Configuration ---

def create_output_directory(folder_name):
    path = Path.cwd() / folder_name
    path.mkdir(parents=True, exist_ok=True)
    return path

def adjust_column_widths(writer, sheet_name, df):
    if df.empty:
        return
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df):
        series = df[col]
        max_len = (
            max(
                (
                    series.astype(str).map(len).max(),
                    len(str(series.name)),
                )
            )
            + 2
        )
        max_len = min(max_len, 70) # Cap max width
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=idx + 1).column_letter
        ].width = max_len


def process_json_file(
    json_file_path,
    mandatory_main_fields,
    unnecessary_fields,
    second_sheet_fields,
    desired_main_col_order,
    output_dir,
):
    all_main_data_rows = []
    all_extra_data_rows = []

    try:
        with open(json_file_path, "r", encoding="utf-8") as f:
            data_from_file = json.load(f)

        items_to_process = (
            [data_from_file]
            if isinstance(data_from_file, dict)
            else data_from_file
            if isinstance(data_from_file, list)
            else []
        )

        if not items_to_process:
            print(
                f"Warning: Content of {json_file_path.name} is not a processable JSON object or list. Skipping."
            )
            return

        for item_idx, item in enumerate(items_to_process):
            if not isinstance(item, dict):
                print(
                    f"Warning: Found non-dictionary item at index {item_idx} in {json_file_path.name}. Skipping item."
                )
                continue

            main_data_obj = {}
            extra_data_obj = {}

            for key, value in item.items():
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

            if main_data_obj:
                all_main_data_rows.append(main_data_obj)
            if extra_data_obj:
                all_extra_data_rows.append(extra_data_obj)

        if not all_main_data_rows and not all_extra_data_rows:
            print(f"No processable data found in {json_file_path.name}. No Excel file created.")
            return

        df_main = pd.DataFrame(all_main_data_rows)
        df_extra = pd.DataFrame(all_extra_data_rows)

        if not df_main.empty:
            current_main_cols = df_main.columns.tolist()
            ordered_cols = []
            
            for col in desired_main_col_order:
                if col in current_main_cols:
                    ordered_cols.append(col)
                    current_main_cols.remove(col)
            
            url_col_present = "url" in current_main_cols
            if url_col_present:
                current_main_cols.remove("url")
            
            ordered_cols.extend(current_main_cols)
            
            if "url" in df_main.columns: # Check if 'url' actually exists in the DataFrame
                if "url" not in ordered_cols: # If it wasn't part of desired_main_col_order
                    ordered_cols.append("url")
                elif ordered_cols[-1] != "url": # If it was in desired but not last, move it
                    ordered_cols.remove("url")
                    ordered_cols.append("url")

            df_main = df_main.reindex(columns=ordered_cols, fill_value=None)


        output_excel_filename = f"{json_file_path.stem}.xlsx"
        output_excel_path = output_dir / output_excel_filename

        with pd.ExcelWriter(
            output_excel_path, engine="openpyxl"
        ) as writer:
            if not df_main.empty:
                df_main.to_excel(
                    writer, sheet_name=MAIN_SHEET_NAME, index=False
                )
                adjust_column_widths(writer, MAIN_SHEET_NAME, df_main)
            elif not df_extra.empty:
                 pd.DataFrame().to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)

            if not df_extra.empty:
                df_extra.to_excel(
                    writer, sheet_name=EXTRA_SHEET_NAME, index=False
                )
                adjust_column_widths(writer, EXTRA_SHEET_NAME, df_extra)
            elif not df_main.empty:
                 pd.DataFrame().to_excel(writer, sheet_name=EXTRA_SHEET_NAME, index=False)

        print(f"Successfully created Excel file: {output_excel_path}")

    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from {json_file_path.name}. Skipping.")
    except Exception as e:
        print(
            f"An unexpected error occurred while processing {json_file_path.name}: {e}"
        )
        traceback.print_exc()

def main():
    output_directory_path = create_output_directory(OUTPUT_SUBFOLDER)
    print(f"Output will be saved to: {output_directory_path.resolve()}")

    json_file_paths_input = input(
        "Enter the paths to your JSON files, separated by commas (or a single path): "
    )
    if not json_file_paths_input:
        print("No file paths provided. Exiting.")
        return

    # Split the input string by commas
    raw_paths = json_file_paths_input.split(",")
    
    json_file_paths = []
    for p_raw in raw_paths:
        # Strip leading/trailing whitespace from the segment
        p_stripped = p_raw.strip()
        # If the stripped segment starts and ends with a double quote, remove them
        if p_stripped.startswith('"') and p_stripped.endswith('"'):
            p_cleaned = p_stripped[1:-1]
        else:
            p_cleaned = p_stripped
        
        if p_cleaned: # Ensure it's not an empty string after stripping
            json_file_paths.append(Path(p_cleaned))

    if not json_file_paths:
        print("No valid file paths were extracted from the input. Exiting.")
        return

    for file_path in json_file_paths:
        if not file_path.is_file():
            print(f"Warning: File not found at '{file_path}'. Skipping.")
            continue
        if file_path.suffix.lower() != ".json":
            print(
                f"Warning: File '{file_path}' is not a .json file. Skipping."
            )
            continue

        print(f"\nProcessing {file_path.name}...")
        process_json_file(
            file_path,
            MANDATORY_MAIN_FIELDS,
            UNNECESSARY_FIELDS,
            SECOND_SHEET_FIELDS,
            DESIRED_MAIN_COLUMN_ORDER,
            output_directory_path,
        )

if __name__ == "__main__":
    main()