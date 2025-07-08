#!/usr/bin/env python3
"""
Export renamed assets to semicolon-delimited CSV for Bynder.

Accepts spreadsheet and folder paths via command-line arguments.
If arguments are not provided, it prompts the user using Tkinter file dialogs.
Processes rows with non-empty values in columns A, B, and C.
Matches files named "Description-AD ID.ext" in asset folder.
Generates export CSV with exact columns and mappings, falling back to UI-provided
or spreadsheet values for certain fields if command-line arguments are not given.
Exports to user's Downloads folder with timestamped filename.
"""
import os
import sys
import argparse
from datetime import datetime
import pandas as pd
# Only import Tkinter if needed (when running standalone without arguments)
try:
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename, askdirectory
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("Warning: Tkinter not available. Script will only run via command-line arguments.", file=sys.stderr)


def parse_expiration(s):
    if pd.isna(s) or not isinstance(s, str):
        return ""
    s = s.strip()
    if "-" in s:
        parts = s.split("-", 1)
        # Ensure there's a second part before stripping
        if len(parts) > 1:
            return parts[1].strip()
    return s.split(" ")[0] if s else ""

def parse_year_from_value(value):
    if pd.isna(value) or value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        return s[:-2]
    return s

def main():
    parser = argparse.ArgumentParser(description="Export renamed assets to Bynder metadata CSV.")
    parser.add_argument("--spreadsheet", help="Path to the input spreadsheet (.xlsx, .xls, .csv).")
    parser.add_argument("--assets_folder", help="Path to the folder containing renamed video files.")
    parser.add_argument("--wrike_link", help="Optional: Link to Wrike Project (applies to all assets in batch).")
    parser.add_argument("--year", help="Optional: Year (applies to all assets in batch, e.g., 2024).")
    parser.add_argument("--sub_initiative", help="Optional: Sub-Initiative (applies to all assets in batch).")
    parser.add_argument("--location_type", help="Optional: Location Type (applies to all assets in batch).")

    args = parser.parse_args()

    file_path = args.spreadsheet
    folder = args.assets_folder

    # If running standalone or essential arguments are missing, use Tkinter dialogs
    if (not file_path or not folder) and TKINTER_AVAILABLE:
        print("Required arguments not provided. Falling back to Tkinter dialogs.")
        # Hide Tkinter root
        root = Tk()
        root.withdraw()

        if not file_path:
            file_path = askopenfilename(
                title="Select Spreadsheet",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
            )
            if not file_path:
                print("No spreadsheet selected. Exiting.")
                sys.exit(1)

        if not folder:
            folder = askdirectory(title="Select Folder of Assets")
            if not folder:
                print("No folder selected. Exiting.")
                sys.exit(1)
        root.destroy() # Destroy the Tkinter root after dialogs

    elif not file_path or not folder:
        print("Error: Missing required arguments. Please provide --spreadsheet and --assets_folder.")
        if not TKINTER_AVAILABLE:
            print("Tkinter is not installed, so graphical file selection is not possible.")
        sys.exit(1)

    # Read the spreadsheet into a DataFrame
    try:
        if file_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        else: # Assume CSV if not Excel
            df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error reading spreadsheet: {e}")
        sys.exit(1)

    # Filter valid rows (first three columns: A, B, C)
    if len(df.columns) < 3:
        print("Error: Spreadsheet must have at least 3 columns (A, B, C).")
        sys.exit(1)
    
    # Use column names '0', '1', '2' if headers are not present, or the actual column names
    # This assumes first three columns are what you mean by A, B, C
    required_cols = df.columns[:3].tolist()
    
    # Convert all required columns to string before dropping NA to ensure empty strings are handled
    for col in required_cols:
        df[col] = df[col].astype(str).replace('nan', '').str.strip()

    # Drop rows where any of the first three columns are empty after stripping
    valid_rows = df[df[required_cols].ne('').all(axis=1)]


    if valid_rows.empty:
        print("No valid rows found (need non-empty values in the first three columns). Exiting.")
        return

    # Prepare timestamped output in Downloads
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    os.makedirs(downloads_dir, exist_ok=True)
    save_path = os.path.join(downloads_dir, f"adVideo_metadataPrepped_{timestamp}.csv")

    # Define exact headers
    headers = [
        "filename", "name", "description", "Asset Type", "Asset Sub-Type", "Deliverable", "Product SKU",
        "Product SKU Position", "Asset Status", "Usage Rights", "tags", "File Type", "STEP Path",
        "Link to Wrike Project", "Sync to Site", "Generic Dimension Diagram With Measurements", "Admin Status",
        "Product Status", "Product Category", "Product Sub-Category", "Product Collection",
        "Component SKUs", "Stock Level (only relevant for Inline products)", "Restock Date (only relevant for Inline products)",
        "Link to Print Materials", "Link to Lifestyle Images", "Link to Store Images", "Initiative", "Sub-Initiative",
        "Print Tracking Code", "Print Tracking - Start Date", "Print Tracking - End Date", "Year", "Video Expiration",
        "Audio Licensing Expiration", "Ad ID", "Lead Offer Message", "Lead Finance Message", "Video Focus",
        "Video Objective", "Video Type", "Total Run Time (TRT)", "Spot Running (MM/DD/YYYY)", "Language",
        "Season", "Holiday/Special Occasion", "Talent", "Sunset Date (MM/DD/YYYY)", "Location Name", "Store Code",
        "Location Status", "Location Address", "Location Town", "Location State", "Location Zip Code",
        "Location Phone Number", "Location Type", "Location", "Inactive Product", "Partner", "Notes",
        "Sign Facade Color", "Sign Location", "Sign Color", "Sign Text",
        "Reviewed products in lifestyle", "Reviewed Studio Uploads", "Featured SKU", "Image Type",
        "scratchpad", "3D Model Source Files Acquired", "Visible to", "BynderTest", "dim_Length",
        "Bynder Report", "Dimensions", "dim_Height", "Figmage doc id", "dim_Width", "Figmage image extension",
        "Figmage node id", "Figmage page id", "Performance Metric", "DNUCampaign", "DNUFeatures",
        "DNUMaterials", "DNUStyle", "DNUPattern", "DNUPackage SKUs", "DNUSign Size",
        "DNUDistribution Channel", "Dim diagram re-cropped", "Embedded Instructions (for updating existing metadata based on automations)",
        "Mattress Size", "Asset Identifier", "Sync Batch", "Marked for Deletion from Site", "scene7 folder",
        "Variant Type", "Source", "PSA Image Type", "Rights Notes", "Workflow", "Workflow Status",
        "Product Name (STEP)", "Vendor Code", "Family Code", "Hero SKU", "Product Color", "Dropped",
        "Visible on Website", "Sales Channel", "Associated Materials Status", "Product in Studio",
        "DNU_PromoUpdate2", "Additional Files Upload Scratchpad", "Bump", "Carousel Dimensions Diagram Audit", "User Status", "Reviewed for Site Content Refresh", "Image Type Pre-Classification"
    ]

    allowed_exts = {"AI","CR2","CSS","DOC","DOCX","EPS","GIF","GLB","HTML","IDML","INDD","JFIF","JPEG","MOV","MP3","MP4","OTF","PDF","PNG","PPT", "WEBM", "AVI", "MKV"} # Added common video formats
    output_data = []
    
    total_rows = len(valid_rows)
    print(f"Total valid rows to process: {total_rows}")

    for i, (idx, row) in enumerate(valid_rows.iterrows()):
        desc = str(row.get("Description", "")).strip()
        ad_id = str(row.get("AD ID", "")).strip()
        
        # Print progress for UI
        print(f"PROGRESS:{i+1}/{total_rows}")
        
        # Skip if Description or AD ID is empty for a valid row (should be caught by dropna but good to double check)
        if not desc or not ad_id:
            print(f"Warning: Row {idx+2} skipped due to empty 'Description' or 'AD ID'.")
            continue

        new_stem = f"{desc}-{ad_id}"
        match_file = next((f for f in os.listdir(folder) if os.path.splitext(f)[0] == new_stem), None)
        if not match_file:
            print(f"No asset file found for '{new_stem}'. Skipping.")
            continue
        
        ext = os.path.splitext(match_file)[1].lstrip('.').upper()
        file_type = ext if ext in allowed_exts else ""
        spot = str(row.get("Spot Running", "")).strip()
        ad_name = str(row.get("Ad Name", "")).strip()

        entry = {h: "" for h in headers}
        entry["filename"] = match_file
        entry["File Type"] = file_type
        
        # Defaults
        entry["Asset Type"] = "Final Creative Materials"
        entry["Asset Sub-Type"] = "Ad Video"
        entry["Asset Status"] = "Final"
        entry["Usage Rights"] = "Approved for External Usage"
        entry["User Status"] = "Please review metadata"
        entry["Initiative"] = "Promos"
        
        # Determine Video Type based on Ad Name, with fallback to default
        # If the 'Ad Name' column exists and contains "(Animation)", set "Video Type" to "Animation"
        # Otherwise, default to "Live Action".
        if "Ad Name" in row and isinstance(row["Ad Name"], str) and "(Animation)" in row["Ad Name"]:
            entry["Video Type"] = "Animation"
        else:
            entry["Video Type"] = "Live Action" # CHECK ON THIS - Confirmed: Default to Live Action.

        # --- CLI arguments / UI inputs take precedence, then fallback to spreadsheet ---
        
        # Link to Wrike Project
        if args.wrike_link:
            entry["Link to Wrike Project"] = args.wrike_link
        else:
            entry["Link to Wrike Project"] = str(row.get("Link to Wrike Project", "")).strip()

        # Year
        if args.year:
            entry["Year"] = parse_year_from_value(args.year)
        else:
            entry["Year"] = parse_year_from_value(row.get("Year", ""))

        # Sub-Initiative
        if args.sub_initiative:
            entry["Sub-Initiative"] = args.sub_initiative
        else:
            entry["Sub-Initiative"] = str(row.get("Sub-Initiative", "")).strip()
        
        # Location Type
        if args.location_type:
            entry["Location Type"] = args.location_type
        else:
            entry["Location Type"] = str(row.get("Location Type", "")).strip()

        # Matrix mappings (always from spreadsheet row)
        entry["Deliverable"] = str(row.get("Placement(s)", "")).strip()
        entry["Spot Running (MM/DD/YYYY)"] = spot
        entry["Video Expiration"] = parse_expiration(spot)
        entry["Ad ID"] = ad_id
        entry["Lead Offer Message"] = str(row.get("Lead Offer Message", "")).strip()
        entry["Lead Finance Message"] = str(row.get("Lead Finance Message", "")).strip()
        
        # Video Focus - derive from Ad Name
        video_focus = ad_name.split(" (Spanish)")[0].strip()
        # If video_focus is just "____", clear it or handle as needed. Assuming it implies empty.
        entry["Video Focus"] = video_focus if video_focus != "____" else ""


        entry["Video Objective"] = str(row.get("Objective", "")).strip()
        
        # Total Run Time (TRT) conversion
        trt_val = str(row.get("TRT", "")).strip()
        if trt_val == ":06":
            entry["Total Run Time (TRT)"] = "06 seconds"
        elif trt_val == ":15":
            entry["Total Run Time (TRT)"] = "15 seconds"
        elif trt_val == ":30":
            entry["Total Run Time (TRT)"] = "30 seconds"
        else:
            entry["Total Run Time (TRT)"] = trt_val # Keep as is if no match

        # Language
        entry["Language"] = "Spanish" if "Spanish" in ad_name else "English"

        output_data.append(entry)

    # Build DataFrame and export
    out_df = pd.DataFrame(output_data, columns=headers)
    out_df.to_csv(save_path, sep=';', index=False)
    print(f"Your Bynder metadata import CSV has been saved to {save_path}.")
    print("Script finished.")


if __name__ == "__main__":
    main()