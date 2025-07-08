#!/usr/bin/env python3
"""
Rename video files based on spreadsheet columns.

Accepts spreadsheet and folder paths via command-line arguments.
If arguments are not provided, it prompts the user using Tkinter file dialogs.
Processes rows with non-empty values in columns A, B, and C.
Renames files whose name matches the 'Description' column by appending '-AD ID' to the name, retaining the extension.
At the end, prints a summary of renamed files, missing matches, and unmatched folder files.
"""
import os
import sys
import argparse
import pandas as pd
# Only import Tkinter if needed (when running standalone without arguments)
try:
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename, askdirectory
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("Warning: Tkinter not available. Script will only run via command-line arguments.", file=sys.stderr)

def main():
    parser = argparse.ArgumentParser(description="Rename video files based on spreadsheet columns.")
    parser.add_argument("--spreadsheet", help="Path to the input spreadsheet (.xlsx, .xls, .csv).")
    parser.add_argument("--video_folder", help="Path to the folder containing video files to be renamed.")

    args = parser.parse_args()

    file_path = args.spreadsheet
    folder = args.video_folder

    # If running standalone or essential arguments are missing, use Tkinter dialogs
    if (not file_path or not folder) and TKINTER_AVAILABLE:
        print("Required arguments not provided. Falling back to Tkinter dialogs.")
        root = Tk()
        root.withdraw()

        if not file_path:
            file_path = askopenfilename(
                title="Select Spreadsheet",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
            )
            if not file_path:
                print("No file selected. Exiting.")
                sys.exit(1)

        if not folder:
            folder = askdirectory(title="Select Folder of Video Files")
            if not folder:
                print("No folder selected. Exiting.")
                sys.exit(1)
        root.destroy() # Destroy the Tkinter root after dialogs
    
    elif not file_path or not folder:
        print("Error: Missing required arguments. Please provide --spreadsheet and --video_folder.")
        if not TKINTER_AVAILABLE:
            print("Tkinter is not installed, so graphical file selection is not possible.")
        sys.exit(1)


    # Read the spreadsheet into a DataFrame
    try:
        if file_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error reading spreadsheet: {e}")
        sys.exit(1)

    # Determine required columns (first three columns: A, B, C)
    if len(df.columns) < 3:
        print("Error: Spreadsheet must have at least 3 columns (A, B, C).")
        sys.exit(1)
    
    required_cols = df.columns[:3].tolist()
    
    # Convert all required columns to string before dropping NA to ensure empty strings are handled
    for col in required_cols:
        df[col] = df[col].astype(str).replace('nan', '').str.strip()

    # Filter valid rows: non-null in all three required columns
    valid_rows = df[df[required_cols].ne('').all(axis=1)]

    if valid_rows.empty:
        print("No valid rows found (need non-empty in columns A, B, and C). Exiting.")
        return

    # Initialize lists to store detailed information
    successful_renames = 0
    could_not_rename_files = [] # To store files that couldn't be renamed due to no match
    files_in_folder_without_spreadsheet_match = [] # To store files in the folder not matched by any spreadsheet entry

    # Get initial files in the folder to compare against later
    initial_files_in_folder_full_paths = {os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))}
    initial_files_in_folder_stems = {os.path.splitext(f)[0] for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))}

    # Keep track of the descriptions from the spreadsheet that successfully led to a rename
    matched_spreadsheet_descriptions = set()

    total_rows = len(valid_rows)
    print(f"Total valid rows to process: {total_rows}")

    # Process each valid row
    for i, (idx, row) in enumerate(valid_rows.iterrows()):
        desc = str(row.get("Description", "")).strip()
        ad_id = str(row.get("AD ID", "")).strip()
        
        # Print progress for UI
        print(f"PROGRESS:{i+1}/{total_rows}")

        # Skip if Description or AD ID is empty (redundant with valid_rows filter, but safe)
        if not desc or not ad_id:
            # This case should ideally not be hit due to `valid_rows` filter, but good for robustness.
            print(f"Row {idx+2}: Missing Description or AD ID. Skipping.")
            continue # Don't count this as a file that couldn't be renamed from the folder's perspective

        new_stem = f"{desc}-{ad_id}"
        found_match_in_folder = False
        original_filename_matched = None # To store the filename that was found and potentially renamed

        # Search for a matching file in the folder based on 'Description'
        for filename_in_folder in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, filename_in_folder)):
                stem_in_folder, ext_in_folder = os.path.splitext(filename_in_folder)
                
                if stem_in_folder == desc:
                    src = os.path.join(folder, filename_in_folder)
                    dst = os.path.join(folder, new_stem + ext_in_folder)
                    original_filename_matched = filename_in_folder # Store the original name
                    
                    if src != dst: # Only rename if the name actually changes
                        try:
                            os.rename(src, dst)
                            print(f"Renamed '{filename_in_folder}' to '{new_stem + ext_in_folder}'")
                            successful_renames += 1
                            found_match_in_folder = True
                            matched_spreadsheet_descriptions.add(desc) # Mark this description as matched
                            break # Move to next spreadsheet row
                        except Exception as e:
                            print(f"Error renaming '{filename_in_folder}': {e}")
                            could_not_rename_files.append(filename_in_folder) # Add to list of failed renames
                            found_match_in_folder = True # Treat as "attempted to match" to not double count later
                            break
                    else:
                        print(f"File '{filename_in_folder}' already has the target name. Skipping rename.")
                        successful_renames += 1 # Count as success if already correctly named
                        found_match_in_folder = True
                        matched_spreadsheet_descriptions.add(desc)
                        break # Move to next spreadsheet row

        if not found_match_in_folder:
            # This means the 'Description' from the spreadsheet didn't find any file in the folder
            could_not_rename_files.append(f"No file found in folder for Description: '{desc}' from spreadsheet")

    # --- Final Summary Output ---
    
    # Calculate files in the folder that were not matched by any 'Description' value from the spreadsheet
    # We iterate through the initial files in the folder and check if their stem was ever a 'Description'
    # that led to a successful match.
    
    # Get the current list of files in the folder after all renames
    current_files_in_folder_stems = {os.path.splitext(f)[0] for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))}

    # Files that were in the folder at the start and whose stem *was not* a 'Description' that led to a match
    # or whose new name doesn't match any target new_stem from the spreadsheet
    
    # The most robust way is to compare the *initial* set of files in the folder against all 'Description' values
    # in the spreadsheet (regardless of whether they were successfully renamed or not, just if a match was attempted).
    
    # Collect all 'Description' values from the spreadsheet
    all_spreadsheet_descriptions = {str(row.get("Description", "")).strip() for _, row in valid_rows.iterrows() if str(row.get("Description", "")).strip()}

    for filename_in_folder in initial_files_in_folder_full_paths:
        stem_in_folder, _ = os.path.splitext(os.path.basename(filename_in_folder))
        if stem_in_folder not in all_spreadsheet_descriptions and \
           f"{stem_in_folder}" not in matched_spreadsheet_descriptions: # Check if the initial stem or its new stem was successfully renamed
            
            # We need to be careful here: if a file was renamed from "old_name.mp4" to "new_name-ADID.mp4",
            # then "old_name" was a match. We only want to list files whose *original* name (stem)
            # was never found in the spreadsheet's 'Description' column, *and* wasn't successfully renamed.
            
            # This means if a file "video.mp4" was in the folder, and "video" was in the spreadsheet
            # and got renamed, it shouldn't be counted as "unmatched".
            # The current `matched_spreadsheet_descriptions` only tracks the `desc` that *led to* a rename.

            # A simpler approach for "files in your folder that don't have a match in your spreadsheet":
            # Compare the initial list of file stems in the folder with the set of all 'Description' values
            # present in the valid rows of the spreadsheet.
            if stem_in_folder not in all_spreadsheet_descriptions:
                files_in_folder_without_spreadsheet_match.append(os.path.basename(filename_in_folder))


    print("\n--- Script Summary ---")
    print(f"{successful_renames} video files were successfully renamed.")

    if could_not_rename_files:
        print(f"\n{len(could_not_rename_files)} video files couldn't be renamed because there was no file found matching Description value in spreadsheet or an error occurred during renaming:")
        for file_info in could_not_rename_files:
            print(f"- {file_info}")
    else:
        print("\n0 video files couldn't be renamed due to missing matches or errors.")

    if files_in_folder_without_spreadsheet_match:
        # Remove duplicates if any (e.g., if a file was moved/deleted mid-script, though unlikely with os.rename)
        files_in_folder_without_spreadsheet_match = sorted(list(set(files_in_folder_without_spreadsheet_match)))
        print(f"\nThere were {len(files_in_folder_without_spreadsheet_match)} files in your folder that don't have a match in your spreadsheet:")
        for file_name in files_in_folder_without_spreadsheet_match:
            print(f"- {file_name}")
    else:
        print("\nAll files in your folder had a corresponding match in your spreadsheet.")

    print("\nScript finished.")


if __name__ == "__main__":
    main()
