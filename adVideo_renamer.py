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

    # Counters for summary
    successful_renames = 0
    missing_matches_count = 0
    
    # Keep track of the original stems of files in the folder and their new stems if renamed
    # This helps accurately calculate unmatched files later
    initial_files_in_folder_stems = {os.path.splitext(f)[0] for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))}
    
    # Use a set to store stems of files that *were successfully processed* (renamed or already correct)
    processed_stems_from_spreadsheet = set()

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
            print(f"Row {idx+2}: Missing Description or AD ID. Skipping.")
            missing_matches_count += 1
            continue

        new_stem = f"{desc}-{ad_id}"
        found_match_in_folder = False

        # Search for a matching file in the folder based on 'Description'
        for filename_in_folder in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, filename_in_folder)):
                stem_in_folder, ext_in_folder = os.path.splitext(filename_in_folder)
                
                if stem_in_folder == desc:
                    src = os.path.join(folder, filename_in_folder)
                    dst = os.path.join(folder, new_stem + ext_in_folder)
                    
                    if src != dst: # Only rename if the name actually changes
                        try:
                            os.rename(src, dst)
                            print(f"Renamed '{filename_in_folder}' to '{new_stem + ext_in_folder}'")
                            successful_renames += 1
                            processed_stems_from_spreadsheet.add(new_stem) # Mark the new stem as processed
                            found_match_in_folder = True
                            break # Move to next spreadsheet row
                        except Exception as e:
                            print(f"Error renaming '{filename_in_folder}': {e}")
                            missing_matches_count += 1 # Count as a failed rename/match attempt
                            found_match_in_folder = True # Treat as "attempted to match" to not double count later
                            break
                    else:
                        print(f"File '{filename_in_folder}' already has the target name. Skipping rename.")
                        successful_renames += 1 # Count as success if already correctly named
                        processed_stems_from_spreadsheet.add(new_stem)
                        found_match_in_folder = True
                        break # Move to next spreadsheet row

        if not found_match_in_folder:
            print(f"No file found for Description '{desc}'.")
            missing_matches_count += 1

    # --- Original Summary Output Format ---
    total_files_in_folder_at_start = len(initial_files_in_folder_stems)
    
    # Calculate files that were in the folder but were not matched by any 'Description' value
    # This is tricky because a file might have been renamed.
    # The most accurate way to get "unmatched files in folder" is to list files *after* renaming
    # and compare them to the set of target names (`new_stem`) that were successfully created.

    final_files_in_folder_stems = {os.path.splitext(f)[0] for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))}

    # Unmatched files in the folder = files whose current stem is not in the set of successful target stems
    unmatched_folder_files_count = len(final_files_in_folder_stems) - successful_renames
    # Note: this counts files that were already named correctly as "matched".

    print(f"{successful_renames} video files were successfully renamed. ")
    print(f"{missing_matches_count} video files couldn't be renamed because there was no file found matching Description value in spreadsheet. ")
    print(f"There were {unmatched_folder_files_count} files in your folder that don't have a match in your spreadsheet. See activity log for details.")
    print("Script finished.")


if __name__ == "__main__":
    main()