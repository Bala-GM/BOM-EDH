import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Constants
TARGET_FOLDER = r'D:\MPNlibrary'
MPN_LIBRARY_FILE = os.path.join(TARGET_FOLDER, 'MPNLibrary.xlsx')
MISMATCH_FILE = os.path.join(TARGET_FOLDER, 'MissMatch_MPNLibrary.xlsx')
DUPLICATE_LOG = os.path.join(TARGET_FOLDER, 'Duplicate_Entries.txt')
REQUIRED_COLUMNS = ['MPN', 'TYPE', 'CID', 'PKG', 'SDJ']

# Ensure target folder exists
os.makedirs(TARGET_FOLDER, exist_ok=True)

def browse_and_append():
    # Hide the root Tkinter window
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if not file_path:
        print("No file selected.")
        return

    try:
        new_df = pd.read_excel(file_path)

        if not all(col in new_df.columns for col in REQUIRED_COLUMNS):
            print(f"ERROR: File must contain columns: {REQUIRED_COLUMNS}")
            return

        new_df = new_df[REQUIRED_COLUMNS].copy()
        new_df['UPD'] = pd.Timestamp.now()

        # Load existing MPNLibrary if exists
        if os.path.exists(MPN_LIBRARY_FILE):
            existing_df = pd.read_excel(MPN_LIBRARY_FILE)
        else:
            existing_df = pd.DataFrame(columns=REQUIRED_COLUMNS + ['UPD'])

        # Prepare final list to append
        append_list = []

        # Track duplicates and mismatches
        duplicate_entries = []
        mismatch_records = []

        for _, new_row in new_df.iterrows():
            mpn_matches = existing_df[existing_df['MPN'] == new_row['MPN']]

            if not mpn_matches.empty:
                exact_match = mpn_matches[
                    (mpn_matches['TYPE'] == new_row['TYPE']) &
                    (mpn_matches['CID'] == new_row['CID']) &
                    (mpn_matches['PKG'] == new_row['PKG']) &
                    (mpn_matches['SDJ'] == new_row['SDJ'])
                ]
                if not exact_match.empty:
                    # Duplicate found, skip and log
                    duplicate_entries.append(new_row.to_dict())
                    continue
                else:
                    # Mismatch found, log both old and new
                    for _, mismatch_row in mpn_matches.iterrows():
                        mismatch_records.append(mismatch_row)
                    mismatch_records.append(new_row)
                    continue

            # If MPN not found in existing data, it's a valid new entry
            append_list.append(new_row)

        # Append valid new rows to library
        if append_list:
            final_append_df = pd.DataFrame(append_list)
            updated_df = pd.concat([existing_df, final_append_df], ignore_index=True)
            updated_df.to_excel(MPN_LIBRARY_FILE, index=False)
            print(f"{len(append_list)} new record(s) appended to MPNLibrary.xlsx")
        else:
            print("No new records to append.")

        # Save mismatches
        if mismatch_records:
            mismatch_df = pd.DataFrame(mismatch_records)
            mismatch_df.to_excel(MISMATCH_FILE, index=False)
            print(f"{len(mismatch_records)} mismatch record(s) saved to MissMatch_MPNLibrary.xlsx")

        # Save duplicates to text file
        if duplicate_entries:
            with open(DUPLICATE_LOG, 'a') as f:
                for entry in duplicate_entries:
                    line = ', '.join([str(entry[col]) for col in REQUIRED_COLUMNS]) + '\n'
                    f.write(line)
            print(f"{len(duplicate_entries)} duplicate record(s) logged in Duplicate_Entries.txt")

    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    browse_and_append()
