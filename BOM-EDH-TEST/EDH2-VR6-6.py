import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Constants
TARGET_FOLDER = r'D:\MPNlibrary'
MPN_LIBRARY_FILE = os.path.join(TARGET_FOLDER, 'MPNLibrary.xlsx')
REQUIRED_COLUMNS = ['MPN', 'TYPE', 'CID', 'PKG', 'SDJ']

# Ensure target folder exists
os.makedirs(TARGET_FOLDER, exist_ok=True)

def browse_and_append():
    # Suppress the main Tkinter root window
    root = tk.Tk()
    root.withdraw()

    # Open file browser
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if not file_path:
        print("No file selected.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Check if required columns exist
        if not all(col in df.columns for col in REQUIRED_COLUMNS):
            print(f"ERROR: File must contain the columns: {REQUIRED_COLUMNS}")
            return

        # Filter required columns and add timestamp
        df_filtered = df[REQUIRED_COLUMNS].copy()
        df_filtered['UPD'] = pd.Timestamp.now()

        # Append or create Excel file
        if os.path.exists(MPN_LIBRARY_FILE):
            existing_df = pd.read_excel(MPN_LIBRARY_FILE)
            combined_df = pd.concat([df_filtered, existing_df], ignore_index=True)
        else:
            combined_df = df_filtered

        # Save to Excel
        combined_df.to_excel(MPN_LIBRARY_FILE, index=False)
        print(f"SUCCESS: Data appended to {MPN_LIBRARY_FILE}")

    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    browse_and_append()
