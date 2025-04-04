import os
import sys
import pandas as pd
from tkinter import Tk, filedialog

# PY VER-1.0.1 APR|03|04|2025

# Define target folder
TARGET_FOLDER = r'D:\MPNlibrary'
MPN_LIBRARY_FILE = os.path.join(TARGET_FOLDER, 'MPNLibrary.xlsx')
DUPLICATE_FILE = os.path.join(TARGET_FOLDER, 'Duplicate_MPNLibrary.xlsx')
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Ensure target folder exists
os.makedirs(TARGET_FOLDER, exist_ok=True)

# Function to check file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Open file dialog
root = Tk()
root.withdraw()  # Hide the root window
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])

if file_path:
    if allowed_file(file_path):
        try:
            # Read new Excel file
            df_new = pd.read_excel(file_path, usecols=['MPN', 'TYPE', 'PKG', 'SDJ'])
            
            # Load existing MPNLibrary if it exists
            if os.path.exists(MPN_LIBRARY_FILE):
                df_existing = pd.read_excel(MPN_LIBRARY_FILE)
                
                # Find duplicates
                df_duplicates = df_existing.merge(df_new, on=['MPN', 'TYPE', 'PKG', 'SDJ'], how='inner')
                if not df_duplicates.empty:
                    print("Duplicate data found. Saving to Duplicate_MPNLibrary.xlsx")
                    df_duplicates.to_excel(DUPLICATE_FILE, index=False)
                
                # Append new data
                df_combined = pd.concat([df_existing, df_new]).drop_duplicates()
            else:
                df_combined = df_new
            
            # Save updated MPNLibrary
            df_combined.to_excel(MPN_LIBRARY_FILE, index=False)
            print(f'File saved to {MPN_LIBRARY_FILE}')
        except Exception as e:
            print(f'Error processing file: {str(e)}')
    else:
        print('Invalid file format. Only .xls and .xlsx allowed.')
else:
    print('No file selected.')
    
#sys.exit()
