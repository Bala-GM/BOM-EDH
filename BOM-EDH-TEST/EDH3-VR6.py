import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog

# PY VER-1.0.1 APR|03|04|2025

def browse_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    return file_path

def update_excel_file(file_path):
    if not file_path:
        print("No file selected.")
        return
    
    # Load the browsed Excel file
    try:
        df_browsed = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        print(f"Error loading browsed file: {e}")
        return
    
    required_headers = ['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'SDJ']
    if not all(header in df_browsed.columns for header in required_headers):
        print("The browsed file does not contain the required headers.")
        return
    
    # Load MPNLibrary.xlsx
    mpn_library_path = r'D:\MPNlibrary\MPNLibrary.xlsx'
    if not os.path.exists(mpn_library_path):
        print("MPNLibrary.xlsx not found in D:\\MPNlibrary.")
        return
    
    try:
        df_library = pd.read_excel(mpn_library_path, dtype=str)
    except Exception as e:
        print(f"Error loading MPNLibrary.xlsx: {e}")
        return
    
    required_lib_headers = ['MPN', 'TYPE', 'PKG', 'SDJ']
    if not all(header in df_library.columns for header in required_lib_headers):
        print("MPNLibrary.xlsx does not contain the required headers.")
        return
    
    # Merge data based on 'MPN'
    df_updated = df_browsed.merge(df_library, on='MPN', how='left', suffixes=('', '_new'))
    
    # Replace 'TYPE', 'PKG', 'SDJ' if MPN matches
    for column in ['TYPE', 'PKG', 'SDJ']:
        df_updated[column] = df_updated[column + '_new'].combine_first(df_updated[column])
        df_updated.drop(columns=[column + '_new'], inplace=True)
    
    # Save back to the original file
    df_updated.to_excel(file_path, index=False)
    print(f"File updated and saved: {file_path}")

if __name__ == "__main__":
    file_to_update = browse_file()
    update_excel_file(file_to_update)
