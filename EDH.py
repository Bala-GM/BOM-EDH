import os
import sys
import pandas as pd
import re
import tkinter as tk
from tkinter import Tk, filedialog
from tkinter import filedialog, messagebox
from datetime import datetime

# PY VER-1.0.1 APR|03|04|2025

# ============================ Program 1: BOM EDH Analyser ============================
def pro_1():
    
    print("\nRunning BOM EDH Analyser (Program 1)...")

    print("\n")
    
    def browse_file():
        root = tk.Tk()
        root.withdraw()
        return filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def validate_data(df, required_columns):
        missing_data = [f"Missing column: {col}" for col in required_columns if col not in df.columns]
        for col in required_columns:
            if col in df.columns:
                missing_data.extend([f"Row {i + 2}, Column '{col}' is missing." for i, val in df[col].items() if pd.isna(val) or val == ""])
        return missing_data

    def save_error_log(file_path, errors):
        error_file = os.path.join(os.path.dirname(file_path), "error_log.txt")
        with open(error_file, "w") as f:
            f.write("\n".join(errors))
        messagebox.showinfo("Error Log Saved", f"Errors saved in {error_file}")

    def save_cleaned_data(df, file_path):
        save_name = input("Enter the name for the cleaned file: ")
        save_path = os.path.join(os.path.dirname(file_path), f"{save_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df.to_excel(save_path, index=False, engine="openpyxl")
        messagebox.showinfo("File Saved", f"Cleaned data saved as {save_path}")

    def extract_shape(description):
        desired_shapes = {"0201", "0402", "0603", "0805", "1206"}
        match = re.search(r'\b\d{4}\b', str(description))
        return match.group() if match and match.group() in desired_shapes else None

    def extract_shapediode(description):
        """Extracts diode package shape from the DES column."""
        desired_shapes = {"SMA", "SOD-123", "SOT-23", "SOT23", "DPAK", "D2PAK", "TO-263AB", "TVS", "DO-214AC", "SOD-80"}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None

    def extract_shapeic(description):
        """Extracts diode package shape from the DES column."""
        desired_shapes = {"SOP8", "8SOP", "SOP16"}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None

    def extract_shapetht(description):
        """Extracts diode package shape from the DES column."""
        desired_shapes = {"THT", "tht", "Dip", "DIP", "dip", "SDIP", "sdip"}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None

    def process_excel():
        file_path = browse_file()
        if not file_path:
            messagebox.showerror("Error", "No file selected")
            return
        
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in [".xlsx", ".xls"]:
            messagebox.showerror("Error", "Selected file is not an Excel file.")
            return

        try:
            df = pd.read_excel(file_path, dtype=str, engine="openpyxl" if file_ext == ".xlsx" else "xlrd")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the file:\n{e}")
            return

        required_columns = ["CRD", "DES", "MUF", "MPN", "QTY"]
        missing_data = validate_data(df, required_columns)
        if missing_data:
            user_choice = messagebox.askyesnocancel("Data Issues Found", "\n".join(missing_data) + "\n\nReplace missing values with NaN and continue?")
            if user_choice is None:
                return
            elif user_choice:
                for col in required_columns:
                    if col in df.columns:
                        df[col] = df[col].fillna("NaN")
            else:
                save_error_log(file_path, missing_data)
                messagebox.showerror("Process Aborted", "Process was aborted due to missing data.")
                return

        df["SHP"] = df["DES"].apply(extract_shape)
        df["DID"] = df["DES"].apply(extract_shapediode)
        df["SOIC"] = df["DES"].apply(extract_shapeic)
        df["THT"] = df["DES"].apply(extract_shapetht)

        # Extract component types
        def extract_component(des, keywords):
            for keyword in keywords:
                if keyword.lower() in str(des).lower():
                    return keyword
            return None

        df['TYPE'] = df['DES'].apply(lambda x: extract_component(x, ["CAP", "RES", "IND"]))
        df['SPERESCOMP'] = df['DES'].apply(lambda x: "MELF" if "MELF" in str(x).upper() else "")
        df['SPETHTCOMP'] = df['DES'].apply(lambda x: "THT" if any(k in str(x).lower() for k in ["THT", "dip"]) else "")
        df['SPESODCOMP'] = df['DES'].apply(lambda x: extract_component(x, ["ZENER", "DIODE", "SOD"]))
        df['SPECAPCOMP'] = df['DES'].apply(lambda x: extract_component(x, ["TAN", "Tantalum", "Aluminium", "ALLUM", "ALUM", "Electrolytic", "ALU"]))
        df['SPEFERINDCOMP'] = df['DES'].apply(lambda x: extract_component(x, ["IND", "FERRITEBEAD", "FERRITE", "BEAD", "INDUCTOR"]))
        
        df['TYPE'] = df[['TYPE', 'SPERESCOMP', 'SPETHTCOMP', 'SPESODCOMP', 'SPECAPCOMP', 'SPEFERINDCOMP']].apply(lambda x: ''.join(x.dropna().astype(str)), axis=1)
        df['TYPE'] = df['TYPE'].replace({"FERRITEBEAD": "IND", "BEAD": "IND", "FERRITE": "IND", "INDIND": "IND"})
        df["SHP"] = df[['TYPE', 'SHP']].apply(lambda x: '-'.join(x.dropna().astype(str)), axis=1)

        # Remove intermediate columns
        df.drop(columns=['SPERESCOMP', 'SPETHTCOMP', 'SPESODCOMP', 'SPECAPCOMP', 'SPEFERINDCOMP'], inplace=True)
        
        # Rename SHP to PKG
        df.rename(columns={'SHP': 'PKG'}, inplace=True)

        # Standardize TYPE names
        df['TYPE'] = df['TYPE'].replace({"CAP": "CAPACITOR", "RES": "RESISTOR", "IND": "INDUCTOR"})

        # Standardize PKG names
        replacements = {"CAP-0201": "C0201", "CAP-0402": "C0402", "CAP-0603": "C0603", "CAP-0805": "C0805", "CAP-1206": "C1206", 
                        "RES-0201": "R0201", "RES-0402": "R0402", "RES-0603": "R0603", "RES-0805": "R0805", "RES-1206": "R1206", 
                        "IND-0201": "C0201", "IND-0402": "C0402", "IND-0603": "C0603", "IND-0805": "C0805", "IND-1206": "C1206",
                        "DIODE": "", "MELFDIODE": "", "ZENER": "", "TAN": ""} 
        df['PKG'] = df['PKG'].replace(replacements)
        
        df["PKG"] = df[['PKG', 'DID', 'SOIC', 'THT']].apply(lambda x: ''.join(x.dropna().astype(str)), axis=1)

        # Reorder columns
        df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG']]
        #df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'DID', 'SOIC', 'THT']]
        #df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'DID', 'SPERESCOMP', 'SPETHTCOMP', 'SPESODCOMP', 'SPECAPCOMP', 'SPEFERINDCOMP']]
        
        # Define the set of values to check for SDJ column
        sdj_values = {"C0201", "C0402", "C0603", "C0805", "C1206", "R0201", "R0402", "R0603", "R0805", "R1206"}

        # Initialize SDJ column with default empty value
        df["SDJ"] = ""

        # Assign SDJ = 2 for specific PKG values
        df.loc[df["PKG"].isin(sdj_values), "SDJ"] = 2

        # Define the set of values to check for diode packages
        sdj_valuesdid = {"SMA", "SOD-123", "SOT-23", "SOT23", "DPAK", "D2PAK", "TO-263AB", "TVS", "DO-214AC", "SOD-80"}

        # Assign SDJ = 3 for diode packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuesdid), "SDJ"] = 3

        # Define the set of values to check for SOIC packages
        sdj_valuessoic8 = {"SOP8", "8SOP", "SOIC8", "8PIN-SOIC"}

        # Assign SDJ = 8 for SOIC packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuessoic8), "SDJ"] = 8
        
        # Define the set of values to check for SOIC packages
        sdj_valuessoic16 = {"SOIC16", "16SOP", "SOP16"}

        # Assign SDJ = 16 for SOIC packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuessoic16), "SDJ"] = 16

        # Reorder columns
        df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'SDJ']]
        #df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'DID', 'SOIC', 'THT', 'SDJ']]
        #df = df[['CRD', 'DES', 'MUF', 'MPN', 'QTY', 'TYPE', 'PKG', 'DID', 'SDJ', 'SPERESCOMP', 'SPETHTCOMP', 'SPESODCOMP', 'SPECAPCOMP', 'SPEFERINDCOMP']]

        save_cleaned_data(df, file_path)
        
    if __name__ == "__main__":
        process_excel()

        input("\nPress Enter to return to the menu...")
#-------------------------------------------------------------------------------------------------------------------
# ============================ Program 2: MPN Library Creator ============================
def pro_2():
    
    print("\nRunning MPN Library Creator (Program 2)...")
    
    print("\n")
    
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
        
        input("\nPress Enter to return to the menu...")  
#-------------------------------------------------------------------------------------------------------------------
# ============================ Program 3: MPN Library Lookup EDH ============================
def pro_3():
    
    print("\nRunning MPN Library Lookup EDH (Program 3)...")
    
    print("\n")
    
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
        
        input("\nPress Enter to return to the menu...")  
        
#-------------------------------------------------------------------------------------------------------------------

def main():
    while True:  
        print("\033[92;40mBOM EDH\033[0m \033[1;34;40mSYRMA\033[0m \033[1;36;40mSGS\033[0m \n\n\033[92;40mBOM EDH PY V-2.0.1 APR|04|04|2025 Select a program: \033[0m")
        print("\n")
        print("\033[1;36;40mProgramming Steps\033[0m")
        print("\n")
        print("1. BOM EDH Analyser: V-1.0.1")  
        print("2. MPN Library Creator: V-1.0.0")  
        print("3. MPN Library Lookup EDH: V-1.0.0")  
        print("\n")
        print("X. \033[1;31;40mExit Program\033[0m")  
        print("\n")

        choice = input("\033[1;36;40mChoose the program number: \033[0m")

        if choice == '1':
            pro_1()
        elif choice == '2':
            pro_2()
        elif choice == '3':
            pro_3()
        elif choice.upper() == 'X':    
            print("\n")
            print("\033[1;31;40mExiting the program.\033[0m")
            print("\nThank You")
            
            root = tk.Tk()
            root.withdraw()  
            messagebox.showinfo("Program Terminated", "Exiting the Program")
            sys.exit()
        else:
            print("Invalid choice. Please try again.\n")
            input("\nPress Enter to return to the menu...")  

if __name__ == "__main__":
    main()
    
#pyinstaller -F -i CRYPTO.ico Manipulator.py
#pyinstaller --onefile --windowed --name BOM-EDH BOM.ico EDH.py
#pyinstaller --onefile --console --name BOM-EDH --icon=BOM.ico EDH.py
#pyinstaller -F --onefile --console --name BOM-EDH --icon=BOM.ico EDH.py