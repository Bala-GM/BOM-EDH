import os
import pandas as pd
import numpy as np
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# py FLASK VER-1.0.0 APR|02|04|2025

def browse_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def validate_data(df, required_columns):
    missing_data = []
    for col in required_columns:
        if col not in df.columns:
            missing_data.append(f"Missing column: {col}")
        else:
            for index, value in df[col].items():
                if pd.isna(value) or value == "":
                    missing_data.append(f"Row {index + 2}, Column '{col}' is missing.")
    return missing_data

def save_error_log(file_path, errors):
    error_file = os.path.join(os.path.dirname(file_path), "error_log.txt")
    with open(error_file, "w") as f:
        f.write("\n".join(errors))
    messagebox.showinfo("Error Log Saved", f"Errors saved in {error_file}")

def save_cleaned_data(df, file_path):
    save_name = input("Enter the name for the cleaned file: ")
    save_name = f"{save_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    save_path = os.path.join(os.path.dirname(file_path), save_name)
    df.to_excel(save_path, index=False, engine="openpyxl")
    messagebox.showinfo("File Saved", f"Cleaned data saved as {save_path}")

def extract_shape(description):
    desired_shapes = {"0201", "0402", "0603", "0805", "1206"}
    shape_match = re.search(r'\b\d{4}\b', str(description))
    return shape_match.group() if shape_match and shape_match.group() in desired_shapes else None

def extract_component_type(DES):
    desired_COMP = ("CAP", "RES", "IND")
    for comp_type in desired_COMP:
        if comp_type.lower() in str(DES).lower():
            return comp_type
    return None

def check_melf_resistor(DES):
    return "MELF" if "MELF" in str(DES).upper() else ""

def check_tht_component(DES):
    desired_SPETHTCOMP = ("THT", "DIP")
    return "THT" if any(comp.lower() in str(DES).lower() for comp in desired_SPETHTCOMP) else ""

def extract_sod_component(DES):
    desired_SPESODCOMP = ("ZENER", "DIODE", "SOD")
    for sod_type in desired_SPESODCOMP:
        if sod_type.lower() in str(DES).lower():
            return sod_type
    return ""

def extract_capacitor_type(DES):
    desired_SPECAPCOMP = ("TAN", "TANTALUM", "ALUMINIUM", "ALLUM", "ALUM", "ELECTROLYTIC", "ALU")
    for cap_type in desired_SPECAPCOMP:
        if cap_type.lower() in str(DES).lower():
            return cap_type
    return ""

def extract_ferrite_inductor(DES):
    desired_SPEFERINDCOMP = ("IND", "FERRITEBEAD", "FERRITE", "BEAD", "INDUCTOR")
    for ind_type in desired_SPEFERINDCOMP:
        if ind_type.lower() in str(DES).lower():
            return ind_type
    return ""

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
        engine = "openpyxl" if file_ext == ".xlsx" else "xlrd"
        df = pd.read_excel(file_path, dtype=str, engine=engine)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the file:\n{e}")
        return

    required_columns = ["CRD", "DES", "MUF", "MPN", "QTY"]
    missing_data = validate_data(df, required_columns)

    if missing_data:
        error_message = "\n".join(missing_data)
        user_choice = messagebox.askyesnocancel("Data Issues Found", f"{error_message}\n\nReplace missing values with NaN and continue?")
        
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
    df["TYPE"] = df["DES"].apply(extract_component_type)
    df["SPERESCOMP"] = df["DES"].apply(check_melf_resistor)
    df["SPETHTCOMP"] = df["DES"].apply(check_tht_component)
    df["SPESODCOMP"] = df["DES"].apply(extract_sod_component)
    df["SPECAPCOMP"] = df["DES"].apply(extract_capacitor_type)
    df["SPEFERINDCOMP"] = df["DES"].apply(extract_ferrite_inductor)

    save_cleaned_data(df[["CRD", "DES", "MUF", "MPN", "QTY", "SHP", "TYPE", "SPERESCOMP", "SPETHTCOMP", "SPESODCOMP", "SPECAPCOMP", "SPEFERINDCOMP"]], file_path)

if __name__ == "__main__":
    process_excel()
