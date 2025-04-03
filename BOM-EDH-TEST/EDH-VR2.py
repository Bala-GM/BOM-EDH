import os
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# py FLASK VER-1.0.0 APR|02|04|2025

def browse_file():
    """Opens a file dialog to select an Excel file."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def validate_data(df, required_columns):
    """Checks for missing columns and empty values."""
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
    """Saves validation errors to a log file."""
    error_file = os.path.join(os.path.dirname(file_path), "error_log.txt")
    with open(error_file, "w") as f:
        f.write("\n".join(errors))
    messagebox.showinfo("Error Log Saved", f"Errors saved in {error_file}")

def save_cleaned_data(df, file_path):
    """Saves the cleaned data as an Excel file."""
    save_name = input("Enter the name for the cleaned file: ")
    save_name = f"{save_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    save_path = os.path.join(os.path.dirname(file_path), save_name)
    df.to_excel(save_path, index=False, engine="openpyxl")
    messagebox.showinfo("File Saved", f"Cleaned data saved as {save_path}")

def extract_shape(description):
    """Extracts shape (0201, 0402, 0603, etc.) from the DES column."""
    desired_shapes = {"0201", "0402", "0603", "0805", "1206"}
    shape_match = re.search(r'\b\d{4}\b', str(description))  # Finds a 4-digit number
    if shape_match:
        shape = shape_match.group()
        if shape in desired_shapes:
            return shape
    return None

def process_excel():
    """Handles the complete process: file selection, validation, shape extraction, and saving."""
    file_path = browse_file()
    if not file_path:
        messagebox.showerror("Error", "No file selected")
        return
    
    # **Ensure correct file type**
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in [".xlsx", ".xls"]:
        messagebox.showerror("Error", "Selected file is not an Excel file.")
        return

    # **Read the Excel file safely**
    try:
        engine = "openpyxl" if file_ext == ".xlsx" else "xlrd"  # Correct engine selection
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

    # **Extract Shape Data**
    df["SHP"] = df["DES"].apply(extract_shape)

    # **Save the cleaned data with SHP column**
    save_cleaned_data(df[["CRD", "DES", "MUF", "MPN", "QTY", "SHP"]], file_path)

if __name__ == "__main__":
    process_excel()
