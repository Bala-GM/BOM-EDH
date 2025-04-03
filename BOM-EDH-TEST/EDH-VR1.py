import os
import pandas as pd
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
    df.to_excel(save_path, index=False)
    messagebox.showinfo("File Saved", f"Cleaned data saved as {save_path}")

def process_excel():
    file_path = browse_file()
    if not file_path:
        messagebox.showerror("Error", "No file selected")
        return
    
    df = pd.read_excel(file_path, dtype=str)
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
            save_cleaned_data(df[required_columns], file_path)
        else:
            save_error_log(file_path, missing_data)
            messagebox.showerror("Process Aborted", "Process was aborted due to missing data.")
    else:
        save_cleaned_data(df[required_columns], file_path)

if __name__ == "__main__":
    process_excel()
