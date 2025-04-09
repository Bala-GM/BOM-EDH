import pandas as pd
from tkinter import filedialog, Tk
import os

# Setup paths
MPN_LIBRARY_PATH = r"D:\MPNlibrary\MPNLibrary.xlsx"

# Load MPN library
if not os.path.exists(MPN_LIBRARY_PATH):
    print("❌ MPNLibrary not found.")
    exit()

df_library = pd.read_excel(MPN_LIBRARY_PATH)

# Select BOM file
root = Tk()
root.withdraw()
bom_path = filedialog.askopenfilename(title="Select BOM file", filetypes=[("Excel files", "*.xls;*.xlsx")])

if bom_path:
    df_bom = pd.read_excel(bom_path)

    # Ensure MPN column exists
    if 'MPN' not in df_bom.columns:
        print("❌ 'MPN' column not found in BOM.")
        exit()

    # Merge BOM with library
    df_updated = df_bom.merge(df_library[['MPN', 'TYPE', 'CID', 'PKG', 'SDJ']], on='MPN', how='left')

    # Save updated BOM
    updated_path = os.path.splitext(bom_path)[0] + "_Updated.xlsx"
    df_updated.to_excel(updated_path, index=False)
    print(f"✅ BOM updated and saved at: {updated_path}")
else:
    print("⚠️ No file selected.")
