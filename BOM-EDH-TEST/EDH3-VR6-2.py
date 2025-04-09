import pandas as pd
from tkinter import filedialog, Tk
import os

# Paths
MPN_LIBRARY_PATH = r"D:\MPNlibrary\MPNLibrary.xlsx"

# Load MPNLibrary
if not os.path.exists(MPN_LIBRARY_PATH):
    print("❌ MPNLibrary.xlsx not found.")
    exit()

df_library = pd.read_excel(MPN_LIBRARY_PATH, usecols=['MPN', 'TYPE', 'CID', 'PKG', 'SDJ'])
df_library['MPN'] = df_library['MPN'].astype(str).str.strip()

# Select BOM file
root = Tk()
root.withdraw()
bom_path = filedialog.askopenfilename(title="Select BOM file", filetypes=[("Excel files", "*.xls;*.xlsx")])

if not bom_path:
    print("❌ No BOM file selected.")
    exit()

df_bom = pd.read_excel(bom_path)
df_bom['MPN'] = df_bom['MPN'].astype(str).str.strip()

# Merge BOM with library, keeping only the columns to update
df_updated = df_bom.merge(df_library, on='MPN', how='left', suffixes=('', '_lib'))

# Update only TYPE, CID, PKG, SDJ in BOM if MPN matches
for col in ['TYPE', 'CID', 'PKG', 'SDJ']:
    df_updated[col] = df_updated[f'{col}_lib'].combine_first(df_updated[col])
    df_updated.drop(columns=[f'{col}_lib'], inplace=True)

# Save updated BOM
updated_path = os.path.splitext(bom_path)[0] + "_UpdatedWithMPNLibrary.xlsx"
df_updated.to_excel(updated_path, index=False)
print(f"✅ BOM updated and saved at: {updated_path}")
