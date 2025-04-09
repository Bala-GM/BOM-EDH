import os
import sys
import pandas as pd
import re
import tkinter as tk
from tkinter import Tk, filedialog
from tkinter import filedialog, messagebox
from datetime import datetime

# PY VER-2.0.2 APR|08|04|2025

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
    
    def extract_shape(description): #chip-component
        desired_shapes = {"01005", "0201", "0402", "0603", "0805", "1008", "1206", "1210", "1806", "1812", "2010", "2512", "2920"}
        match = re.search(r'\b\d{4}\b', str(description))
        return match.group() if match and match.group() in desired_shapes else None

    def extract_shapesmdtht(description):
        """Extracts SMD / THT package shape from the DES column."""
        desired_shapes = {"Thru-Hole", "THT", "SMD", "SMT", "RADIAL", "Thru Hole", "THD", "AXIAL",
                          "Surface Mount Solder", "Surface Mount Solder Pad", "T/R", "RAD", "AWG",
                          "long,life", "MKP", "CAP,TH", "Disc", "HEADER", "0805", "1206", "KINKED",
                          "Kink Lead", "TAPED ON REEL"} #"""DIP,SDIP"""
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapescid(description):
        """Extracts Component ID package shape from the DES column."""
        desired_shapes = {"Transformer", "SOD", "SOT", "CAP", "COG", "NPO", "MLCC", "X7R", "IND" ,"Ferrite", "Choke", "Thermistor",
                          "Diode", "Rectifier", "Zener", "MELF", "Fuse", "varistor", "Transistor", "Crystal", "MOSFET", "RES"} #"""DIP,SDIP"""
        #ESD Suppressor TVS Uni-Dir 24V 2-Pin SOD-323 T/R == taking res for Diode package
        #ESD Suppressor TVS Bi-Dir 5V Automotive AEC-Q101 3-Pin SOT-23 T/R == taking res for SOT package
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapesmdchip(description):
        """Extracts SMD chip package shape from the DES column."""
        desired_shapes = {"01005", "0201", "0402", "0603", "0805", "1008", "1206", "1210", "1806", "1812", "2010", "2512", "2920"} #"""DIP,SDIP"""
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapechk(description): #diode
        """Extracts SOD SOT package shape from the DES column."""
        desired_shapes = {"SOD", "SOT"}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapediode(description): #diode
        """Extracts 2 pin/ 2 pos diode package shape from the DES column."""
        desired_shapes = {"Mini-MELF", "SOD-123", "TO-263AB", "DO-214AC", "DO-214-AC", "LS4148", "SOD-80C-2", "SMA", "SOD-323", "SOD-123", 
                          "SMA", "DO-214AC", "SMB", "Mini-MELF", "SMP", "SMAFL", "DO-214AA", "SMBFL", "SMAF", "SMBF", "SOD-123F", 
                          "DO-219AB", "SOD-123FL", "DO-219AD", "SC-76", "SOD-523", "SC-79", "SOD-882", "X1-DFN1006-2", "SOD-882L", 
                          "X2-DFN1006-2", "SOD-923", "TLM2D3D6", "SOD123", "SOD-80", "TUMD-M", "TUMD2M", "LED"} 
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapesot(description): #Transistor
        """Extracts SOT Transistor diode package shape from the DES column."""
        desired_shapes = {"SOT-23", "TO-236AB", "SOT-23F", "SOT-323", "SC-70", "SOT-523", "SC-89", "SOT-523W", "SC-75", "SOT-883L", "X2-DFN1006-3", 
                         "SOT-883VL", "X3-DFN1006-3", "D2PAK", "TO-263", "DPAK", "TO-252", "SOT-923", "TLM364", "TO-277A", "TSOT-26", "LFPAK",  
                         "PowerPAK", "TO-220-3", "TO-220AB", "SOT23-3", "SOT23-5", "SOT23-6", "SOT-223" ,"SOT252-2"}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapeic(description): #IC
        """Extracts SOIC QFN BGA package shape from the DES column."""
        desired_shapes = {"SOIC", "UQFN", "LQFP", "SOP", "TSOP", "QFP", "QFN", "BGA", "µBGA", "CBGA", "PBGA", "CSP", "DFN", "PGA", "LGA", "SiP", "MCM", "DIP",
                          "SOJ", "PLCC", "DIMM", "SIMM", "TSSOP", "VSSOP",
                          "SOIC8", "8SOIC", "SOIC-8", "8-SOIC", "8SOP", "SOP8",
                          "SOIC12", "12SOIC", "SOIC-12", "12-SOIC", "12SOP", "SOP12",
                          "SOIC14", "14SOIC", "SOIC-14", "14-SOIC", "14SOP", "SOP14",
                          "SOIC16", "16SOIC", "SOIC-16", "16-SOIC", "16SOP", "SOP16",}
        for shape in desired_shapes:
            if shape.lower() in str(description).lower():
                return shape
        return None
    
    def extract_shapepinsmt(description):
        """Extracts PIN package shape from the description."""
        # Define the desired shapes as a regex pattern
        desired_shapes = r"\b(300-pin|299-pin|298-pin|297-pin|296-pin|295-pin|294-pin|293-pin|292-pin|291-pin|290-pin|289-pin|288-pin|287-pin|286-pin|285-pin|284-pin|283-pin|282-pin|281-pin|280-pin|279-pin|278-pin|277-pin|276-pin|275-pin|274-pin|273-pin|272-pin|271-pin|270-pin|269-pin|268-pin|267-pin|266-pin|265-pin|264-pin|263-pin|262-pin|261-pin|260-pin|259-pin|258-pin|257-pin|256-pin|255-pin|254-pin|253-pin|252-pin|251-pin|250-pin|249-pin|248-pin|247-pin|246-pin|245-pin|244-pin|243-pin|242-pin|241-pin|240-pin|239-pin|238-pin|237-pin|236-pin|235-pin|234-pin|233-pin|232-pin|231-pin|230-pin|229-pin|228-pin|227-pin|226-pin|225-pin|224-pin|223-pin|222-pin|221-pin|220-pin|219-pin|218-pin|217-pin|216-pin|215-pin|214-pin|213-pin|212-pin|211-pin|210-pin|209-pin|208-pin|207-pin|206-pin|205-pin|204-pin|203-pin|202-pin|201-pin|200-pin|199-pin|198-pin|197-pin|196-pin|195-pin|194-pin|193-pin|192-pin|191-pin|190-pin|189-pin|188-pin|187-pin|186-pin|185-pin|184-pin|183-pin|182-pin|181-pin|180-pin|179-pin|178-pin|177-pin|176-pin|175-pin|174-pin|173-pin|172-pin|171-pin|170-pin|169-pin|168-pin|167-pin|166-pin|165-pin|164-pin|163-pin|162-pin|161-pin|160-pin|159-pin|158-pin|157-pin|156-pin|155-pin|154-pin|153-pin|152-pin|151-pin|150-pin|149-pin|148-pin|147-pin|146-pin|145-pin|144-pin|143-pin|142-pin|141-pin|140-pin|139-pin|138-pin|137-pin|136-pin|135-pin|134-pin|133-pin|132-pin|131-pin|130-pin|129-pin|128-pin|127-pin|126-pin|125-pin|124-pin|123-pin|122-pin|121-pin|120-pin|119-pin|118-pin|117-pin|116-pin|115-pin|114-pin|113-pin|112-pin|111-pin|110-pin|109-pin|108-pin|107-pin|106-pin|105-pin|104-pin|103-pin|102-pin|101-pin|100-pin|99-pin|98-pin|97-pin|96-pin|95-pin|94-pin|93-pin|92-pin|91-pin|90-pin|89-pin|88-pin|87-pin|86-pin|85-pin|84-pin|83-pin|82-pin|81-pin|80-pin|79-pin|78-pin|77-pin|76-pin|75-pin|74-pin|73-pin|72-pin|71-pin|70-pin|69-pin|68-pin|67-pin|66-pin|65-pin|64-pin|63-pin|62-pin|61-pin|60-pin|59-pin|58-pin|57-pin|56-pin|55-pin|54-pin|53-pin|52-pin|51-pin|50-pin|49-pin|48-pin|47-pin|46-pin|45-pin|44-pin|43-pin|42-pin|41-pin|40-pin|39-pin|38-pin|37-pin|36-pin|35-pin|34-pin|33-pin|32-pin|31-pin|30-pin|29-pin|28-pin|27-pin|26-pin|25-pin|24-pin|23-pin|22-pin|21-pin|20-pin|19-pin|18-pin|17-pin|16-pin|15-pin|14-pin|13-pin|12-pin|11-pin|10-pin|9-pin|8-pin|7-pin|6-pin|5-pin|4-pin|3-pin|2-pin|1-pin)\b"
        
        # Search for the pattern in the description
        match = re.search(desired_shapes, str(description), re.IGNORECASE)
        
        # Return the found shape or None
        return match.group(0) if match else None
    
    def extract_shapepinpos(description):
        """Extracts PIN package shape (1 to 300 Pins) from the description."""
        # Regex: Match 1 to 300 followed by optional space or dash and a keyword
        pattern = r'\b(3[0-0][0]|2[0-9]{2}|1[0-9]{2}|[1-9][0-9]?|300)\s*[-]?\s*(Pin|Pins|POS|Pos|Positions|Feet|P|contacts?)\b'
        
        match = re.search(pattern, str(description), re.IGNORECASE)
        
        return match.group(0) if match else None
    
    def extract_shapepossmt(description):
        """Extracts POS package shape from the description."""
        # Define the desired shapes as a regex pattern
        desired_shapes = r"\b(300 pos|299 pos|298 pos|297 pos|296 pos|295 pos|294 pos|293 pos|292 pos|291 pos|290 pos|289 pos|288 pos|287 pos|286 pos|285 pos|284 pos|283 pos|282 pos|281 pos|280 pos|279 pos|278 pos|277 pos|276 pos|275 pos|274 pos|273 pos|272 pos|271 pos|270 pos|269 pos|268 pos|267 pos|266 pos|265 pos|264 pos|263 pos|262 pos|261 pos|260 pos|259 pos|258 pos|257 pos|256 pos|255 pos|254 pos|253 pos|252 pos|251 pos|250 pos|249 pos|248 pos|247 pos|246 pos|245 pos|244 pos|243 pos|242 pos|241 pos|240 pos|239 pos|238 pos|237 pos|236 pos|235 pos|234 pos|233 pos|232 pos|231 pos|230 pos|229 pos|228 pos|227 pos|226 pos|225 pos|224 pos|223 pos|222 pos|221 pos|220 pos|219 pos|218 pos|217 pos|216 pos|215 pos|214 pos|213 pos|212 pos|211 pos|210 pos|209 pos|208 pos|207 pos|206 pos|205 pos|204 pos|203 pos|202 pos|201 pos|200 pos|199 pos|198 pos|197 pos|196 pos|195 pos|194 pos|193 pos|192 pos|191 pos|190 pos|189 pos|188 pos|187 pos|186 pos|185 pos|184 pos|183 pos|182 pos|181 pos|180 pos|179 pos|178 pos|177 pos|176 pos|175 pos|174 pos|173 pos|172 pos|171 pos|170 pos|169 pos|168 pos|167 pos|166 pos|165 pos|164 pos|163 pos|162 pos|161 pos|160 pos|159 pos|158 pos|157 pos|156 pos|155 pos|154 pos|153 pos|152 pos|151 pos|150 pos|149 pos|148 pos|147 pos|146 pos|145 pos|144 pos|143 pos|142 pos|141 pos|140 pos|139 pos|138 pos|137 pos|136 pos|135 pos|134 pos|133 pos|132 pos|131 pos|130 pos|129 pos|128 pos|127 pos|126 pos|125 pos|124 pos|123 pos|122 pos|121 pos|120 pos|119 pos|118 pos|117 pos|116 pos|115 pos|114 pos|113 pos|112 pos|111 pos|110 pos|109 pos|108 pos|107 pos|106 pos|105 pos|104 pos|103 pos|102 pos|101 pos|100 pos|99 pos|98 pos|97 pos|96 pos|95 pos|94 pos|93 pos|92 pos|91 pos|90 pos|89 pos|88 pos|87 pos|86 pos|85 pos|84 pos|83 pos|82 pos|81 pos|80 pos|79 pos|78 pos|77 pos|76 pos|75 pos|74 pos|73 pos|72 pos|71 pos|70 pos|69 pos|68 pos|67 pos|66 pos|65 pos|64 pos|63 pos|62 pos|61 pos|60 pos|59 pos|58 pos|57 pos|56 pos|55 pos|54 pos|53 pos|52 pos|51 pos|50 pos|49 pos|48 pos|47 pos|46 pos|45 pos|44 pos|43 pos|42 pos|41 pos|40 pos|39 pos|38 pos|37 pos|36 pos|35 pos|34 pos|33 pos|32 pos|31 pos|30 pos|29 pos|28 pos|27 pos|26 pos|25 pos|24 pos|23 pos|22 pos|21 pos|20 pos|19 pos|18 pos|17 pos|16 pos|15 pos|14 pos|13 pos|12 pos|11 pos|10 pos|9 pos|8 pos|7 pos|6 pos|5 pos|4 pos|3 pos|2 pos|1 pos)\b"
        
        # Search for the pattern in the description
        match = re.search(desired_shapes, str(description), re.IGNORECASE)
        
        # Return the found shape or None
        return match.group(0) if match else None

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

        required_columns = ["CRD", "DES", "MFG", "MPN", "QTY"]
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
            
        # Create a new column 'DESC' as a copy of 'DES'
        df['DESC'] = df['DES'].copy()

        # Replace the word 'Suppressor' with an empty string in 'DESC'
        df['DESC'] = df['DESC'].str.replace('Suppressor', '', regex=False)
        df['DESC'] = df['DESC'].str.replace('Insulation Cap', 'Insulation', regex=False)
        df['DESC'] = df['DESC'].str.replace('Grade', '', regex=False)
            
        df["TYPE"] = df["DESC"].apply(extract_shapesmdtht)
        df["SHP"] = df["DESC"].apply(extract_shape)
        df["CID"] = df["DESC"].apply(extract_shapescid)
        df["CID-CHK"] = df["DESC"].apply(extract_shapechk)
        df["SHP-C"] = df["DESC"].apply(extract_shapesmdchip)
        df["SHP-D"] = df["DESC"].apply(extract_shapediode)
        df["SHP-T"] = df["DESC"].apply(extract_shapesot)
        df["SHP-I"] = df["DESC"].apply(extract_shapeic)
        df["SHP-PIN"] = df["DESC"].apply(extract_shapepinsmt)
        df["SHP-POS"] = df["DESC"].apply(extract_shapepossmt)
        df["SDJ-A2"] = df["DESC"].apply(extract_shapepinpos)
        
        # Standardize PKG names
        replacetype = {"Thru-Hole": "THT", "Thru Hole": "THT", "RADIAL": "THT", "AXIAL": "THT", "THD": "THT", "SMT": "SMD", "T/R": "SMD",
                       "Surface Mount Solder Pad": "SMD", "Surface Mount Solder": "SMD", "RAD": "THT", "AWG": "THT", "HEADER": "THT",
                       "long,life": "THT", "MKP": "THT", "CAP,TH": "THT", "Disc": "THT", "0805": "SMD", "1206": "SMD", "KINKED": "THT",
                       "Kink Lead": "THT", "TAPED ON REEL": "SMD"}
         
        df['TYPE'] = df['TYPE'].replace(replacetype)
        
        # Standardize CID names
        replactypeCID = {"SOD": "DIODE", "SOT": "TRANSISTOR", "MOSFET": "TRANSISTOR", "NPO": "CAPACITOR", "COG": "CAPACITOR", "MLCC": "CAPACITOR",
                         "Crystal": "CRYSTAL", "Ferrite": "INDUCTOR", "Choke": "INDUCTOR", "Rectifier": "DIODE", "X7R": "CAPACITOR", "MELF": "DIODE",
                         "Thermistor": "THERMISTOR", "Fuse": "FUSE", "Transformer": "TRANSFORMER", "CAP": "CAPACITOR", "Zener": "DIODE",
                          "varistor": "VARISTOR", "IND": "INDUCTOR", "Transistor": "TRANSISTOR", "Diode": "DIODE", "RES": "RESISTOR"}
         
        df['CID'] = df['CID'].replace(replactypeCID)
        
        df["CIDC"] = df[['CID', 'CID-CHK']].apply(lambda x: '-'.join(x.dropna().astype(str)), axis=1)
        
        # Standardize CIDC names
        replactypeCIDC = {"RESISTOR-SOD": "DIODE", "RESISTOR-SOT": "TRANSISTOR", "DIODE-SOD": "DIODE", "DIODE-SOT": "TRANSISTOR", "TRANSISTOR-SOT": "TRANSISTOR"}
         
        df['CIDC'] = df['CIDC'].replace(replactypeCIDC)
        
        df["SHP-C"] = df[['CID', 'SHP-C']].apply(lambda x: '-'.join(x.dropna().astype(str)), axis=1)
        
        # Replace the word 'Suppressor' with an empty string in SHP-C'
        df['SHP-C'] = df['SHP-C'].str.replace('DIODE', '', regex=False)
        df['SHP-C'] = df['SHP-C'].str.replace('TRANSISTOR', '', regex=False)
        df['SHP-C'] = df['SHP-C'].str.replace('TRANSFORMER', '', regex=False)
        df['SHP-C'] = df['SHP-C'].str.replace('CAPACITOR-', 'C', regex=False)
        df['SHP-C'] = df['SHP-C'].str.replace('INDUCTOR-', 'I', regex=False)
        df['SHP-C'] = df['SHP-C'].str.replace('RESISTOR-', 'R', regex=False)
        
        df["SHP"] = df[['SHP-C', 'SHP-D', 'SHP-T', 'SHP-I']].apply(lambda x: ''.join(x.dropna().astype(str)), axis=1)

        # Reorder columns
        df = df[['CRD', 'DES', 'MFG', 'MPN', 'QTY', 'TYPE', 'CIDC', 'SHP', 'SHP-C', 'SHP-D', 'SHP-T', 'SHP-I', 'SHP-PIN', 'SHP-POS', 'SDJ-A2']]
        
        # Rename SHP to PKG
        df.rename(columns={'SHP': 'PKG'}, inplace=True)
        df.rename(columns={'CIDC': 'CID'}, inplace=True)
        
        df["SDJ-A"] = df[['SHP-PIN', 'SHP-POS']].apply(lambda x: ''.join(x.dropna().astype(str)), axis=1)
        
        #df["SDJ-A2"] = df[['SHP-PIN1', 'SHP-PIN2', 'SHP-PIN3']].apply(lambda x: ''.join(x.dropna().astype(str)), axis=1)
        
        # Define the set of values to check for SDJ column
        sdj_values = {"C01005", "C0201", "C0402", "C0603", "C0805", "C1008", "C1206", "C1210", "C1806", "C1812", "C2010", "C2512", "C2920",
                      "R01005", "R0201", "R0402", "R0603", "R0805", "R1008", "R1206", "R1210", "R1806", "R1812", "R2010", "R2512", "R2920",
                      "I01005", "I0201", "I0402", "I0603", "I0805", "I1008", "I1206", "I1210", "I1806", "I1812", "I2010", "I2512", "I2920",
                      "SOD-323", "SOD-123", "SMA", "DO-214AC", "SMB",
                      "Mini-MELF", "SMP", "SMAFL", "DO-214AA", "SMBFL", "SMAF", "SMBF",
                      "SOD-123F", "DO-219AB", "SOD-123FL", "DO-219AD", "SC-76", "SOD-523", "SC-79",
                      "SOD-882", "X1-DFN1006-2", "SOD-882L", "X2-DFN1006-2", "SOD-923", "TLM2D3D6"}

        # Initialize SDJ-M column with default empty value 2
        df["SDJ-M"] = ""

        # Assign SDJ = 2 for specific PKG values
        df.loc[df["PKG"].isin(sdj_values), "SDJ-M"] = '2-Pin' #2

        # Define the set of values to check for diode packages 3
        sdj_valuesdid3 = {"SOT-23", "TO-236AB", "SOT-23F", "SOT-323", "SC-70", "SOT-523", "SC-89", "SOT-523W",
                         "SC-75", "SOT-883L", "X2-DFN1006-3", "SOT-883VL", "X3-DFN1006-3", "D2PAK", "TO-263",
                         "DPAK", "TO-252", "SOT-923", "TLM364", "TO-277A"}

        # Assign SDJ-M = 3 for diode packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuesdid3), "SDJ-M"] = '3-Pin' #3
        
        # Define the set of values to check for diode packages 4
        sdj_valuesdid4 = {"HD DIP", "TO-269AA", "LPDIP", "SMDIP", "SOT-143", "TO-253AA", "SOT-223",
                          "TO-261AA", "SOT-543", "SC-75-4", "SOT-89", "TO-243AA", "SC-82AB"}

        # Assign SDJ-M = 4 for diode packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuesdid4), "SDJ-M"] = '4-Pin' #4
        
        # Define the set of values to check for diode packages 5
        sdj_valuesdid5 = {"SOT-953", "SOT-89-5", "SOT23-5", "SOT-23-5", "SC-88A"}

        # Assign SDJ-M = 5 for diode packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuesdid5), "SDJ-M"] = '5-Pin' #5
        
        # Define the set of values to check for diode packages 6
        sdj_valuesdid6 = {"SOT-26", "SOT23-6", "SC-74", "SOT-363", "SSC-88", "SOT-563", "SC-89-6", "SOT-963", "TLM621"}

        # Assign SDJ-M = 6 for diode packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuesdid6), "SDJ-M"] = '6-Pin' #6

        # Define the set of values to check for SOIC packages
        sdj_valuessoic8 = {"SOP8", "8SOP", "SOIC8", "8PIN-SOIC"}

        # Assign SDJ-M = 8 for SOIC packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuessoic8), "SDJ-M"] = '8-Pin' #8
        
        # Define the set of values to check for SOIC packages
        sdj_valuessoic16 = {"SOIC16", "16SOP", "SOP16"}

        # Assign SDJ-M = 16 for SOIC packages without erasing previous values
        df.loc[df["PKG"].isin(sdj_valuessoic16), "SDJ-M"] = '16-Pin' #16 
        
        # Reorder columns
        df = df[['CRD', 'DES', 'MFG', 'MPN', 'QTY', 'TYPE', 'CID', 'PKG', 'SDJ-A', 'SDJ-M', 'SDJ-A2']]
        
        def compare_sdj(row):
            val_a = str(row['SDJ-A']).strip()
            val_m = str(row['SDJ-M']).strip()
            
            if not val_a and not val_m:
                return ''
            elif val_a and not val_m:
                return val_a
            elif val_m and not val_a:
                return val_m
            elif val_a == val_m:
                return val_a
            else:
                return f'{val_a}/{val_m}'

        df['SDJ'] = df.apply(compare_sdj, axis=1)

        # Reorder columns
        df = df[['CRD', 'DES', 'MFG', 'MPN', 'QTY', 'TYPE', 'CID', 'PKG', 'SDJ-A', 'SDJ-M', 'SDJ-A2', 'SDJ']]
        
        save_cleaned_data(df, file_path)
        
    if __name__ == "__main__":
        process_excel()

        input("\nPress Enter to return to the menu...")
#-------------------------------------------------------------------------------------------------------------------
# ============================ Program 2: MPN Library Creator ============================
def pro_2():
    
    print("\nRunning MPN Library Creator (Program 2)...")
    
    print("\n")
    
    import os
    import pandas as pd
    import tkinter as tk
    from tkinter import filedialog
    from datetime import datetime

    # Constants
    TARGET_FOLDER = r'D:\MPNlibrary'
    MPN_LIBRARY_FILE = os.path.join(TARGET_FOLDER, 'MPNLibrary.xlsx')
    MISMATCH_FILE = os.path.join(TARGET_FOLDER, 'MissMatch_MPNLibrary.xlsx')
    DUPLICATE_LOG = os.path.join(TARGET_FOLDER, 'Duplicate_Entries.txt')
    REQUIRED_COLUMNS = ['MPN', 'TYPE', 'CID', 'PKG', 'SDJ']

    # Ensure target folder exists
    os.makedirs(TARGET_FOLDER, exist_ok=True)

    def browse_and_append():
        # Hide root tkinter window
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            print("No file selected.")
            return

        try:
            excel_file_name = os.path.basename(file_path)
            current_time_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

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

            append_list = []
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
                        duplicate_entries.append((new_row.to_dict(), excel_file_name, current_time_str))
                        continue
                    else:
                        for _, mismatch_row in mpn_matches.iterrows():
                            mismatch_records.append(mismatch_row)
                        mismatch_records.append(new_row)
                        continue

                append_list.append(new_row)

            if append_list:
                final_append_df = pd.DataFrame(append_list)
                updated_df = pd.concat([existing_df, final_append_df], ignore_index=True)
                updated_df.to_excel(MPN_LIBRARY_FILE, index=False)
                print(f"{len(append_list)} new record(s) appended to MPNLibrary.xlsx")
            else:
                print("No new records to append.")

            if mismatch_records:
                mismatch_df = pd.DataFrame(mismatch_records)
                mismatch_df.to_excel(MISMATCH_FILE, index=False)
                print(f"{len(mismatch_records)} mismatch record(s) saved to MissMatch_MPNLibrary.xlsx")

            if duplicate_entries:
                with open(DUPLICATE_LOG, 'a') as f:
                    for entry, file_name, timestamp in duplicate_entries:
                        line = ', '.join([str(entry[col]) for col in REQUIRED_COLUMNS])
                        f.write(f"{line} | File: {file_name} | Time: {timestamp}\n")
                print(f"{len(duplicate_entries)} duplicate record(s) logged in Duplicate_Entries.txt")

        except Exception as e:
            print(f"ERROR: {e}")

    if __name__ == "__main__":
        browse_and_append()
            
    input("\nPress Enter to return to the menu...")
#-------------------------------------------------------------------------------------------------------------------
# ============================ Program 3: MPN Library Lookup EDH ============================
def pro_3():
    
    print("\nRunning MPN Library Lookup EDH (Program 3)...")
    
    print("\n")
    
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

            
    input("\nPress Enter to return to the menu...")
    
#-------------------------------------------------------------------------------------------------------------------

def main():
    while True:  
        print("\033[92;40mBOM EDH\033[0m \033[1;34;40mSYRMA\033[0m \033[1;36;40mSGS\033[0m \n\n\033[92;40mBOM EDH PY V-4.0.0 APR|08|04|2025  \033[0m")
        print("\n")
        print("\033[1;36;40mSelect a program:\033[0m")
        print("\n")
        print("1. BOM EDH Analyser: V-2.0.0")  
        print("2. MPN Library Creator: V-2.0.1")  
        print("3. MPN Library Lookup EDH: V-2.0.1")  
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
            messagebox.showinfo("Program Description", "PROGRAM NAME 'BOM-EDH-VR10,1VR-1,2VR-6-8,3VR-6-2'\n\n'PY V-4.0.0 APR|08|04|2025' brief note on STANDALONE SOFTWARE 'Application to help BOM segeration'\n\nMIT License\n\nCopyright (C) <2025>  <BALA GANESH>\n\n")
            sys.exit()
        else:
            print("Invalid choice. Please try again.\n")
            input("\nPress Enter to return to the menu...")

if __name__ == "__main__":
    main()
    
#pyinstaller -F --onefile --console --name BOM-EDH --icon=BOM.ico EDH.py