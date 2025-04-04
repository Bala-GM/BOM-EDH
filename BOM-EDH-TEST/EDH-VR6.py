
import os
import pandas as pd
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# PY VER-1.0.1 APR|03|04|2025

from EDH1 import pro_1
from EDH2 import pro_2
from EDH3 import pro_3

def main():
# Display menu
    print("\033[92;40mBOM EDH\033[0m \033[1;34;40mSYRMA\033[0m \033[1;36;40mSGS\033[0m \n\n\033[92;40mBOM EDH PY V-2.0.1 APR|04|04|2025 Select a program: \033[0m")
    print("\n")
    print("\033[1;36;40mPrograming Steps\033[0m")
    print("\n")
    print("1. BOM EDH Analyser: V-1.0.1") #89P13
    print("2. MPNLibrary Creater: V-1.0.0") #89P13
    print("3. MPNLibrary Lookup EDH: V-1.0.0") #89P13
    print("\n")
    print("X. \033[1;31;40mExit Program\033[0m") #70599
    print("\n")
    
# Get user choice
    choice = input("\033[1;36;40mChoose the program number: \033[0m")

    # Run the selected program
    if choice == '1':
        pro_1()
    elif choice == '2':
        pro_2()
    elif choice == '3':
        pro_3()
    elif choice == 'X':    
  
        print("\n")
        print("\033[1;31;40mExiting the program.\033[0m")
        print("\nThank You")
        
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showinfo("Program Terminated", "Exiting the Program")
        sys.exit()
    else:
        print("Invalid choice. Exiting.")
        print("\nThank You")

    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Program Terminated", "Enter Invalid choice!")
    sys.exit()
    
if __name__ == "__main__":
    main()
    
    #pyinstaller -F -i CRYPTO.ico Manipulator.py
      
#pyinstaller --onefile --windowed --name BOM-EDH BOM.ico EDH.py
#pyinstaller --onefile --console --name BOM-EDH --icon=BOM.ico EDH.py
