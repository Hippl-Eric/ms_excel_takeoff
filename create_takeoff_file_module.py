import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from takeoff import create_new_takeoff

def main():
    # Load template file
    temp_file = tk.filedialog.askopenfilename()
    wb = load_workbook(temp_file)
    dest_file = tk.filedialog.asksaveasfile(mode='w', initialfile="Takeoff - Project Name", defaultextension=".xlsx")
    wb.save(dest_file.name)
    pass
    # Load workbook
    # Create new takeoff
    # Prompt where to save file
    # Save

if __name__ == "__main__":
    main()