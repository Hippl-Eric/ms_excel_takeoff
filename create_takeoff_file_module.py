import os
import tkinter as tk
from tkinter import filedialog
from dotenv import load_dotenv
from openpyxl import load_workbook
from takeoff import create_new_takeoff

def main():
    
    # Request and parse user inputs
    project_name = input("Project Name: ")
    
    # Number of soldier beams
    while True:
        try:
            num_rows = int(input("Number of SBs: ")) - 1
            if num_rows < 0:
                raise ValueError
            break
        except ValueError:
            print("Number of SBs must be an integer greater than 0")
    
    # Determine drilled or driven
    install = input("Drilled? (Y/N): ")
    if install.lower() == "no" or install.lower() == "n":
        drilled = False
    else:
        drilled = True

    # Load directory locations and template file
    load_dotenv()
    template_dir = os.getenv("TEMPLATE_DIR")
    template_file = os.getenv("TEMPLATE_FILE")
    template_file_path = os.path.join(template_dir, template_file)
    
    # Set the default destinate directory
    dest_dir = os.getenv("BID_DIR")
    dest_dir = os.path.join(dest_dir, project_name, "PRICING")
    
    # Load template file workbook
    while True:
        try:
            wb = load_workbook(filename = template_file_path)
            break
        except PermissionError as e:
            print(e)
            input("Please close template file.  Press ENTER to continue... or Ctrl-c to quit.")

    create_new_takeoff(wb, project_name, num_rows, drilled)
    
    # Ask user for destination directory and file name
    while True:
        try:
            dest_file = (tk.filedialog.asksaveasfile(
                mode='w', 
                initialdir=dest_dir,
                initialfile=f"Takeoff - {project_name}", 
                defaultextension=".xlsx")
            )
            break
        except PermissionError as e:
            print(e)
            input("You are about to overwrite this file.  Please close file, and press ENTER.  Otherwise press Crtl-c to quit.")
    
    # Save the workbook
    try:
        wb.save(dest_file.name)
        
    # User clicked cancel in tkinter "asksaveasfile" prompt, don't save file
    except AttributeError:
        pass

if __name__ == "__main__":
    main()