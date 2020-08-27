import os

from openpyxl import load_workbook
from dotenv import load_dotenv

from helpers import cell_search, copy_row, fix_sum_row_cells

def main():

    # Request and parse user inputs
    project_name = input("Project Name: ")
    num_rows = int(input("Number of SBs: ")) - 1
    drilled = True
    install = input("Drilled? (Y/N): ")
    if install == "no" or install == "No" or install == "n" or install == "N" or install == "NO":
        drilled = False

    # Load directory locations and template file
    load_dotenv()
    init_directory = os.getenv("TEMPLATE_DIR") # template location
    dest_directory = os.getenv("BID_DIR") # bids location
    template_file = os.getenv("TEMPLATE_FILE")

    create_new_takeoff(template_file, project_name, num_rows, drilled, init_directory, dest_directory)

    # Success
    return print("Success")

def create_new_takeoff(template_file, project_name, num_rows, drilled, temp_dir = "", dest_dir = ""):
"""Description

[template_file] string filename.  Cannot start with "\""
[project_name] string
[num_rows] int
[drilled] bool
[temp_dir] Optional string. Directory of template file. Default is current directory.
[dest_dir] Optional string. Directory to save new file. Default is current directory.
"""

    # Load base template file
    wb = load_workbook(filename = f"{temp_dir}{template_file}")
    ws = wb["Takeoff-SB"]

    # Load cells and determine number of columns
    all_cells = tuple(ws.rows)
    num_col = len(all_cells[0])

    # Search for project name, drilled/driven, and first row locations
    name_cell = cell_search(all_cells, "PROJECTNAME -  TAKEOFF")
    drill_cell = cell_search(all_cells, "DRIVEN/DRILLED")
    title_row_cell = cell_search(all_cells, "SB Nos.")

    # Input project name
    name_cell.value = f"{project_name} - Takeoff"

    # Input drilled or driven
    if drilled:
        drill_cell.value = "DRILLED"
    else:
        drill_cell.value = "DRIVEN"

    # Locate the first cell row
    first_cell_row_index = title_row_cell.row + 1

    # Copy the first row values and styling, and paste on all added rows
    for first_cell_row in ws.iter_rows(min_row = first_cell_row_index, max_row = first_cell_row_index, max_col = num_col):
        copy_row(ws, first_cell_row, num_rows)

    # TODO change value for # SB column C

    # Correct forumlas in sum row
    sum_row_index = first_cell_row_index + num_rows + 2
    for sum_row in ws.iter_rows(min_row = sum_row_index, max_row = sum_row_index + 1, max_col = num_col):
        for cell in sum_row:
            fix_sum_row_cells(ws, cell, num_rows)

    # TODO
    # Set the new print area
    # ws.print_area = 'A1:F10'

    # Save in new project directory
    # TODO add path to bid files, handle error if path not found
    file_name = f"Takeoff - {project_name}.xlsx"
    wb.save(file_name)


if __name__ == "__main__":
    # main()
    create_new_takeoff("BASE Takeoff.xlsx", "New Project 1", 13, False)