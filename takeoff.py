import os

from openpyxl import load_workbook
from dotenv import load_dotenv

from helpers import cell_search, copy_row, fix_sum_row_cells

def main():

    # Request and parse user inputs
    # TODO parse user inputs
    project_name = input("Project Name: ")
    num_rows = int(input("Number of SBs: ")) - 1
    drilled = True
    install = input("Drilled? (Y/N): ")
    if install == "no" or install == "No" or install == "n" or install == "N" or install == "NO":
        drilled = False

    # Load directory locations and template file
    load_dotenv()
    template_dir = os.getenv("TEMPLATE_DIR") # template location
    dest_dir = os.getenv("BID_DIR") # bids location
    template_file = os.getenv("TEMPLATE_FILE")

    create_new_takeoff(template_file, project_name, num_rows, drilled, template_dir, dest_dir)

    # Success
    return print("Success")

def create_new_takeoff(template_file, project_name, num_rows, drilled, temp_dir, dest_dir):
    """Description

    [template_file] string, filename
    [project_name] string
    [num_rows] int
    [drilled] bool
    [temp_dir] string, directory of template file
    [dest_dir] string, directory to save new file
    """

    # Load base template file
    wb = load_workbook(filename = f"{temp_dir}\\{template_file}")
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

    # Correct forumlas in sum row
    sum_row_index = first_cell_row_index + num_rows + 2
    for sum_row in ws.iter_rows(min_row = sum_row_index, max_row = sum_row_index + 1, max_col = num_col):
        for cell in sum_row:
            fix_sum_row_cells(ws, cell, num_rows)
            
    # Change value for "SB Nos."
    sb_column = title_row_cell.column
    start_sb = 1
    for row_num in range(first_cell_row_index, first_cell_row_index + num_rows + 1):
        ws.cell(row=row_num, column=sb_column, value=str(start_sb))
        start_sb += 1

    # Correct the print area
    old_area = ws.print_area[0]
    colon_idx = old_area.find(":")
    end_coordinate = old_area[colon_idx + 1:]

    # Convert end_coordinate to cell, offset, return to coordinate
    end_coordinate_cell_obj = ws[end_coordinate]
    end_coordinate_new_cell_obj = end_coordinate_cell_obj.offset(row=num_rows, column=0)
    end_coordinate_new = end_coordinate_new_cell_obj.coordinate

    # Set print area
    ws.print_area = f"{old_area[:colon_idx + 1]}{end_coordinate_new}"

    # Save workbook in new project directory
    file_name = f"{dest_dir}\\{project_name}\\PRICING\\Takeoff - {project_name}.xlsx"
    while file_name != f"quit\\Takeoff - {project_name}.xlsx":
        try:
            wb.save(file_name)
            break
        except FileNotFoundError:
            file_path = input("Bid file location not found. Specifiy path to save file, or 'quit':  ")
            file_name = f"{file_path}\\Takeoff - {project_name}.xlsx"

if __name__ == "__main__":
    main()
