import os

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.formula import Tokenizer
from dotenv import load_dotenv
from copy import copy

def main():

    # Request and parse user inputs
    project_name = input("Project Name: ")
    num_rows = int(input("Number of SBs: ")) - 1
    drilled = True
    install = input("Drilled? (Y/N): ")
    if install == "no" or install == "No" or install == "n" or install == "N" or install == "NO":
        drilled = False

    create_new_takeoff(project_name, num_rows, drilled)

    # Success
    return print("Success")

def create_new_takeoff(project_name: str, num_rows: int, drilled: bool):

    # Load base template file
    load_dotenv()
    init_directory = os.getenv("TEMPLATE_DIR") # template location
    dest_directory = os.getenv("BID_DIR") # bids location
    wb = load_workbook(filename = f"{init_directory}\BASE Takeoff.xlsx")
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

    # Insert new rows below first cell row
    ws.insert_rows(first_cell_row_index + 1, num_rows)

    # Copy the first row values and styling, and paste on all added rows
    for first_cell_row in ws.iter_rows(min_row = first_cell_row_index, max_row = first_cell_row_index, max_col = num_col):
        copy_row(first_cell_row, num_rows)

    # TODO change value for # SB column C

    # TODO fix sum row formulas
    sum_row_index = first_cell_row_index + num_rows + 2
    for row in ws.iter_rows(min_row = sum_row_index, max_row = sum_row_index + 1, max_col = num_col):
        for cell in row:
            fix_sum_row_cells(cell, num_rows)


    # TODO
    # Set the new print area
    # ws.print_area = 'A1:F10'

    # Save in new project directory
    # TODO add path to bid files, handle error if path not found
    file_name = f"Takeoff - {project_name}.xlsx"
    wb.save(file_name)

def cell_search(cells, value):
    """Return cell location if equal to value"""
    for row in cells:
        for cell in row:
            if cell.value == value:
                # TODO should this return a copy of cell? (may be pointing to all_cells tuple)
                return cell
    raise AssertionError(f"value: {value} not found") 

def copy_row(base_row, int_count):
    """Copy all attributes of the [base_row] and paste [int_count] number of rows below the [base_row]"""

    # Create list of style attributes to check later
    # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.styleable.html#openpyxl.styles.styleable.StyleableObject
    style_list = ["alignment", "border", "fill", "font", "number_format", "protection", "quotePrefix"]

    # Copy cells from top row to all inserted rows
    for n in range(1, int_count + 1):
        for cell in base_row:

            # Select the new cell
            new_cell = cell.offset(row=n, column=0)

            # Translate the formula (if present) or just the value
            try:
                new_cell.value = Translator(cell.value, origin=cell.coordinate).translate_formula(new_cell.coordinate)
            except TypeError:
                new_cell.value = cell.value

            # Copy any styling
            if cell.has_style:
                for style in style_list:
                    setattr(new_cell, style, copy(getattr(cell, style)))

def fix_sum_row_cells(cell, int_count):


    # Select the cells old location
    old_cell = cell.offset(row=(-int_count), column=0)

    # Translate the formula (if present) or just the value
    try:
        cell.value = Translator(cell.value, origin=old_cell.coordinate).translate_formula(cell.coordinate)
        # check for ranges
        # tok = Tokenizer(cell.value)
        # formula = tok.formula
    except TypeError:
        pass

if __name__ == "__main__":
    # main()
    create_new_takeoff("New Project 1", 13, False)