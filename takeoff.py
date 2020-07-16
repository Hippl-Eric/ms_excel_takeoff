from openpyxl import load_workbook

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

def create_new_takeoff(project_name, num_rows, drilled):

    # Load base template file
    # TODO load values from .env file
    init_directory = "" # template location
    dest_directory = "" # bids location
    wb = load_workbook(filename = "BASE Takeoff.xlsx")
    ws = wb["Takeoff-SB"]

    # Load cells
    all_cells = []

    # Search for project name, drilled/driven, and first row, return locations
    name_cell = cell_search(all_cells, "PROJECTNAME -  TAKEOFF")
    drill_cell = cell_search(all_cells, "DRIVEN/DRILLED")
    first_row_cell = cell_search(all_cells, "SB Nos.")

    # Input project name
    name_cell.value = project_name

    # Input drilled or driven
    if drilled:
        drill_cell.value = "DRILLED"
    else:
        drill_cell.value = "DRIVEN"

    # Insert rows below first row
    copy_insert()

    # TODO
    # Set the new print area
    # ws.print_area = 'A1:F10'

    # Save in new project directory
    # TODO add path to bid files, handle error if path not found
    file_name = f"Takeoff - {project_name}.xlsx"
    wb.save(file_name)

def cell_search(cells, value):
    pass

def copy_insert():
    pass

if __name__ == "__main__":
    main()