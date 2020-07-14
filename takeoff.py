from openpyxl import load_workbook

def main():

    # Request and parse arguments
    project_name = "test project name 123"
    num_rows = 10
    drilled = False

    # Load base template file
    wb = load_workbook(filename = 'BASE Takeoff.xlsx')
    ws = wb["Takeoff-SB"]

    # Input project name
    ws['A1'] = project_name

    # Input drilled or driven
    if drilled:
        ws['F2'] = "DRILLED"
    else:
        ws["F2"] = "DRIVEN"

    # Insert rows
    ws.insert_rows(6, num_rows)


    # TODO
    # Set the new print area
    # ws.print_area = 'A1:F10'

    # Save in new project directory
    file_name = f"Takeoff - {project_name}.xlsx"
    wb.save(file_name)

    # Success
    return print("Success")

if __name__ == "__main__":
    main()