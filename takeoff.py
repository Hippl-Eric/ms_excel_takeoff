from openpyxl import load_workbook

def main():

    # Request and parse arguments
    project_name = "test project name 123"
    num_rows = 25
    drilled = False

    # Load base template file
    wb = load_workbook(filename = 'BASE Takeoff.xlsx')
    ws = wb["Takeoff-SB"]

    name = ws['A1']

    ws['A1'] = project_name

    file_name = f"Takeoff - {project_name}.xlsx"

    wb.save(file_name)

    # Input project name

    # Input drilled or driven

    # Insert rows

    # Save in new project directory

    return print("Success")

if __name__ == "__main__":
    main()