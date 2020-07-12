from openpyxl import load_workbook

def main():

    # Set iteration count
    num_rows = 2

    # Load Template File
    wb = load_workbook(filename="input.xlsx")
    ws = wb["Sheet1"]

    # Insert Rows
    ws.insert_rows(2, num_rows)

    for n in range(1, num_rows + 1):
        for row in ws.iter_rows(min_row=1, max_col=20, max_row=1):
            # first_cell = row[0]
            # the_row = first_cell.row
            for cell in row:
                new_cell = ws.cell(row=cell.row + n, column=cell.col_idx, value=cell.value)

                # print(cell)

    file_name = "output.xlsx"
    wb.save(file_name)

    # Success
    return print("Success")

if __name__ == "__main__":
    main()