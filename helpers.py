from openpyxl.formula.translate import Translator
from openpyxl.formula import Tokenizer

from copy import copy

def cell_search(cells, value):
    """Return cell location if equal to value
    
    [cells] lump tuple of all cells.
    [value] value of any type to be checked.
    """
    for row in cells:
        for cell in row:
            if cell.value == value:
                return cell
    raise AssertionError(f"value: {value} not found") 

def copy_row(work_sheet, base_row, int_count):
    """Copy all attributes of the [base_row] and paste [int_count] number of rows below the [base_row]
    
    [work_sheet] active worksheet.
    [base_row] must be a tuple containing a single row cell range.
    [int_count] must be a positive integer.
    """

    # Create list of style attributes to check later
    # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.styleable.html#openpyxl.styles.styleable.StyleableObject
    style_list = ["alignment", "border", "fill", "font", "number_format", "protection", "quotePrefix"]

    # Insert rows below the base_row
    work_sheet.insert_rows(base_row[0].row + 1, int_count)

    # Iterate over each cell in the base row
    for cell in base_row:

        # Copy cell attributes and paste to all inserted rows
        for n in range(1, int_count + 1):

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

def fix_sum_row_cells(work_sheet, cell, int_count):
    """Correct formulas for the summation row below inserted rows.
    No modifications to styling neccesary.
    
    [work_sheet] active worksheet.
    [cell] single cell.
    [int_count] must be a positive integer.
    """

    # Select the cells old location
    old_cell = cell.offset(row=(-int_count), column=0)

    # Translate the formula (if present)
    try:
        cell.value = Translator(cell.value, origin=old_cell.coordinate).translate_formula(cell.coordinate)
        
        # Correct formulas with ranges
        cell.value = correct_range_row(work_sheet, cell.value, int_count)

    # Skip cells with no formulas
    except TypeError:
        pass

def correct_range_row(work_sheet, formula_string, int_count):
    """Recursive helper function for correcting formulas that include ranges.
    Built in openpyxl formula Translator does not handle ranges with rows inserted.
    
    [work_sheet] active worksheet
    [formula_string] string cell.value.
    [int_count] must be a positive integer.
    """

    if ":" not in formula_string:
        return formula_string

    else:
        # Find first range coordinate before ":"
        colon_idx = formula_string.find(":")
        paren_idx = formula_string.rfind("(", 0, colon_idx)
        coordinate = formula_string[paren_idx + 1: colon_idx]

        # Convert coordinate to cell
        cell_obj = work_sheet[coordinate]

        # Use offset to correct cell coordinate
        new_cell_obj = cell_obj.offset(row=(-int_count), column=0)

        # Return the cells new coordinate to string
        new_coordinate = new_cell_obj.coordinate

        # Put formula string back together
        left_half = formula_string[0: paren_idx + 1] + new_coordinate + formula_string[colon_idx: colon_idx + 1]
        right_half = correct_range_row(work_sheet, formula_string[colon_idx + 1:], int_count)
        new_formula = left_half + right_half

        return new_formula
