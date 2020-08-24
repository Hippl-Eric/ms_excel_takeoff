from openpyxl.formula.translate import Translator
from openpyxl.formula import Tokenizer

from copy import copy

def cell_search(cells, value):
    """Return cell location if equal to value"""
    for row in cells:
        for cell in row:
            if cell.value == value:
                return cell
    raise AssertionError(f"value: {value} not found") 

def copy_row(base_row, int_count):
    """Copy all attributes of the [base_row] and paste [int_count] number of rows below the [base_row]
    
    [base_row] must be a tuple containing a single row cell range.
    [int_count] must be a positive integer.
    """

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

    # Translate the formula (if present)
    try:
        cell.value = Translator(cell.value, origin=old_cell.coordinate).translate_formula(cell.coordinate)
        
        # Check for formuals with ranges
        if ":" in cell.value:
            correct_range_row(cell.value, int_count)

    # Skip cells with no formulas
    except TypeError:
        pass

def correct_range_row(formula, int_count):
    """Recursive helper function for correcting formulas that include ranges.
    The built in openpyxl formula Translator does not handle ranges with rows inserted."""
    colon_index = formula.find(":")
