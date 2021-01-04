from openpyxl.formula.translate import Translator
from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import coordinate_from_string

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
        
        # Correct formulas in base row with coordinate or range greater than base row index prior to copy
        correct_formula(cell=cell, start_row_idx=base_row[0].row, int_count=int_count)

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

def correct_formula(cell, start_row_idx, int_count):
    """
    Correct the [cell]'s formula after [int_count] number of rows inserted below [start_row_idx]
    Mimics Excel's built in functionality for a cell's formula to be relative (not absolute)
    """
    # Parse only cells with formulas
    try:
        formula = cell.value.startswith("=")
    except AttributeError:
        pass
    else:
        if formula:
            
            # Tokenize formula and make list of all coordinates and ranges
            formula_str = cell.value
            tok = Tokenizer(formula_str)
            original_vals = [t.value for t in tok.items if t.subtype == 'RANGE']
            
            # Create list for parsed values to be added to
            parsed_vals = []
            
            # Check all original coordinate and range values
            for orig_cell_range in original_vals:
                
                # If coorindate
                if ":" not in orig_cell_range:
                    
                    # Parse
                    parsed_cell_range = correct_row_index(orig_cell_range, start_row_idx, int_count)
                
                # If range
                else:
                    
                    # Split coordinates at ":"
                    colon_idx = orig_cell_range.find(":")
                    first_coordinate = orig_cell_range[0:colon_idx]
                    second_coordinate = orig_cell_range[colon_idx+1:]
                    
                    # Send each coordinate to parse
                    first_parsed = correct_row_index(first_coordinate, start_row_idx, int_count)
                    second_parsed = correct_row_index(second_coordinate, start_row_idx, int_count)
                    
                    # Put range back together
                    parsed_cell_range = f"{first_parsed}:{second_parsed}"
                    
                # Add parsed values to the parsed list (may be unchanged)
                parsed_vals.append(parsed_cell_range)
                
            # Replace all original coordinates of formula with parsed coordinates
            for orig, parsed in zip(original_vals, parsed_vals):
                formula_str = formula_str.replace(orig, parsed, 1)
            
            # Set the cell's value to new formula string
            cell.value = formula_str
    
def correct_row_index(cell_coordinate, start_row_idx, int_count):
    """
    Helper function to correct_formula()
    Determines whether cell's coordinate needs to be corrected (row index > [start_row_idx])
    Adds [int_count] to row index
    """
    
    # Break into column, row
    col, row = coordinate_from_string(cell_coordinate)
    
    # Add int_count to row index
    if row > start_row_idx:
        new_row = row + int_count
        cell_coordinate = cell_coordinate.replace(str(row), str(new_row))
    
    return cell_coordinate
    