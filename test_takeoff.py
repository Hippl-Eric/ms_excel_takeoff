import os
import unittest

from openpyxl import load_workbook
from dotenv import load_dotenv

import takeoff

class TestHelperFunctions(unittest.TestCase):
    def setUp(self):

        # Grab and set inputs
        test_dir = os.getenv("TEST_DIR")
        test_file = f"\\{os.getenv('TEST_FILE')}"
        check_file = f"\\{os.getenv('CHECK_FILE')}"
        num_rows = int(os.getenv("NUM_ROWS"))
        project_name = "Unit_Test"
        if os.getenv("DRILLED") == "True":
            drilled = True
        else:
            drilled = False
        
        # Create a takeoff (excel) file to test
        takeoff.create_new_takeoff(test_file, project_name, num_rows, drilled, test_dir, test_dir)

        # Load the test workbook
        self.test_wb = load_workbook(filename = f"{test_dir}//Takeoff - Unit_Test.xlsx")
        self.test_ws = self.test_wb["Takeoff-SB"]
        self.test_cells = tuple(self.test_ws)

        # Load the check workbook
        self.check_wb = load_workbook(filename = f"{test_dir}//{check_file}")
        self.check_ws = self.check_wb["Takeoff-SB"]
        self.check_cells = tuple(self.check_ws)

    def tearDown(self):
        pass

    def test_create_new_takeoff(self):

        # Check num rows
        self.assertEqual(len(self.test_cells), len(self.check_cells))

        # Check num columns
        self.assertEqual(len(self.test_cells[0]), len(self.check_cells[0]))

        # Check print area
        self.assertEqual(self.test_ws.print_area[0], self.check_ws.print_area[0])

        # Check each cell
        style_list = ["alignment", "border", "fill", "font", "number_format", "protection", "quotePrefix"]
        for test_row, check_row in zip(self.test_cells, self.check_cells):
            for test_cell, check_cell in zip(test_row, check_row):

                # Check value, coordinate, and has_style
                self.assertEqual(test_cell.value, check_cell.value)
                self.assertEqual(test_cell.coordinate, check_cell.coordinate)
                self.assertEqual(test_cell.has_style, check_cell.has_style)

                # Check Style
                # TODO                            

if __name__ == "__main__":
    load_dotenv()
    unittest.main()
