import os
import unittest

from openpyxl import load_workbook
from dotenv import load_dotenv

import takeoff

class TestHelperFunctions(unittest.TestCase):
    def setUp(self):
        self.test_dir = os.getenv("TEST_DIR")
        self.test_file = f"\\{os.getenv('TEST_FILE')}"
        self.check_file = f"\\{os.getenv('CHECK_FILE')}"
        self.num_rows = int(os.getenv("NUM_ROWS"))
        self.project_name = "Unit_Test"
        if os.getenv("DRILLED") == "True":
            self.drilled = True
        else:
            self.drilled = False

    def tearDown(self):
        pass

    def test_create_new_takeoff(self):

        takeoff.create_new_takeoff(self.test_file, self.project_name, self.num_rows, self.drilled, self.test_dir, self.test_dir)

        test_wb = load_workbook(filename = f"{self.test_dir}//Takeoff - Unit_Test.xlsx")
        test_ws = test_wb["Takeoff-SB"]
        test_cells = tuple(test_ws)

        check_wb = load_workbook(filename = f"{self.test_dir}//{self.check_file}")
        check_ws = check_wb["Takeoff-SB"]
        check_cells = tuple(check_ws)

        # Check num rows
        self.assertEqual(len(test_cells), len(check_cells))

        # Check num columns
        self.assertEqual(len(test_cells[0]), len(check_cells[0]))

        # Check print area
        self.assertEqual(test_ws.print_area[0], check_ws.print_area[0])

        # Check each cell
        style_list = ["alignment", "border", "fill", "font", "number_format", "protection", "quotePrefix"]
        for test_row, check_row in zip(test_cells, check_cells):
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
