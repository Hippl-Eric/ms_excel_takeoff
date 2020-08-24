import unittest

from openpyxl import load_workbook

import helpers

class TestHelperFunctions(unittest.TestCase):

    def setUp(self):
        wb = load_workbook(filename="test_workbook.xlsx")
        ws = wb.active
        all_cells = tuple(ws.rows)
        num_col = len(all_cells[0])


    def tearDown(self):
        pass

    def test_cell_search(self):
        pass

    def test_copy_row(self):
        pass

    def test_fix_sum_row_cells(self):
        pass

if __name__ == "__main__":
    unittest.main()
    