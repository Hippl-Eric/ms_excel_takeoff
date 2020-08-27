import os
import unittest

from openpyxl import load_workbook
from dotenv import load_dotenv

import helpers

class TestHelperFunctions(unittest.TestCase):

    def setUp(self):
        self.directory = os.getenv("TEST_DIR")
        self.wb = load_workbook(filename=f"{self.directory}test_workbook.xlsx")
        self.ws = self.wb.active
        self.all_cells = tuple(self.ws.rows)
        self.num_col = len(self.all_cells[0])

    def tearDown(self):
        del self.wb
        del self.ws
        del self.all_cells
        del self.num_col

    def test_cell_search(self):
        result_1 = self.ws['A1']
        result_2 = self.ws['B1']
        result_3 = self.ws['C1']
        result_4 = self.ws['C5']
        self.assertEqual(helpers.cell_search(self.all_cells, 'TEXT'), result_1)
        self.assertEqual(helpers.cell_search(self.all_cells, 'more Text'), result_2)
        self.assertEqual(helpers.cell_search(self.all_cells, '=not_a_formula'), result_3)
        self.assertEqual(helpers.cell_search(self.all_cells, 'SOME_TEXT'), result_4)
        with self.assertRaises(AssertionError):
            helpers.cell_search(self.all_cells, "foobar")

    def test_copy_row(self):
        base_row_idx = 3
        num_rows = 5
        for base_row in self.ws.iter_rows(min_row=base_row_idx, max_row=base_row_idx, max_col=self.num_col):
            helpers.copy_row(self.ws, base_row, num_rows)
        self.wb.save(f'{self.directory}result_copy_row.xlsx')


    def test_fix_sum_row_cells(self):
        sum_row_idx = 16
        num_rows = 5
        for sum_row in self.ws.iter_rows(min_row=sum_row_idx, max_row=sum_row_idx, max_col=self.num_col):
            for cell in sum_row:
                helpers.fix_sum_row_cells(self.ws, cell, num_rows)
        self.wb.save(f'{self.directory}result_fix_sum_row_cells.xlsx')

if __name__ == "__main__":
    load_dotenv()
    unittest.main()
    