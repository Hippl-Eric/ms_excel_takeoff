import unittest

from openpyxl import load_workbook

import helpers

class TestHelperFunctions(unittest.TestCase):

    def setUp(self):
        self.wb = load_workbook(filename="test_workbook.xlsx")
        self.ws = self.wb.active
        self.all_cells = tuple(self.ws.rows)
        self.num_col = len(self.all_cells[0])


    def tearDown(self):
        pass

    def test_cell_search(self):
        result_1 = self.ws['A1']
        result_2 = self.ws['B1']
        result_3 = self.ws['C1']
        result_4 = self.ws['C5']
        self.assertEqual(helpers.cell_search(self.all_cells, 'TEXT'), result_1)
        self.assertEqual(helpers.cell_search(self.all_cells, 'more Text'), result_2)
        self.assertEqual(helpers.cell_search(self.all_cells, '=not_a_formula'), result_2)
        self.assertEqual(helpers.cell_search(self.all_cells, 'SOME_TEXT'), result_4)

    def test_copy_row(self):
        base_row = self.ws['A3:D3']
        int_count = 3
        

    def test_fix_sum_row_cells(self):
        pass

if __name__ == "__main__":
    unittest.main()
    