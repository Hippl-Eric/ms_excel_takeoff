import os
import unittest

from openpyxl import load_workbook
from dotenv import load_dotenv

import takeoff

class TestHelperFunctions(unittest.TestCase):
    def setUp(self):
        self.test_dir = os.getenv("TEST_DIR")

    def tearDown(self):
        pass

    def test_create_new_takeoff(self):
        test_file = "\\BASE Takeoff.xlsx"
        project_name = "Unit_Test"
        num_rows = 6
        drilled = True

        takeoff.create_new_takeoff(test_file, project_name, num_rows, drilled, self.test_dir, self.test_dir)

if __name__ == "__main__":
    load_dotenv()
    unittest.main()
