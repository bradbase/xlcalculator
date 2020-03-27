
# Excel reference: https://support.office.com/en-us/article/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Mid
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class Test_Mid(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/MID.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_start_num_must_be_integer(self):
        with self.assertRaises(ExcelError):
            Mid.mid('Romain', 1.1, 2)


    def test_num_chars_must_be_integer(self):
        with self.assertRaises(ExcelError):
            Mid.mid('Romain', 1, 2.1)


    def test_start_num_must_be_superior_or_equal_to_1(self):
        with self.assertRaises(ExcelError):
            Mid.mid('Romain', 0, 3)


    def test_num_chars_must_be_positive(self):
        with self.assertRaises(ExcelError):
            Mid.mid('Romain', 1, -1)


    def test_mid(self):
        self.assertEqual(Mid.mid('Romain', 3, 4), 'main')
        self.assertEqual(Mid.mid('Romain', 1, 2), 'Ro')
        self.assertEqual(Mid.mid('Romain', 3, 6), 'main')


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )

    
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
