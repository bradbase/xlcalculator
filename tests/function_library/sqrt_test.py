
# Excel reference: https://support.office.com/en-us/article/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf

import unittest

import pandas as pd

from xlcalculator.function_library import Sqrt
from xlcalculator.exceptions import ExcelError
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator


class Test_Sqrt(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SQRT.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_first_argument_validity(self):
        self.assertIsInstance(Sqrt.sqrt(-16), ExcelError )


    def test_positive_integers(self):
        self.assertEqual(Sqrt.sqrt(16), 4)


    def test_float(self):
        self.assertEqual(Sqrt.sqrt(.25), .5)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_B1(self):
        # only one exception need be raised.
        """
        ======================================================================
        ERROR: test_evaluation_B1 (tests.function_library.sqrt_test.Test_Sqrt)
        ----------------------------------------------------------------------
        Traceback (most recent call last):
          File "C:\\Users\\bradb\\Documents\\Python\\xlcalculator\\xlcalculator\\evaluator\\evaluator.py", line 63, in evaluate
            value = eval(cell.formula.python_code)
          File "<string>", line 1, in <module>
          File "C:\\Users\\bradb\\Documents\\Python\\xlcalculator\\xlcalculator\\function_library\\sqrt.py", line 22, in sqrt
            raise ExcelError('#NUM!', '{} must be non-negative'.format( number ))
        xlcalculator.exceptions.exceptions.ExcelError: #NUM!

        During handling of the above exception, another exception occurred:

        Traceback (most recent call last):
          File "C:\\Users\\bradb\\Documents\\Python\\xlcalculator\\tests\\function_library\\sqrt_test.py", line 44, in test_evaluation_B1
            self.evaluator.evaluate('Sheet1!B1')
          File "C:\\Users\\bradb\\Documents\\Python\\xlcalculator\\xlcalculator\\evaluator\\evaluator.py", line 80, in evaluate
            raise Exception("Problem evalling: {} for {}, {}".format(exception, cell.address, cell.formula.python_code))
        Exception: Problem evalling: #NUM! for Sheet1!B1, Sqrt.sqrt(self.eval_ref("Sheet1!A2"))
        """
        # with self.assertRaises(ExcelError):
        #     self.evaluator.evaluate('Sheet1!B1')

        with self.assertRaises(Exception) as context:
            self.evaluator.evaluate('Sheet1!B1')
            self.assertTrue('#NUM!' in context.exception)


    @unittest.skip("Can't work as we have not implemented function ABS")
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
