
import unittest

import pandas as pd

from koala_xlcalculator.function_library import Choose
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator

class TestChoose(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/CHOOSE.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_choose(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        range_01 = pd.DataFrame([[2, 1],[3, 4]])
        range_02 = pd.DataFrame([[1, 2],[4, 3]])
        choose_result_00 = Choose.choose('2', range_00, range_01, range_02)
        result_00 = pd.DataFrame([[2, 1],[3, 4]])
        self.assertTrue(result_00.equals(choose_result_00))


    def test_choose_evaluation_int(self):
        value_00 = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( 12, value_00 )


    def test_choose_evaluation_cell(self):
        value_01 = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( 4, value_01 )


    @unittest.skip("Range isn't being tokenised properly in choose.")
    def test_choose_evaluation_range(self):
        value_00 = self.evaluator.evaluate('Sheet1!A3')
        result_00 = pd.DataFrame([1, 2, 3])
        self.assertTrue( result_00.equals(value_00) )
