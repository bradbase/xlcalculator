import unittest

from xlcalculator import evaluator, model
from . import testing


class TestEvaluator(unittest.TestCase):

    def setUp(self):
        self.model = model.Model()
        self.model.construct_from_json_file(
            testing.get_resource("model.json"), build_code=True)
        self.evaluator = evaluator.Evaluator(self.model)

    def test_evaluate(self):
        evaluated_result_00 = self.evaluator.evaluate('First!A2')
        result_00 = 0.1
        self.assertEqual(result_00, evaluated_result_00)

        evaluated_result_01 = self.evaluator.evaluate('Seventh!C1')
        result_01 = 3
        self.assertEqual(result_01, evaluated_result_01)

        evaluated_result_02 = self.evaluator.evaluate('Ninth!B1')
        result_02 = 101
        self.assertEqual(result_02, evaluated_result_02)

        evaluated_result_03 = self.evaluator.evaluate('Hundred')
        result_03 = 100
        self.assertEqual(result_03, evaluated_result_03)

        evaluated_result_04 = self.evaluator.evaluate('Tenth!C1')
        result_04 = 102
        self.assertEqual(result_04, evaluated_result_04)

        evaluated_result_06 = self.evaluator.evaluate('Tenth!C2')
        result_06 = 102
        self.assertEqual(result_06, evaluated_result_06)

        evaluated_result_07 = self.evaluator.evaluate('Tenth!C3')
        result_07 = 102
        self.assertEqual(result_07, evaluated_result_07)

    def test_set_value(self):
        self.evaluator.set_cell_value('First!A2', 88)
        evaluated_result_00 = self.evaluator.model.cells['First!A2'].value
        result_00 = 88
        self.assertEqual(result_00, evaluated_result_00)
        # Put it back the way we found it.
        self.evaluator.set_cell_value('First!A2', 0.1)

    def test_set_value_evaluate(self):
        self.evaluator.set_cell_value('First!A2', 88)
        evaluated_result_00 = self.evaluator.evaluate('First!A2')
        result_00 = 88
        self.assertEqual(result_00, evaluated_result_00)
        # Put it back the way we found it.
        self.evaluator.set_cell_value('First!A2', 0.1)

    def test_divide_eval(self):
        div_compiler = model.ModelCompiler()
        div_model = div_compiler.read_and_parse_archive(
            testing.get_resource("division.xlsx"))
        div_evaluator = evaluator.Evaluator(div_model)

        excel_value_00 = div_evaluator.get_cell_value('Sheet1!A1')
        value_00 = div_evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value_00, value_00)

        excel_value_01 = div_evaluator.get_cell_value('Sheet1!B1')
        value_01 = div_evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value_01, value_01)

    def test_subtract_eval(self):
        sub_compiler = model.ModelCompiler()
        sub_model = sub_compiler.read_and_parse_archive(
            testing.get_resource("subtraction.xlsx"))
        sub_evaluator = evaluator.Evaluator(sub_model)

        excel_value_00 = sub_evaluator.get_cell_value('Sheet1!A1')
        value_00 = sub_evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value_00, value_00)

        excel_value_01 = sub_evaluator.get_cell_value('Sheet1!B1')
        value_01 = sub_evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value_01, value_01)

        excel_value_02 = sub_evaluator.get_cell_value('Sheet1!C1')
        value_02 = sub_evaluator.evaluate('Sheet1!C1')
        self.assertEqual(excel_value_02, value_02)

    def test_addition_eval(self):
        add_compiler = model.ModelCompiler()
        add_model = add_compiler.read_and_parse_archive(
            testing.get_resource("addition.xlsx"))
        add_evaluator = evaluator.Evaluator(add_model)

        excel_value_00 = add_evaluator.get_cell_value('Sheet1!A1')
        value_00 = add_evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value_00, value_00)

        excel_value_01 = add_evaluator.get_cell_value('Sheet1!B1')
        value_01 = add_evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value_01, value_01)

        excel_value_02 = add_evaluator.get_cell_value('Sheet1!C1')
        value_02 = add_evaluator.evaluate('Sheet1!C1')
        self.assertEqual(excel_value_02, value_02)

    def test_multiplication_eval(self):
        add_compiler = model.ModelCompiler()
        add_model = add_compiler.read_and_parse_archive(
            testing.get_resource("multiplication.xlsx"))
        add_evaluator = evaluator.Evaluator(add_model)

        excel_value_00 = add_evaluator.get_cell_value('Sheet1!A1')
        value_00 = add_evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value_00, value_00)

        excel_value_01 = add_evaluator.get_cell_value('Sheet1!B1')
        value_01 = add_evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value_01, value_01)

        excel_value_02 = add_evaluator.get_cell_value('Sheet1!C1')
        value_02 = add_evaluator.evaluate('Sheet1!C1')
        self.assertEqual(excel_value_02, value_02)

        excel_value_03 = add_evaluator.get_cell_value('Sheet1!D1')
        value_03 = add_evaluator.evaluate('Sheet1!D1')
        self.assertEqual(excel_value_03, value_03)
