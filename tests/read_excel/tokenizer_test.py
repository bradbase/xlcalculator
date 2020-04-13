
import unittest

from xlcalculator.read_excel.tokenizer import ExcelParser

from ..formulas import *
from ..xlcalculator_test import XlCalculatorTestCase


class TestExcelParser(XlCalculatorTestCase):

    def setUp(self):

        self.operand_range_00 = ['A1:B1']
        self.operand_range_01 = ['A1', 'B1']


    # def teardown(self):
    #     pass

    def test_cell_minus(self):
        parser = ExcelParser()
        parser.parse('-A1')
        self.assertASTNodesEqual(formula_cell_minus.tokens, parser.tokens.items)


    def test_cells_minus(self):
        parser = ExcelParser()
        parser.parse('-A1,-B1')
        self.assertASTNodesEqual(formula_cells_minus.tokens, parser.tokens.items)


    def test_cell_minus_with_sheet(self):
        parser = ExcelParser()
        parser.parse('-Sheet1!A1')
        self.assertASTNodesEqual(formula_cell_minus_with_sheet.tokens, parser.tokens.items)


    def test_cell_address(self):
        parser = ExcelParser()
        parser.parse('A1')
        self.assertASTNodesEqual(formula_cell_address.tokens, parser.tokens.items)


    def test_cell_address_with_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1')
        self.assertASTNodesEqual(formula_cell_address_with_sheet.tokens, parser.tokens.items)


    def test_cells_union(self):
        parser = ExcelParser()
        parser.parse('A1,B1')
        self.assertASTNodesEqual(formula_cells_union.tokens, parser.tokens.items)


    def test_cells_union_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1,Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_union_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_union_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1,Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_union_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_addition(self):
        parser = ExcelParser()
        parser.parse('A1+B1')
        self.assertASTNodesEqual(formula_cells_addition.tokens, parser.tokens.items)


    def test_cells_addition_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1+Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_addition_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_addition_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1+Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_addition_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_subtract(self):
        parser = ExcelParser()
        parser.parse('A1-B1')
        self.assertASTNodesEqual(formula_cells_subtract.tokens, parser.tokens.items)


    def test_cells_subtract_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1-Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_subtract_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_subtract_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1-Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_subtract_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_multiply(self):
        parser = ExcelParser()
        parser.parse('A1*B1')
        self.assertASTNodesEqual(formula_cells_multiply.tokens, parser.tokens.items)


    def test_cells_multiply_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1*Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_multiply_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_multiply_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1*Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_multiply_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_divide(self):
        parser = ExcelParser()
        parser.parse('A1/B1')
        self.assertASTNodesEqual(formula_cells_divide.tokens, parser.tokens.items)


    def test_cells_divide_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1/Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_divide_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_divide_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1/Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_divide_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_is_equal(self):
        parser = ExcelParser()
        parser.parse('A1==B1')
        self.assertASTNodesEqual(formula_cells_is_equal.tokens, parser.tokens.items)


    def test_cells_is_equal_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1==Sheet1!B1')
        self.assertASTNodesEqual(formula_cells_is_equal_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_is_equal_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1==Sheet2!B1')
        self.assertASTNodesEqual(formula_cells_is_equal_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_function(self):
        parser = ExcelParser()
        parser.parse('SUM(A1,B1)')
        self.assertASTNodesEqual(formula_cells_function.tokens, parser.tokens.items)


    def test_cells_function_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse('SUM(Sheet1!A1,Sheet1!B1)')
        self.assertASTNodesEqual(formula_cells_function_with_same_sheet.tokens, parser.tokens.items)


    def test_cells_is_equal_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse('SUM(Sheet1!A1,Sheet2!B1)')
        self.assertASTNodesEqual(formula_cells_is_equal_with_different_sheet.tokens, parser.tokens.items)


    def test_cells_function_and_operator(self):
        parser = ExcelParser()
        parser.parse('SUM(A1+B1,C1)')
        self.assertASTNodesEqual(formula_cells_function_and_operator.tokens, parser.tokens.items)


    def test_range_address(self):
        parser = ExcelParser()
        parser.parse('A1:B1')
        self.assertASTNodesEqual(formula_range_address.tokens, parser.tokens.items)


    def test_range_address_with_sheet(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1:B1')
        self.assertASTNodesEqual(formula_range_address_with_sheet.tokens, parser.tokens.items)


    def test_range_address_with_same_sheets(self):
        parser = ExcelParser()
        parser.parse('Sheet1!A1:Sheet1!B1')
        self.assertASTNodesEqual(formula_range_address_with_same_sheets.tokens, parser.tokens.items)


    def test_range_address_with_different_sheets(self):
        # THIS CAN'T HAPPEN IN REAL LIFE
        parser = ExcelParser()
        parser.parse('Sheet1!A1:Sheet2!B1')
        self.assertASTNodesEqual(formula_range_address_with_different_sheets.tokens, parser.tokens.items)


    def test_range_address_union(self):
        parser = ExcelParser()
        parser.parse('A1:B1,A2:B2')
        self.assertASTNodesEqual(formula_range_address_union.tokens, parser.tokens.items)


    def test_range_address_function(self):
        parser = ExcelParser()
        parser.parse('SUM(A1:B1)')
        self.assertASTNodesEqual(formula_range_address_function.tokens, parser.tokens.items)
