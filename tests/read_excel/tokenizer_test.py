
import unittest

from koala_xlcalculator.read_excel.tokenizer import ExcelParser
from koala_xlcalculator.read_excel.tokenizer import f_token


class TestExcelParser(unittest.TestCase):

    def setUp(self):
        self.formula_000 = [f_token(tvalue='A1', ttype='operand', tsubtype='range')]
        self.formula_001 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')]
        self.formula_002 = [f_token(tvalue='A1:B1', ttype='operand', tsubtype='range')]
        self.formula_003 = [f_token(tvalue='Sheet1!A1:B1', ttype='operand', tsubtype='range')]
        self.formula_004 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_005 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_006 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_007 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_008 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_009 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_010 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_011 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_012 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_013 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_014 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_015 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_016 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_017 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_018 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_019 = [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range')]
        self.formula_020 = [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_021 = [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')]
        self.formula_022 = [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='B1', ttype='operand', tsubtype='range')]
        self.formula_023 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_024 = [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_025 = [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')]
        self.formula_026 = [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')]
        self.formula_027 = [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')]
        self.formula_028 = [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='C1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')]
        self.formula_029 = [f_token(tvalue='Sheet1!A1:Sheet1!B1', ttype='operand', tsubtype='range')]
        self.formula_030 = [f_token(tvalue='Sheet1!A1:Sheet2!B1', ttype='operand', tsubtype='range')]
        self.formula_031 = [f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='A2:B2', ttype='operand', tsubtype='range')]
        self.formula_032 = [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')]


    # def teardown(self):
    #     pass

    def test_cell_minus(self):
        parser = ExcelParser()
        parser.parse("-A1")
        self.assertEqual(self.formula_019, parser.tokens.items)


    def test_cells_minus(self):
        parser = ExcelParser()
        parser.parse("-A1,-B1")
        self.assertEqual(self.formula_020, parser.tokens.items)


    def test_cell_minus_with_sheet(self):
        parser = ExcelParser()
        parser.parse("-Sheet1!A1")
        self.assertEqual(self.formula_021, parser.tokens.items)


    def test_cell_address(self):
        parser = ExcelParser()
        parser.parse("A1")
        self.assertEqual(self.formula_000, parser.tokens.items)


    def test_cell_address_with_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1")
        self.assertEqual(self.formula_001, parser.tokens.items)


    def test_cells_union(self):
        parser = ExcelParser()
        parser.parse("A1,B1")
        self.assertEqual(self.formula_004, parser.tokens.items)


    def test_cells_union_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1,Sheet1!B1")
        self.assertEqual(self.formula_005, parser.tokens.items)


    def test_cells_union_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1,Sheet2!B1")
        self.assertEqual(self.formula_006, parser.tokens.items)


    def test_cells_addition(self):
        parser = ExcelParser()
        parser.parse("A1+B1")
        self.assertEqual(self.formula_007, parser.tokens.items)


    def test_cells_addition_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1+Sheet1!B1")
        self.assertEqual(self.formula_008, parser.tokens.items)


    def test_cells_addition_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1+Sheet2!B1")
        self.assertEqual(self.formula_009, parser.tokens.items)


    def test_cells_subtract(self):
        parser = ExcelParser()
        parser.parse("A1-B1")
        self.assertEqual(self.formula_010, parser.tokens.items)


    def test_cells_subtract_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1-Sheet1!B1")
        self.assertEqual(self.formula_011, parser.tokens.items)


    def test_cells_subtract_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1-Sheet2!B1")
        self.assertEqual(self.formula_012, parser.tokens.items)


    def test_cells_multiply(self):
        parser = ExcelParser()
        parser.parse("A1*B1")
        self.assertEqual(self.formula_013, parser.tokens.items)


    def test_cells_multiply_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1*Sheet1!B1")
        self.assertEqual(self.formula_014, parser.tokens.items)


    def test_cells_multiply_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1*Sheet2!B1")
        self.assertEqual(self.formula_015, parser.tokens.items)


    def test_cells_divide(self):
        parser = ExcelParser()
        parser.parse("A1/B1")
        self.assertEqual(self.formula_016, parser.tokens.items)


    def test_cells_divide_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1/Sheet1!B1")
        self.assertEqual(self.formula_017, parser.tokens.items)


    def test_cells_divide_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1/Sheet2!B1")
        self.assertEqual(self.formula_018, parser.tokens.items)


    def test_cells_is_equal(self):
        parser = ExcelParser()
        parser.parse("A1==B1")
        self.assertEqual(self.formula_022, parser.tokens.items)


    def test_cells_is_equal_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1==Sheet1!B1")
        self.assertEqual(self.formula_023, parser.tokens.items)


    def test_cells_is_equal_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1==Sheet2!B1")
        self.assertEqual(self.formula_024, parser.tokens.items)
        

    def test_cells_function(self):
        parser = ExcelParser()
        parser.parse("SUM(A1,B1)")
        self.assertEqual(self.formula_025, parser.tokens.items)


    def test_cells_function_with_same_sheet(self):
        parser = ExcelParser()
        parser.parse("SUM(Sheet1!A1,Sheet1!B1)")
        self.assertEqual(self.formula_026, parser.tokens.items)


    def test_cells_is_equal_with_different_sheet(self):
        parser = ExcelParser()
        parser.parse("SUM(Sheet1!A1,Sheet2!B1)")
        self.assertEqual(self.formula_027, parser.tokens.items)


    def test_cells_function_and_operator(self):
        parser = ExcelParser()
        parser.parse("SUM(A1+B1,C1)")
        self.assertEqual(self.formula_028, parser.tokens.items)


    def test_range_address(self):
        parser = ExcelParser()
        parser.parse("A1:B1")
        self.assertEqual(self.formula_002, parser.tokens.items)


    def test_range_address_with_sheet(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1:B1")
        self.assertEqual(self.formula_003, parser.tokens.items)


    def test_range_address_with_same_sheets(self):
        parser = ExcelParser()
        parser.parse("Sheet1!A1:Sheet1!B1")
        self.assertEqual(self.formula_029, parser.tokens.items)


    def test_range_address_with_different_sheets(self):
        # THIS CAN'T HAPPEN IN REAL LIFE
        parser = ExcelParser()
        parser.parse("Sheet1!A1:Sheet2!B1")
        self.assertEqual(self.formula_030, parser.tokens.items)


    def test_range_address_union(self):
        parser = ExcelParser()
        parser.parse("A1:B1,A2:B2")
        self.assertEqual(self.formula_031, parser.tokens.items)


    def test_range_address_function(self):
        parser = ExcelParser()
        parser.parse("SUM(A1:B1)")
        self.assertEqual(self.formula_032, parser.tokens.items)
