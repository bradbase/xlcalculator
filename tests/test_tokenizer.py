from xlcalculator.tokenizer import f_token, ExcelParser

from . import testing


class ExcelParserTest(testing.XlCalculatorTestCase):

    def setUp(self):
        self.operand_range_00 = ['A1:B1']
        self.operand_range_01 = ['A1', 'B1']

    def parse(self, formula):
        parser = ExcelParser()
        return parser.parse(formula).items

    def test_cell_minus(self):
        self.assertASTNodesEqual(
            self.parse('-A1'),
            [
                f_token(tvalue='-', ttype='operator-prefix', tsubtype=''),
                f_token(tvalue='A1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_minus(self):
        self.assertASTNodesEqual(
            self.parse('-A1,-B1'),
            [
                f_token(tvalue='-', ttype='operator-prefix', tsubtype=''),
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='operator-infix', tsubtype='union'),
                f_token(tvalue='-', ttype='operator-prefix', tsubtype=''),
                f_token(tvalue='B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cell_minus_with_sheet(self):
        self.assertASTNodesEqual(
            self.parse('-Sheet1!A1'),
            [
                f_token(tvalue='-', ttype='operator-prefix', tsubtype=''),
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cell_address(self):
        self.assertASTNodesEqual(
            self.parse('A1'),
            [
                f_token(tvalue='A1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cell_address_with_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_union(self):
        self.assertASTNodesEqual(
            self.parse('A1,B1'),
            [
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='operator-infix', tsubtype='union'),
                f_token(tvalue='B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_union_with_same_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1,Sheet1!B1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='operator-infix', tsubtype='union'),
                f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_union_with_different_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1,Sheet2!B1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='operator-infix', tsubtype='union'),
                f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_addition(self):
        self.assertASTNodesEqual(
            self.parse('A1+B1'),
            [
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='+', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_addition_with_same_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1+Sheet1!B1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='+', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_addition_with_different_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1+Sheet2!B1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='+', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_is_equal(self):
        self.assertASTNodesEqual(
            self.parse('A1=B1'),
            [
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='=', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_is_equal_with_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sheet1!A1=Sheet1!B1'),
            [
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='=', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_cells_function(self):
        self.assertASTNodesEqual(
            self.parse('SUM(A1,B1)'),
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='argument', tsubtype=''),
                f_token(tvalue='B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )

    def test_cells_function_with_sheet(self):
        self.assertASTNodesEqual(
            self.parse('SUM(Sheet1!A1,Sheet1!B1)'),
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='argument', tsubtype=''),
                f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )

    def test_cells_function_and_operator(self):
        self.assertASTNodesEqual(
            self.parse('SUM(A1+B1,C1)'),
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1', ttype='operand', tsubtype='range'),
                f_token(tvalue='+', ttype='operator-infix', tsubtype='math'),
                f_token(tvalue='B1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='argument', tsubtype=''),
                f_token(tvalue='C1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )

    def test_range_address(self):
        self.assertASTNodesEqual(
            self.parse('A1:B1'),
            [
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_range_address_with_sheet(self):
        self.assertASTNodesEqual(
            self.parse('Sh1!A1:B1'),
            [
                f_token(tvalue='Sh1!A1:B1', ttype='operand', tsubtype='range')
            ]
        )

    def test_range_address_with_same_sheets(self):
        self.assertASTNodesEqual(
            self.parse('Sh1!A1:Sh1!B1'),
            [
                f_token(tvalue='Sh1!A1:Sh1!B1',
                        ttype='operand', tsubtype='range')
            ]
        )

    def test_range_address_with_different_sheets(self):
        # THIS CAN'T HAPPEN IN REAL LIFE
        self.assertASTNodesEqual(
            self.parse('Sh1!A1:Sh2!B1'),
            [
                f_token(tvalue='Sh1!A1:Sh2!B1',
                        ttype='operand', tsubtype='range')
            ]
        )

    def test_range_address_union(self):
        self.assertASTNodesEqual(
            self.parse('A1:B1,A2:B2'),
            [
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
                f_token(tvalue=',', ttype='operator-infix', tsubtype='union'),
                f_token(tvalue='A2:B2', ttype='operand', tsubtype='range')
            ]
        )

    def test_range_address_function(self):
        self.assertASTNodesEqual(
            self.parse('SUM(A1:B1)'),
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop'),
            ]
        )
