
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode
from networkx import info, nodes, edges, is_isomorphic

from xlcalculator.model.model import Model
from xlcalculator.xlcalculator_types import XLCell, XLFormula, XLRange
from xlcalculator.read_excel.tokenizer import f_token

from ..formulas import *
from ..xlcalculator_test import XlCalculatorTestCase
from ..function_library import testing


class TestModel(XlCalculatorTestCase):

    def setUp(self):
        infile = open(testing.get_resource("model.json"), "rb")
        json_bytes = infile.read()
        infile.close()
        data = decode(json_bytes, keys=True, classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        self.model = Model()
        self.model.cells = deepcopy(self.cells)
        self.model.defined_names = deepcopy(self.defined_names)
        self.model.ranges = deepcopy(self.ranges)
        self.model.formulae = deepcopy(self.formulae)


    def test_build_code(self):

        this_model = Model()
        this_model.cells = self.cells
        this_model.defined_names = self.defined_names
        this_model.ranges = self.ranges
        this_model.formulae = self.formulae

        this_model.build_code()

        self.assertEqual(self.model, this_model)

        # TODO: fix test to check that the python code has bee correctly written
        # for cell in this_model.cells:
        #     if this_model.cells[cell].formula is not None:
        #         print( this_model.cells[cell].formula.python_code )


    def test_shunting_yard(self):
        this_model = deepcopy(self.model)

        cell_address = this_model.shunting_yard(formula_cell_address, self.model.defined_names)
        self.assertASTNodesEqual(formula_cell_address.reverse_polish_tokens, cell_address)

        cell_address_with_sheet = this_model.shunting_yard(formula_cell_address_with_sheet, self.model.defined_names)
        self.assertASTNodesEqual(formula_cell_address_with_sheet.reverse_polish_tokens, cell_address_with_sheet)

        cells_union_with_same_sheet = this_model.shunting_yard(formula_cells_union_with_same_sheet, self.model.defined_names)
        self.assertASTNodesEqual(formula_cells_union_with_same_sheet.reverse_polish_tokens, cells_union_with_same_sheet)


    @unittest.skip("I don't know how to easily compare networks. is_isomorphic isn't doing what I though it did.")
    def test_build_ast(self):
        this_model = deepcopy(self.model)

        range_address_function = this_model.shunting_yard(formula_range_address_function, self.model.defined_names)
        ast_graph, stack = this_model.build_ast(range_address_function)
        self.assertTrue( is_isomorphic(formula_range_address_function.ast_graph, ast_graph) )
        self.assertEqual(formula_range_address_function.stack, stack)


    def test_create_node(self):
        this_model = deepcopy(self.model)

        range_node = RangeNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'), None)
        created_range_node = this_model.create_node(f_token(tvalue='A1', ttype='operand', tsubtype='range'))
        self.assertIsInstance(created_range_node, RangeNode)
        self.assertEqual(range_node.__repr__(), created_range_node.__repr__())


    def test_set_value(self):
        this_model = deepcopy(self.model)

        # Cells we know are in the model.

        this_model.set_cell_value('First!A2', 88)
        cell_value_00 = 88
        self.assertEqual(cell_value_00, this_model.cells['First!A2'].value)

        this_model.set_cell_value(XLCell('First!A2', 22), 0.1)
        cell_value_01 = 0.1
        self.assertEqual(cell_value_01, this_model.cells['First!A2'].value)

        this_model.set_cell_value('First!A1', 'Header_1')
        cell_value_00 = 'Header_1'
        self.assertEqual(cell_value_00, this_model.cells['First!A1'].value)

        # Cells we know are not in the model, they get added

        this_model.set_cell_value('Sheet1!A1', 88)
        cell_value_00 = 88
        self.assertEqual(cell_value_00, this_model.cells['Sheet1!A1'].value)

        this_model.set_cell_value(XLCell('Sheet1!A1', 22), 0.1)
        cell_value_01 = 0.1
        self.assertEqual(cell_value_01, this_model.cells['Sheet1!A1'].value)


    def test_get_value(self):
        this_model = deepcopy(self.model)

        # cells we know are in the model
        get_cell_value_00 = this_model.get_cell_value('First!A2')
        cell_value_00 = 0.1
        self.assertEqual(cell_value_00, get_cell_value_00)

        get_cell_value_01 = this_model.get_cell_value('First!A1')
        cell_value_01 = 'Header 1'
        self.assertEqual(cell_value_01, get_cell_value_01)

        get_cell_value_02 = this_model.get_cell_value(XLCell('First!A1', 22))
        cell_value_02 = 'Header 1'
        self.assertEqual(cell_value_02, get_cell_value_02)

        get_cell_value_03 = this_model.get_cell_value(XLCell('First!A2', 22))
        cell_value_03 = 0.1
        self.assertEqual(cell_value_03, get_cell_value_03)

        # Cells we know are not in the model, all default to the value 0

        get_cell_value_04 = this_model.get_cell_value('Sheet1!A1')
        cell_value_04 = 0
        self.assertEqual(cell_value_04, get_cell_value_04)






    # def teardown(self):
    #     pass
