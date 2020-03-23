
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode
from networkx import info, nodes, edges, is_isomorphic

from koala_xlcalculator.model.model import Model
from koala_xlcalculator.koala_types import XLCell, XLFormula, XLRange
from koala_xlcalculator.read_excel.tokenizer import f_token

from ..formulas import *


class TestModel(unittest.TestCase):

    def setUp(self):
        infile = open(r"./tests/resources/model.json", "rb")
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


    def test_shunting_yard(self):
        this_model = deepcopy(self.model)

        cell_address = this_model.shunting_yard(formula_cell_address.address, self.model.defined_names)
        self.assertEqual(formula_cell_address.reverse_polish_tokens, cell_address)

        cell_address_with_sheet = this_model.shunting_yard(formula_cell_address_with_sheet.address, self.model.defined_names)
        self.assertEqual(formula_cell_address_with_sheet.reverse_polish_tokens, cell_address_with_sheet)

        cells_union_with_same_sheet = this_model.shunting_yard(formula_cells_union_with_same_sheet.address, self.model.defined_names)
        self.assertEqual(formula_cells_union_with_same_sheet.reverse_polish_tokens, cells_union_with_same_sheet)


    @unittest.skip("I don't know how to easily compare networks. is_isomorphic isn't doing what I though it did.")
    def test_build_ast(self):
        this_model = deepcopy(self.model)

        range_address_function = this_model.shunting_yard(formula_range_address_function.address, self.model.defined_names)
        ast_graph, stack = this_model.build_ast(range_address_function)
        self.assertTrue( is_isomorphic(formula_range_address_function.ast_graph, ast_graph) )
        self.assertEqual(formula_range_address_function.stack, stack)


    def test_create_node(self):
        this_model = deepcopy(self.model)

        range_node = RangeNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'), None)
        created_range_node = this_model.create_node(f_token(tvalue='A1', ttype='operand', tsubtype='range'))
        self.assertIsInstance(created_range_node, RangeNode)
        self.assertEqual(range_node, created_range_node)

    # def teardown(self):
    #     pass
