
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode
from networkx import info, nodes, edges, is_isomorphic

from xlcalculator.model.model import Model
from xlcalculator.types import XLCell, XLFormula, XLRange
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
