import unittest

from networkx import info

from xlcalculator.xlcalculator_types import OperatorNode
# from xlcalculator.xlcalculator_types import OperandNode
from xlcalculator.xlcalculator_types import RangeNode
from xlcalculator.xlcalculator_types import FunctionNode
from xlcalculator.read_excel.tokenizer import f_token

from ..formulas import *


class ASTNodesTest(unittest.TestCase):

    def test_operator_node(self):
        node = OperatorNode(formula_range_address_union.tokens[1], None)
        self.assertEqual(
            node.emit(formula_range_address_union.ast_graph, 'Sheet1'),
            'eval_ref("Sheet1!A1:B1"),eval_ref("Sheet1!A2:B2")'
        )

    # TODO: Might be good to resolve the address representation descrepency between RangeNode and XLRange. Maybe there's synergy for large ranges.
    def test_range_node(self):
        range_node = RangeNode(formula_range_address_function.tokens[1], None)
        # this is different to the way XLRange stores range addresses
        get_cells = ['A1','B1']
        self.assertEqual(get_cells, range_node.get_cells())

        # emit
        range_node_emit = range_node.emit(formula_range_address_function.ast_graph, 'Sheet1')
        emit = 'eval_ref("Sheet1!A1:B1")'
        self.assertEqual(emit, range_node_emit)


    def test_function_node(self):
        function_node = FunctionNode(formula_range_address_function.tokens[0], None)
        function_node_emit = function_node.emit(formula_range_address_function.ast_graph, 'Sheet1')
        emit = 'SUM(eval_ref("Sheet1!A1:B1"))'
        self.assertEqual(emit, function_node_emit)


class RangeNodeTest(unittest.TestCase):

    def create_node(self, range):
        return RangeNode(
            f_token(tvalue=range, ttype='operand', tsubtype='range'), None)

    def test_get_cells(self):
        self.assertEqual(
            self.create_node('A1').get_cells(), ['A1'])
        self.assertEqual(
            self.create_node('A1:A3').get_cells(), [['A1'], ['A2'], ['A3']])
        self.assertEqual(
            self.create_node('S!A1:A2').get_cells(), [['S!A1'], ['S!A2']])
