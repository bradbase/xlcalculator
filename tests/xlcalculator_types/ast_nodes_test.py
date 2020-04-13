import unittest

from networkx import info

from xlcalculator.xlcalculator_types import OperatorNode
# from xlcalculator.xlcalculator_types import OperandNode
from xlcalculator.xlcalculator_types import RangeNode
from xlcalculator.xlcalculator_types import FunctionNode
from xlcalculator.read_excel.tokenizer import f_token

from ..formulas import *


class TestASTNodes(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_operator_node(self):
        operator_node = OperatorNode(formula_range_address_union.tokens[1], None)
        operator_node_emit = operator_node.emit(formula_range_address_union.ast_graph, 'Sheet1')
        emit = 'self.eval_ref("Sheet1!A1:B1"),self.eval_ref("Sheet1!A2:B2")'
        self.assertEqual(emit, operator_node_emit)


    @unittest.skip("No examples currently exist in the example spreadsheet")
    def test_operand_node(self):
        pass

    # TODO: Might be good to resolve the address representation descrepency between RangeNode and XLRange. Maybe there's synergy for large ranges.
    def test_range_node(self):
        range_node = RangeNode(formula_range_address_function.tokens[1], None)
        get_cells = ['A1','B1'] # this is different to the way XLRange stores range addresses
        self.assertEqual(get_cells, range_node.get_cells())

        # emit
        range_node_emit = range_node.emit(formula_range_address_function.ast_graph, 'Sheet1')
        emit = 'self.eval_ref("Sheet1!A1:B1")'
        self.assertEqual(emit, range_node_emit)


    def test_function_node(self):
        function_node = FunctionNode(formula_range_address_function.tokens[0], None)
        function_node_emit = function_node.emit(formula_range_address_function.ast_graph, 'Sheet1')
        emit = 'xSum.xsum(self.eval_ref("Sheet1!A1:B1"))'
        self.assertEqual(emit, function_node_emit)
