
from collections import namedtuple

import networkx as nx

from koala_xlcalculator.read_excel.tokenizer import f_token
from koala_xlcalculator.koala_types.ast_nodes import RangeNode
from koala_xlcalculator.koala_types.ast_nodes import OperatorNode
# from koala_xlcalculator.koala_types.ast_nodes import OperandNode
from koala_xlcalculator.koala_types.ast_nodes import FunctionNode

from koala_xlcalculator.model.model import Model


Formula = namedtuple('Formula', 'address tokens reverse_polish_tokens ast_graph stack')


formula_cell_minus = Formula(
        '-A1',
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), None), OperatorNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'), None)],
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_minus = Formula(
        '-A1,-B1',
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cell_minus_with_sheet = Formula(
        '-Sheet1!A1',
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cell_address = Formula(
        'A1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'), None)],
        nx.DiGraph(), # ast_graph
        None
    )


formula_cell_address_with_sheet = Formula(
        'Sheet1!A1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), None)],
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_union = Formula(
        'A1,B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_union_with_same_sheet = Formula(
        'Sheet1!A1,Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), None), RangeNode(f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range'), None), OperatorNode(f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), None)],
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_union_with_different_sheet = Formula(
        'Sheet1!A1,Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_addition = Formula(
        'A1+B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_addition_with_same_sheet = Formula(
        'Sheet1!A1+Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_addition_with_different_sheet = Formula(
        'Sheet1!A1+Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_subtract = Formula(
        'A1-B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_subtract_with_same_sheet = Formula(
        'Sheet1!A1-Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_subtract_with_different_sheet = Formula(
        'Sheet1!A1-Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_multiply = Formula(
        'A1*B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_multiply_with_same_sheet = Formula(
        'Sheet1!A1*Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_multiply_with_different_sheet = Formula(
        'Sheet1!A1*Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_divide = Formula(
        'A1/B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_divide_with_same_sheet = Formula(
        'Sheet1!A1/Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_divide_with_different_sheet = Formula(
        'Sheet1!A1/Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_is_equal = Formula(
        'A1==B1',
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_is_equal_with_same_sheet = Formula(
        'Sheet1!A1==Sheet1!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_is_equal_with_different_sheet = Formula(
        'Sheet1!A1==Sheet2!B1',
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_function = Formula(
        'SUM(A1,B1)',
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

formula_cells_function_with_same_sheet = Formula(
        'SUM(Sheet1!A1,Sheet1!B1)',
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_is_equal_with_different_sheet = Formula(
        'SUM(Sheet1!A1,Sheet2!B1)',
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )


formula_cells_function_and_operator = Formula(
        'SUM(A1+B1,C1)',
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='C1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

formula_range_address = Formula(
        'A1:B1',
        [f_token(tvalue='A1:B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

formula_range_address_with_sheet = Formula(
        'Sheet1!A1:B1',
        [f_token(tvalue='Sheet1!A1:B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

formula_range_address_with_same_sheets = Formula(
        'Sheet1!A1:Sheet1!B1',
        [f_token(tvalue='Sheet1!A1:Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

formula_range_address_with_different_sheets = Formula(
        'Sheet1!A1:Sheet2!B1',
        [f_token(tvalue='Sheet1!A1:Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        nx.DiGraph(), # ast_graph
        None
    )

# model = Model()
# range_thing = {'' : [['A1', 'A2'],['B1','B2']]}
# tokens = model.shunting_yard('A1:B1,A2:B2', range_thing)
# ast_graph = model.build_ast(tokens)[0]
# print(nx.info(ast_graph))
# for node in ast_graph.nodes():
#     print("node", node, type(node), node.tvalue, node.ttype, node.tsubtype)
# for edge in ast_graph.edges():
#     print("edge", edge[0], edge[1], type(edge[0]), type(edge[1]))
#
ast_graph_range_address_union = nx.DiGraph()
ast_graph_range_address_union.add_node(RangeNode(f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), None), pos=1)
ast_graph_range_address_union.add_node(RangeNode(f_token(tvalue='A2:B2', ttype='operand', tsubtype='range'), None), pos=2)
ast_graph_range_address_union.add_node(OperatorNode(f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), None), pos=3)
ast_graph_range_address_union.add_edge(RangeNode(f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), None), OperatorNode(f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), None) )
ast_graph_range_address_union.add_edge(RangeNode(f_token(tvalue='A2:B2', ttype='operand', tsubtype='range'), None), OperatorNode(f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), None) )
formula_range_address_union = Formula(
        'A1:B1,A2:B2',
        [f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='A2:B2', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        ast_graph_range_address_union,
        None
    )


ast_graph_range_address_function = nx.DiGraph()
ast_graph_range_address_function.add_node(RangeNode(f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), None), pos=1)
ast_graph_range_address_function.add_node(FunctionNode(f_token(tvalue='SUM', ttype='function', tsubtype='start'), None), pos=2)
ast_graph_range_address_function.add_edge(RangeNode(f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), None), FunctionNode(f_token(tvalue='SUM', ttype='function', tsubtype='start'), None))
formula_range_address_function = Formula(
        'SUM(A1:B1)',
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [FunctionNode(f_token(tvalue='SUM', ttype='function', tsubtype='start'), None), RangeNode(f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'), None)], # reverse_polish_tokens
        ast_graph_range_address_function, # ast_graph
        FunctionNode(f_token(tvalue='SUM', ttype='function', tsubtype=''), None) # stack
    )
