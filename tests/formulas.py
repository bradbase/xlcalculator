from collections import namedtuple

from xlcalculator.ast_nodes import FunctionNode, RangeNode, OperatorNode
from xlcalculator.model import Model
from xlcalculator.tokenizer import f_token


Formula = namedtuple(
    'Formula',
    'formula sheet_name return_type tokens reverse_polish_tokens stack')


formula_cell_minus = Formula(
        '-A1',
        'Sheet1',
        None,
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='-', ttype='operator-prefix', tsubtype='')), OperatorNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'))],
        None
    )


formula_cells_minus = Formula(
        '-A1,-B1',
        'Sheet1',
        None,
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cell_minus_with_sheet = Formula(
        '-Sheet1!A1',
        'Sheet1',
        None,
        [f_token(tvalue='-', ttype='operator-prefix', tsubtype=''), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cell_address = Formula(
        'A1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='A1', ttype='operand', tsubtype='range'))],
        None
    )


formula_cell_address_with_sheet = Formula(
        'Sheet1!A1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'))],
        None
    )


formula_cells_union = Formula(
        'A1,B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_union_with_same_sheet = Formula(
        'Sheet1!A1,Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [RangeNode(f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range')), RangeNode(f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')), OperatorNode(f_token(tvalue=',', ttype='operator-infix', tsubtype='union'))],
        None
    )


formula_cells_union_with_different_sheet = Formula(
        'Sheet1!A1,Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='operator-infix', tsubtype='union'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_addition = Formula(
        'A1+B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_addition_with_same_sheet = Formula(
        'Sheet1!A1+Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_addition_with_different_sheet = Formula(
        'Sheet1!A1+Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_subtract = Formula(
        'A1-B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_subtract_with_same_sheet = Formula(
        'Sheet1!A1-Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_subtract_with_different_sheet = Formula(
        'Sheet1!A1-Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='-', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_multiply = Formula(
        'A1*B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_multiply_with_same_sheet = Formula(
        'Sheet1!A1*Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_multiply_with_different_sheet = Formula(
        'Sheet1!A1*Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='*', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_divide = Formula(
        'A1/B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_divide_with_same_sheet = Formula(
        'Sheet1!A1/Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_divide_with_different_sheet = Formula(
        'Sheet1!A1/Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='/', ttype='operator-infix', tsubtype='math'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_is_equal = Formula(
        'A1==B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_is_equal_with_same_sheet = Formula(
        'Sheet1!A1==Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_is_equal_with_different_sheet = Formula(
        'Sheet1!A1==Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='=', ttype='operator-infix', tsubtype='logical'), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_function = Formula(
        'SUM(A1,B1)',
        'Sheet1',
        None,
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        None
    )

formula_cells_function_with_same_sheet = Formula(
        'SUM(Sheet1!A1,Sheet1!B1)',
        'Sheet1',
        None,
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet1!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_is_equal_with_different_sheet = Formula(
        'SUM(Sheet1!A1,Sheet2!B1)',
        'Sheet1',
        None,
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='Sheet1!A1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='Sheet2!B1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        None
    )


formula_cells_function_and_operator = Formula(
        'SUM(A1+B1,C1)',
        'Sheet1',
        None,
        [f_token(tvalue='SUM', ttype='function', tsubtype='start'), f_token(tvalue='A1', ttype='operand', tsubtype='range'), f_token(tvalue='+', ttype='operator-infix', tsubtype='math'), f_token(tvalue='B1', ttype='operand', tsubtype='range'), f_token(tvalue=',', ttype='argument', tsubtype=''), f_token(tvalue='C1', ttype='operand', tsubtype='range'), f_token(tvalue='', ttype='function', tsubtype='stop')],
        [], # reverse_polish_tokens
        None
    )

formula_range_address = Formula(
        'A1:B1',
        'Sheet1',
        None,
        [f_token(tvalue='A1:B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )

formula_range_address_with_sheet = Formula(
        'Sheet1!A1:B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1:B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )

formula_range_address_with_same_sheets = Formula(
        'Sheet1!A1:Sheet1!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1:Sheet1!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )

formula_range_address_with_different_sheets = Formula(
        'Sheet1!A1:Sheet2!B1',
        'Sheet1',
        None,
        [f_token(tvalue='Sheet1!A1:Sheet2!B1', ttype='operand', tsubtype='range')],
        [], # reverse_polish_tokens
        None
    )
