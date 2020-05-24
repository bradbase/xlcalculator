import logging

from xlfunctions import xl, math, operator, text
from . import utils

OP_TO_FUNC = {
    "*": operator.OP_MUL,
    "/": operator.OP_DIV,
    "+": operator.OP_ADD,
    "-": operator.OP_SUB,
    "^": math.POWER,
    "&": text.CONCAT,
    "=": operator.OP_EQ,
    "<>": operator.OP_NE,
    ">": operator.OP_GT,
    "<": operator.OP_LT,
    ">=": operator.OP_GE,
    "<=": operator.OP_LE,
}


class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self, token):
        self.token = token

    def __str__(self):
        return str(self.token.tvalue)

    def __getattr__(self, name):
        return getattr(self.token, name)

    def __hash__(self):
        return hash( self.token.tvalue )

    def __eq__(self, other):
        return self.token == other.token

    def eval(self, model, namespace, ref):
        raise NotImplementedError(f'`eval()` of {self}')

    def __repr__(self):
        return (
            f"<{self.__class__.__name__} "
            f"tvalue: {self.tvalue} "
            f"ttype: {self.ttype} "
            f"tsubtype: {self.tsubtype}"
            f">"
        )

    __str__ = __repr__


class OperatorNode(ASTNode):

    def __init__(self, args, ref):
        super().__init__(args)
        # ref is the address of the reference cell
        self.ref = ref if ref != '' else 'None'
        self.left = None
        self.right = None

    def eval(self, model, namespace, ref):
        op = OP_TO_FUNC[self.tvalue]
        return op(
            self.left.eval(model, namespace, ref),
            self.right.eval(model, namespace, ref),
        )


class OperandNode(ASTNode):

    def __init__(self, *args):
        super().__init__(*args)

    def eval(self, model, namespace, ref):
        typ = self.tsubtype
        if typ == "logical":
            return self.tvalue.lower() == "true"
        elif typ == "text" or typ == "error":
            return self.tvalue
        else:
            return xl.convert_number(self.tvalue)


class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range.

       e.g., A5, B3:C20 or INPUT
    """

    def __init__(self, args, ref):
        super().__init__(args)
        # ref is the address of the reference cell
        self.ref = ref if ref != '' else 'None'

    def get_cells(self):
        cells = utils.resolve_ranges(self.tvalue, default_sheet='')[1]
        return cells[0] if len(cells) == 1 else cells

    def eval(self, model, namespace, ref):
        addr = self.token.tvalue

        if '!' not in addr:
            sheet, _, _  = utils.resolve_address(ref)
            addr = f'{sheet}!{addr}'

        if addr in model.cells:
            return model.cells[addr].value

        if addr in model.ranges:
            range_cells = []
            for range_column in model.ranges[addr].cells:
                row = []

                for cell_addr in range_column:
                    row.append(model.cells[cell_addr].value )

                range_cells.append(row)

            model.ranges[addr].value = xl.RangeData(range_cells)
            return model.ranges[addr].value

        # Remove sheet part to look up the name. (The sheet was added by the
        # parser to make complete cell and rnage addresses.
        sheet, _, _  = utils.resolve_address(ref)
        addr = addr[len(sheet)+1:]
        if addr in model.defined_names:
            # TODO: support defined name to be cell, range and formula
            return model.defined_names[addr].value


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, args, ref):
        super().__init__(args)
        # ref is the address of the reference cell
        self.ref = ref if ref != '' else 'None'
        self.args = None

    def eval(self, model, namespace, ref):
        func_name = self.tvalue.upper()
        # Remove the BBB namespace, since we are just supporting
        # everything in one large one.
        func_name = func_name.replace('_XLFN.', '')
        func = namespace[func_name]
        return func(
            *[arg.eval(model, namespace, ref) for arg in self.args]
        )
