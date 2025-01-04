import inspect
from typing import Any, Callable, Union

from xlcalculator.tokenizer import f_token
from xlcalculator.xlfunctions import xl, xlerrors, math, operator, text, func_xltypes

import xlcalculator.utils as utils

PREFIX_OP_TO_FUNC = {
    "-": operator.OP_NEG,
}

POSTFIX_OP_TO_FUNC = {
    "%": operator.OP_PERCENT,
}

INFIX_OP_TO_FUNC = {
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

MAX_EMPTY = 100
xl_or_none = Union[func_xltypes.XlAnything, None]


# TODO: what is the purpose of ranges and cells
class EvalContext:
    cells: dict = None
    ranges: dict = None
    namespace: dict[str, Callable] = None
    seen: list[str] = None
    ref: str = None
    refsheet: str = None
    sheet: str = None

    def __init__(
        self, namespace: dict[str, Callable] = None, ref: str = None, seen: list = None
    ):
        self.seen = seen if seen is not None else []
        self.namespace = namespace if namespace is not None else xl.FUNCTIONS
        self.ref = ref
        self.sheet = self.refsheet = ref.split("!")[0]

    def eval_cell(self, addr):
        raise NotImplementedError()

    def set_sheet(self, sheet=None):
        if sheet is None:
            self.sheet = self.refsheet
        else:
            self.sheet = sheet


# TODO: update class to remove object inheritance, no longer needed since like 3.7+?
# TODO: double check the __iter__ I dont think thats correct ...
class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self, token: f_token):
        self.token = token

    @property
    def tvalue(self) -> str:
        return self.token.tvalue

    @property
    def ttype(self) -> str:
        return self.token.ttype

    @property
    def tsubtype(self) -> str:
        return self.token.tsubtype

    def __eq__(self, other) -> bool:
        return self.token == other.token

    def eval(self, context):
        raise NotImplementedError(f"`eval()` of {self}")

    def __repr__(self) -> str:
        return (
            f"<{self.__class__.__name__} "
            f"tvalue: {repr(self.tvalue)}, "
            f"ttype: {self.ttype}, "
            f"tsubtype: {self.tsubtype}"
            f">"
        )

    def __str__(self) -> str:
        return str(self.tvalue)

    def __iter__(self):
        yield self


# TODO: why in the world would you return an error and not raise it? what is this? rust?
class OperandNode(ASTNode):
    def eval(self, context):
        if self.tsubtype == "logical":
            return func_xltypes.Boolean.cast(self.tvalue)
        elif self.tsubtype == "text":
            return func_xltypes.Text(self.tvalue)
        elif self.tsubtype == "error":
            if self.tvalue in xlerrors.ERRORS_BY_CODE:
                return xlerrors.ERRORS_BY_CODE[self.tvalue](
                    f"Error in cell ${context.ref}"
                )
            return xlerrors.ExcelError(self.tvalue, f"Error in cell ${context.ref}")
        else:
            return func_xltypes.Number.cast(self.tvalue)

    def __str__(self) -> str:
        if self.tsubtype == "logical":
            return self.tvalue.title()
        elif self.tsubtype == "text":
            return '"' + self.tvalue.replace('"', '\\"') + '"'
        return str(self.tvalue)


class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range."""

    # TODO: check the return, that seems odd
    def get_cells(self):
        cells = utils.resolve_ranges(self.tvalue, default_sheet="")[1]
        return cells[0] if len(cells) == 1 else cells

    @property
    def address(self) -> str:
        return self.tvalue

    def full_address(self, context: EvalContext) -> str:
        addr = self.address
        if "!" not in addr:
            addr = f"{context.sheet}!{addr}"
        return addr

    # TODO: remove the Any and figure out the proper return types? xlAnything?
    def eval(self, context: EvalContext) -> Union[func_xltypes.Array, Any]:
        addr = self.full_address(context)

        if addr in context.ranges:
            empty_row = 0
            empty_col = 0
            range_cells = []
            for range_row in context.ranges[addr].cells:
                row_cells = []
                for col_addr in range_row:
                    cell = context.eval_cell(col_addr)
                    if cell.value == "" or cell.value is None:
                        empty_col += 1
                        if empty_col > MAX_EMPTY:
                            break
                    else:
                        empty_col = 0
                    row_cells.append(cell)
                if not row_cells:
                    empty_row += 1
                    if empty_row > MAX_EMPTY:
                        break
                else:
                    empty_row = 0
                range_cells.append(row_cells)
            context.ranges[addr].value = data = func_xltypes.Array(range_cells)
            return data

        value = context.eval_cell(addr)
        context.set_sheet()
        return value


class OperatorNode(ASTNode):
    def __init__(self, token: f_token) -> None:
        super().__init__(token)
        self.left: xl_or_none = None
        self.right: xl_or_none = None

    def eval(self, context: EvalContext) -> func_xltypes.XlAnything:
        if self.ttype == "operator-prefix":
            assert self.left is None, "Left operand for prefix operator"
            op = PREFIX_OP_TO_FUNC[self.tvalue]
            return op(self.right.eval(context))

        elif self.ttype == "operator-infix":
            op = INFIX_OP_TO_FUNC[self.tvalue]
            return op(
                self.left.eval(context),
                self.right.eval(context),
            )
        elif self.ttype == "operator-postfix":
            assert self.right is None, "Right operand for postfix operator"
            op = POSTFIX_OP_TO_FUNC[self.tvalue]
            return op(self.left.eval(context))
        else:
            raise ValueError(f"Invalid operator type: {self.ttype}")

    def __str__(self) -> str:
        left = f"({self.left}) " if self.left is not None else ""
        right = f" ({self.right})" if self.right is not None else ""
        return f"{left}{self.tvalue}{right}"

    def __iter__(self):
        # Return node in resolution order.
        yield self.left
        yield self.right
        yield self


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, token: f_token):
        super().__init__(token)
        self.args = None

    def eval(self, context: EvalContext) -> func_xltypes.XlAnything:
        func_name = self.tvalue.upper()
        # 1. Remove the BBB namespace, since we are just supporting
        #    everything in one large one.
        func_name = func_name.replace("_XLFN.", "")
        # 2. Look up the function to use.
        func = context.namespace[func_name]
        # 3. Prepare arguments.
        sig = inspect.signature(func)
        bound = sig.bind(*self.args)
        args = []
        for pname, pvalue in list(bound.arguments.items()):
            param = sig.parameters[pname]
            ptype = param.annotation
            if ptype == func_xltypes.XlExpr:
                args.append(
                    func_xltypes.Expr(
                        pvalue.eval, (context,), ref=context.ref, ast=pvalue
                    )
                )
            elif param.kind == param.VAR_POSITIONAL and func_xltypes.XlExpr in getattr(
                ptype, "__args__", []
            ):
                args.extend(
                    [
                        func_xltypes.Expr(
                            pitem.eval, (context,), ref=context.ref, ast=pitem
                        )
                        for pitem in pvalue
                    ]
                )
            elif param.kind == param.VAR_POSITIONAL:
                args.extend([pitem.eval(context) for pitem in pvalue])
            else:
                args.append(pvalue.eval(context))
        # 4. Run function and return result.
        return func(*args)

    def __str__(self) -> str:
        args = ", ".join(str(arg) for arg in self.args)
        return f"{self.tvalue}({args})"

    def __iter__(self):
        # Return node in resolution order.
        for arg in self.args:
            yield arg
        yield self
