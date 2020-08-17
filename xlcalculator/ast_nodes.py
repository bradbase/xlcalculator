import inspect

from xlcalculator.xlfunctions import (
    xl,
    xlerrors,
    math,
    operator,
    text,
    func_xltypes
)

from . import utils

PREFIX_OP_TO_FUNC = {
    '-': operator.OP_NEG,
}

POSTFIX_OP_TO_FUNC = {
    '%': operator.OP_PERCENT,
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


class EvalContext:

    cells = None
    ranges = None
    namespace = None
    seen = None
    ref = None

    def __init__(self, namespace=None, ref=None, seen=None):
        self.seen = seen if seen is not None else []
        self.namespace = namespace if namespace is not None else xl.FUNCTIONS
        self.ref = ref

    def eval_cell(self, addr):
        raise NotImplementedError()


class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self, token):
        self.token = token

    @property
    def tvalue(self):
        return self.token.tvalue

    @property
    def ttype(self):
        return self.token.ttype

    @property
    def tsubtype(self):
        return self.token.tsubtype

    def __eq__(self, other):
        return self.token == other.token

    def eval(self, context):
        raise NotImplementedError(f'`eval()` of {self}')

    def __repr__(self):
        return (
            f"<{self.__class__.__name__} "
            f"tvalue: {repr(self.tvalue)}, "
            f"ttype: {self.ttype}, "
            f"tsubtype: {self.tsubtype}"
            f">"
        )

    def __str__(self):
        return str(self.tvalue)

    def __iter__(self):
        yield self


class OperandNode(ASTNode):

    def eval(self, context):
        if self.tsubtype == "logical":
            return func_xltypes.Boolean.cast(self.tvalue)
        elif self.tsubtype == 'text':
            return func_xltypes.Text(self.tvalue)
        elif self.tsubtype == 'error':
            if self.tvalue in xlerrors.ERRORS_BY_CODE:
                return xlerrors.ERRORS_BY_CODE[self.tvalue](
                    f'Error in cell ${context.ref}')
            return xlerrors.ExcelError(
                self.tvalue, f'Error in cell ${context.ref}')
        else:
            return func_xltypes.Number.cast(self.tvalue)

    def __str__(self):
        if self.tsubtype == "logical":
            return self.tvalue.title()
        elif self.tsubtype == "text":
            return '"' + self.tvalue.replace('"', '\\"') + '"'
        return str(self.tvalue)


class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range."""

    def get_cells(self):
        cells = utils.resolve_ranges(self.tvalue, default_sheet='')[1]
        return cells[0] if len(cells) == 1 else cells

    @property
    def address(self):
        return self.tvalue

    def full_address(self, ref):
        addr = self.address
        if '!' not in addr:
            sheet, _, _ = utils.resolve_address(ref)
            addr = f'{sheet}!{addr}'
        return addr

    def eval(self, context):
        addr = self.full_address(context.ref)

        if addr in context.ranges:
            range_cells = [
                [
                    context.eval_cell(addr)
                    for addr in range_row
                ]
                for range_row in context.ranges[addr].cells
            ]
            context.ranges[addr].value = data = func_xltypes.Array(range_cells)
            return data

        return context.eval_cell(addr)


class OperatorNode(ASTNode):

    def __init__(self, token):
        super().__init__(token)
        self.left = None
        self.right = None

    def eval(self, context):
        if self.ttype == 'operator-prefix':
            assert self.left is None, 'Left operand for prefix operator'
            op = PREFIX_OP_TO_FUNC[self.tvalue]
            return op(self.right.eval(context))

        elif self.ttype == 'operator-infix':
            op = INFIX_OP_TO_FUNC[self.tvalue]
            return op(
                self.left.eval(context),
                self.right.eval(context),
            )
        elif self.ttype == 'operator-postfix':
            assert self.right is None, 'Right operand for postfix operator'
            op = POSTFIX_OP_TO_FUNC[self.tvalue]
            return op(self.left.eval(context))
        else:
            raise ValueError(f'Invalid operator type: {self.ttype}')

    def __str__(self):
        left = f'({self.left}) ' if self.left is not None else ''
        right = f' ({self.right})' if self.right is not None else ''
        return f'{left}{self.tvalue}{right}'

    def __iter__(self):
        # Return node in resolution order.
        yield self.left
        yield self.right
        yield self


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, token):
        super().__init__(token)
        self.args = None

    def eval(self, context):
        func_name = self.tvalue.upper()
        # 1. Remove the BBB namespace, since we are just supporting
        #    everything in one large one.
        func_name = func_name.replace('_XLFN.', '')
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
                args.append(func_xltypes.Expr(
                    pvalue.eval, (context,), ref=context.ref, ast=pvalue
                ))
            elif (param.kind == param.VAR_POSITIONAL
                  and func_xltypes.XlExpr in getattr(ptype, '__args__', [])):
                args.extend([
                    func_xltypes.Expr(
                        pitem.eval, (context,), ref=context.ref, ast=pitem
                    )
                    for pitem in pvalue
                ])
            elif (param.kind == param.VAR_POSITIONAL):
                args.extend([
                    pitem.eval(context) for pitem in pvalue
                ])
            else:
                args.append(pvalue.eval(context))
        # 4. Run function and return result.
        return func(*args)

    def __str__(self):
        args = ', '.join(str(arg) for arg in self.args)
        return f'{self.tvalue}({args})'

    def __iter__(self):
        # Return node in resolution order.
        for arg in self.args:
            yield arg
        yield self
