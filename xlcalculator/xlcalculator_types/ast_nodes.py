import logging

from xlfunctions import FUNCTIONS, ExcelError
from . import utils


class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self, token):
        self.token = token

    def __str__(self):
        return self.token.tvalue

    def __getattr__(self, name):
        return getattr(self.token, name)

    def __hash__(self):
        return hash( self.token.tvalue )

    def __eq__(self, other):
        return self.token == other.token

    def children(self, ast):
        """"""

        args = ast.predecessors(self)
        args = sorted(args, key=lambda x: ast.nodes[x]['pos'])
        return args

    def parent(self, ast):
        """"""

        args = list(ast.successors(self))
        return args[0] if args else None

    def find_special_function(self, ast):
        """"""

        found = False
        current = self
        special_functions = ['sumproduct']
        break_functions = ['index']

        while current is not None:
            if current.tvalue.lower() in special_functions:
                found = True
                break

            elif current.tvalue.lower() in break_functions:
                break

            else:
                current = current.parent(ast)

        return found

    def has_operator_or_func_parent(self, ast):
        """"""

        found = False
        current = self

        while current is not None:
            if ((current.ttype[:8] == 'operator'
                     or current.ttype == 'function')
                    and current.tvalue.lower() != 'if'):
                found = True
                break

            else:
                current = current.parent(ast)

        return found

    def has_ind_func_parent(self, ast):
        """"""

        if self.parent(ast) is not None and self.parent(ast).tvalue in IND_FUN:
            return True

        else:
            return False

    def emit(self, ast, sheet_name=None, pointer=False):
        """Emit code"""

        self.token.tvalue

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
    """"""

    def __init__(self, args, ref):
        super().__init__(args)
        # ref is the address of the reference cell
        self.ref = ref if ref != '' else 'None'
        # convert the operator to python equivalents
        self.opmap = {
                 "^":"**",
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

        self.op_range_translator = {
            "*": "multiply",
            "/": "divide",
            "+": "add",
            "-": "substract",
            "==": "is_equal",
            "<>": "is_not_equal",
            ">": "is_strictly_superior",
            "<": "is_strictly_inferior",
            ">=": "is_superior_or_equal",
            "<=": "is_inferior_or_equal"
        }


    def emit(self, ast, sheet_name=None, pointer=False):
        xop = self.tvalue

        # Get the arguments
        args = self.children(ast)

        op = self.opmap.get(xop,xop)

        parent = self.parent(ast)
        # convert ":" operator to a range function
        if op == ":":
            return "eval_ref(%s)" % (','.join([
                a.emit(ast, sheet_name=sheet_name) for a in args]))

        if self.ttype == "operator-prefix":
            return 'apply_one("minus", {}, None, {})'.format(
                args[0].emit(ast, sheet_name=sheet_name), str(self.ref))

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '')
            function = self.op_range_translator.get(op)

            arg1 = str(args[0].emit(ast,sheet_name=sheet_name))
            arg2 = str(args[1].emit(ast,sheet_name=sheet_name))
            return f'{call}("{function}", {arg1}, {arg2})'

        parent = self.parent(ast)

        # TODO silly hack to work around the fact that None < 0 is True
        # (happens on blank cells)
        # This is just wrong. Implement this handling properly in the operator
        # functions.
        left = args[0].emit(ast, sheet_name=sheet_name)
        right = args[1].emit(ast,sheet_name=sheet_name)
        if op == "<" or op == "<=":
            ss = f"({left} if {left} is not None else float('inf')){op}{right}"

        elif op == ">" or op == ">=":
            ss = f'{left}{op}({right} if {right} is not None else float("inf"))'

        else:
            ss = f'{left}{op}{right}'

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "("+ ss + ")"

        return ss


class OperandNode(ASTNode):
    """"""

    def __init__(self, *args):
        super().__init__(*args)


    def emit(self, ast, sheet_name=None, pointer=False):
        """"""

        t = self.tsubtype

        if t == "logical":
            return str(self.tvalue.lower() == "true")

        elif t == "text" or t == "error":
            val = self.tvalue.replace('"','\\"').replace("'","\\'")
            return str('"' + val + '"')

        else:
            return str(self.tvalue)


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


    def emit(self, ast, sheet_name=None, pointer=False):
        """"""

        if isinstance(self.tvalue, ExcelError):
            logging.warning( "Excel Error Code found {}".format(self.tvalue) )
            return self.tvalue

        is_a_named_range = self.tsubtype == "named_range"

        if is_a_named_range:
            my_str = '"{}"'.format( self.token.tvalue )

        else:
            rng = self.tvalue.replace('$','')
            if self.tsubtype == 'pointer':
                print("I FOUND A POINTER!!!", rng)
                my_str = '"' + rng + '"'

            else:
                if '!' in rng:
                    my_str = '"' + rng + '"'

                else:
                    my_str = '"' + sheet_name + "!" + rng + '"'

        return f'eval_ref({my_str})'


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, args, ref):
        super().__init__(args)
        # ref is the address of the reference cell
        self.ref = ref if ref != '' else 'None'

    def emit(self, ast, sheet_name=None, pointer=False):
        func = self.tvalue.upper()

        # Get the arguments
        args = self.children(ast)

        # ARRAY and ARRAYROW are two special functions created by the tokenizer.
        if func == "ARRAY":
            my_str = '['
            if len(args) == 1:
                # only one row
                my_str += args[0].emit(ast, sheet_name=sheet_name)

            else:
                # multiple rows
                my_str += ",".join([
                    '[' + n.emit(ast, sheet_name=sheet_name) + ']'
                    for n in args
                ])

            my_str += ']'

            return my_str

        elif func == "ARRAYROW":
            #simply create a list
            return ",".join([n.emit(ast, sheet_name=sheet_name) for n in args])

        else:
            # Remove the BBB namespace, since we are just supporting
            # everything in one large one.
            func = func.replace('_XLFN.', '')
            # map to the correct name
            if func in FUNCTIONS:
                args_str = ','.join(
                    [n.emit(ast, sheet_name=sheet_name) for n in args])
                return f'{func}({args_str})'

            else:
                message = f"Function {func} is not supported in xlcalculator."
                logging.error(message)

