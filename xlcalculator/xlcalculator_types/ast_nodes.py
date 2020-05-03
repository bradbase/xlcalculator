
import logging
import re
from string import ascii_uppercase

from xlfunctions import *

# from ..function_library import *
from xlfunctions.exceptions import ExcelError

INTERPRETER_OBJECT = "Evaluator"


def to_str(my_string):
    """"""

    try:

        return str(my_string)

    except:
        logging.warning("Couldnt parse {} as string {}".format(my_string, type(my_string)))
        return my_string


def is_range(address):
    """"""

    if isinstance(address, Exception):
        return address

    return address.find(':') > 0


def split_range(rng):
    """"""

    if rng.find('!') > 0:
        start, end = rng.split(':')
        if start.find('!') > 0:
            sh, start = start.split("!")

        if end.find('!') > 0:
            sh, end = end.split("!")

    else:
        sh = None
        start, end = rng.split(':')

    return (sh, start, end)


def split_address(address):
    """"""

    sheet = None
    if address.find('!') > 0:
        sheet, addr = address.split('!')

    else:
        addr = address

    #ignore case
    addr = addr.upper()

    # regular <col><row> format
    if re.match('^[A-Z\$]+[\d\$]+$', addr):
        logging.debug("WE MATCH <col><row> {}".format(addr))
        col, row = [_f for _f in re.split('([A-Z\$]+)', addr) if _f]

    # R<row>C<col> format
    elif re.match('^R\d+C\d+$', addr):
        logging.debug("WE MATCH R<row>C<col> {}".format(addr))
        row, col = addr.split('C')
        row = row[1:]

    # R[<row>]C[<col>] format
    elif re.match('^R\[\d+\]C\[\d+\]$', addr):
        logging.debug("WE MATCH R[<row>]C[<col>] {}".format(addr))
        row, col = addr.split('C')
        row = row[2:-1]
        col = col[2:-1]

    else:
        print("WE DONT MATCH ANYTHING {}".format(addr))
        message = 'Invalid address format ' + addr
        logging.debug(message)
        raise Exception(message)

    return (sheet, col, row)


def col2num(col):
    if not col:
        raise Exception("Column may not be empty")

    tot = 0
    for i,c in enumerate([c for c in col[::-1] if c != "$"]):
        if c == '$': continue
        tot += (ord(c)-64) * 26 ** i

    return tot


def num2col(num):
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)

    s = ''
    q = num
    while q > 0:
        (q,r) = divmod(q,26)
        if r == 0:
            q = q - 1
            r = 26
        s = ascii_uppercase[r-1] + s

    return s


def resolve_range(rng, should_flatten = False, sheet_name=''):
    # print 'RESOLVE RANGE splitting', rng
    if ':' not in rng:
        if '!' in rng:
            rng = rng.split('!')
        return ExcelError('#REF!', info = '%s is not a regular range, nor a named_range' % rng)
    sh, start, end = split_range(rng)

    if sh and sheet_name:
        if sh != sheet_name:
            raise Exception("Mismatched sheets %s and %s" % (sh, sheet_name))
        else:
            sheet_name += '!'
    elif sh and not sheet_name:
        sheet_name = sh + "!"
    elif sheet_name and not sh:
        sheet_name += "!"
    else:
        pass

    # Python strings are unicode capable
    # # `unicode` != `str` in Python2. See `from openpyxl.compat import unicode`
    # if type(sheet_name) == str and str != unicode:
    #     sheet_name = unicode(sheet_name, 'utf-8')
    # if type(rng) == str and str != unicode:
    #     rng = unicode(rng, 'utf-8')

    key = rng + str(should_flatten) + sheet_name

    if not is_range(rng):  return ([sheet_name + rng],1,1)
    # single cell, no range
    if start.isdigit() and end.isdigit():
		# This copes with 5:5 style ranges
        start_col = "A"
        start_row = start
        end_col = "XFD"
        end_row = end
    else:
        sh, start_col, start_row = split_address(start)
        sh, end_col, end_row = split_address(end)

    start_col_idx = col2num(start_col)
    end_col_idx = col2num(end_col);

    start_row = int(start_row)
    end_row = int(end_row)

    # Attempt to use Numpy, not relevant for now

    # num2col_vec = np.vectorize(num2col)
    # r = np.array([range(start_row, end_row + 1),]*nb_col, dtype='a5').T
    # c = num2col_vec(np.array([range(start_col_idx, end_col_idx + 1),]*nb_row))
    # if len(sheet_name)>0:
    #     s = np.chararray((nb_row, nb_col), itemsize=len(sheet_name))
    #     s[:] = sheet_name
    #     c = np.core.defchararray.add(s, c)
    # B = np.core.defchararray.add(c, r)


    # if start_col == end_col:
    #     data = B.T.tolist()[0]
    #     return data, len(data), 1
    # elif start_row == end_row:
    #     data = B.tolist()[0]
    #     return data, 1, len(data)
    # else:
    #     if should_flatten:
    #         return B.flatten().tolist(), 1, nb_col*nb_row
    #     else:
    #         return B.tolist(), nb_row, nb_col

    # single column
    if  start_col == end_col:
        nrows = end_row - start_row + 1
        data = [ "%s%s%s" % (s, c, r) for (s, c, r) in zip([sheet_name]*nrows,[start_col]*nrows,list(range(start_row,end_row+1)))]

        output = data, len(data), 1

    # single row
    elif start_row == end_row:
        ncols = end_col_idx - start_col_idx + 1
        data = [ "%s%s%s" % (s, num2col(c), r) for (s, c, r) in zip([sheet_name]*ncols,list(range(start_col_idx,end_col_idx+1)),[start_row]*ncols)]
        output = data, 1, len(data)

    # rectangular range
    else:
        cells = []
        for r in range(start_row, end_row + 1):
            row = []
            for c in range(start_col_idx, end_col_idx + 1):
                row.append(sheet_name + num2col(c) + str(r))

            cells.append(row)

        if should_flatten:
            # flatten into one list
            l = list(flatten(cells, only_lists=True))
            output = l, len(cells), len(cells[0])
        else:
            output = cells, len(cells), len(cells[0])

    return output


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
            if (current.ttype[:8] == 'operator' or current.ttype == 'function') and current.tvalue.lower() != 'if':
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


class OperatorNode(ASTNode):
    """"""

    def __init__(self, args, ref):
        super().__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell
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
            # OFFSET HANDLER, when the first argument of OFFSET is a range i.e "A1:A2"
            if (parent is not None and (parent.tvalue == 'OFFSET' and parent.children(ast)[0] == self)):
                return '"%s"' % ':'.join([a.emit(ast, sheet_name=sheet_name).replace('"', '') for a in args])

            else:
                return "self.eval_ref(%s)" % (','.join([a.emit(ast, sheet_name=sheet_name) for a in args]))


        if self.ttype == "operator-prefix":
            return '{}.apply_one("minus", {}, None, {})'.format(INTERPRETER_OBJECT, args[0].emit(ast, sheet_name=sheet_name), to_str(self.ref))

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '')
            function = self.op_range_translator.get(op)
            arg1 = args[0]
            arg2 = args[1]

            return INTERPRETER_OBJECT + "." + call + "(%s)" % ','.join(['"'+function+'"', to_str(arg1.emit(ast,sheet_name=sheet_name)), to_str(arg2.emit(ast,sheet_name=sheet_name)), to_str(self.ref)])

        parent = self.parent(ast)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = args[0].emit(ast, sheet_name=sheet_name)
            ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,sheet_name=sheet_name)

        elif op == ">" or op == ">=":
            aa = args[1].emit(ast, sheet_name=sheet_name)
            ss =  args[0].emit(ast, sheet_name=sheet_name) + op + '(' + aa + ' if ' + aa + ' is not None else float("inf"))'

        else:
            ss = args[0].emit(ast, sheet_name=sheet_name) + op + args[1].emit(ast, sheet_name=sheet_name)

        #avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "("+ ss + ")"

        return ss

    def __repr__(self):
        return "tvalue: {} ttype: {} tsubtype: {}".format(self.tvalue, self.ttype, self.tsubtype)

    def __str__(self):
        return self.__repr__()


class OperandNode(ASTNode):
    """"""

    def __init__(self, *args):
        super().__init__(*args)


    def emit(self, ast, sheet_name=None, pointer=False):
        """"""

        t = self.tsubtype

        if t == "logical":
            return to_str(self.tvalue.lower() == "true")

        elif t == "text" or t == "error":
            val = self.tvalue.replace('"','\\"').replace("'","\\'")
            return to_str('"' + val + '"')

        else:
            return to_str(self.tvalue)

    def __repr__(self):
        return "tvalue: {} ttype: {} tsubtype: {}".format(self.tvalue, self.ttype, self.tsubtype)

    def __str__(self):
        return self.__repr__()


class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range, e.g., A5, B3:C20 or INPUT """

    def __init__(self, args, ref):
        super().__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell


    def get_cells(self):
        """"""

        return resolve_range(self.tvalue)[0]


    def emit(self, ast, sheet_name=None, pointer=False):
        """"""

        if isinstance(self.tvalue, ExcelError):
            logging.warning( "Excel Error Code found {}".format(self.tvalue) )
            return self.tvalue

        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        if is_a_named_range:
            my_str = '"{}"'.format( self.token.tvalue )

        else:
            rng = self.tvalue.replace('$','')
            is_a_range = is_range(rng)
            if self.tsubtype == 'pointer':
                print("I FOUND A POINTER!!!", rng)
                my_str = '"' + rng + '"'

            else:
                if is_a_range:
                    sh, start, end = split_range(rng)

                else:
                    try:
                        sh, col, row = split_address(rng)

                    except Exception as exception:
                        logging.warning("Unknown address: {} is not a cell/range reference, nor a named range. Exception type {}".format( to_str(rng), type(exception) ))
                        sh = None

                if sh:
                    my_str = '"' + rng + '"'

                else:
                    my_str = '"' + sheet_name + "!" + rng + '"'

        to_eval = True
        # exception for formulas which use the address and not it content as ":" or "OFFSET"
        parent = self.parent(ast)
        # for OFFSET, it will also depends on the position in the formula (1st position required)
        if (parent is not None and
            (parent.tvalue == ':' or
            (parent.tvalue == 'OFFSET' and parent.children(ast)[0] == self) or
            (parent.tvalue == 'CHOOSE' and parent.children(ast)[0] != self and self.tsubtype == "named_range")) or
            pointer):

            to_eval = False

        # if parent is None and is_a_named_range: # When a named range is referenced in a cell without any prior operation
        #     return 'self.eval_ref(%s)' % (my_str)

        if to_eval == False:
            output = my_str

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s)' % (my_str)

        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s)' % (my_str)

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):

            # we don't use eval_ref here to avoid empty cells (which are not included in Ranges)
            if is_a_named_range:
                output = 'resolve_range(self.named_ranges[%s])' % my_str
            else:
                output = 'resolve_range(%s)' % my_str

        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s)' % (my_str)

        elif (parent is not None and parent.tvalue == 'INDEX' and len(parent.children(ast)) == 3 and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s)' % (my_str)

        # MATCH HANDLER
        elif parent is not None and parent.tvalue == 'MATCH' \
             and (parent.children(ast)[0] == self or len(parent.children(ast)) == 3 and parent.children(ast)[2] == self):
            output = 'self.eval_ref(%s)' % (my_str)

        elif self.find_special_function(ast) or self.has_ind_func_parent(ast):
            output = 'self.eval_ref(%s)' % my_str

        else:
            output = 'self.eval_ref(%s)' % (my_str)

        return output

    def __repr__(self):
        return "tvalue: {} ttype: {} tsubtype: {}".format(self.tvalue, self.ttype, self.tsubtype)

    def __str__(self):
        return self.__repr__()


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, args, ref):

        super().__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell


    def emit(self, ast, sheet_name=None, pointer=False):
        fun = self.tvalue.lower()

        # Get the arguments
        args = self.children(ast)

        if fun == "atan2":
            # swap arguments
            return "atan2(%s,%s)" % (args[1].emit(ast, sheet_name=sheet_name),args[0].emit(ast, sheet_name=sheet_name))

        elif fun == "pi":
            # constant, no parens
            return "pi"

        elif fun == "if":
            # inline the if

            # check if the 'if' is concerning a Range
            is_range = False
            range = None
            childs = args[0].children(ast)

            for child in childs:
                if ':' in child.tvalue and child.tvalue != ':':
                    is_range = True
                    range = child.tvalue
                    break

            if is_range: # hack to filter Ranges when necessary,for instance situations like {=IF(A1:A3 > 0; A1:A3; 0)}
                return '{}.filter(self.eval_ref("{}"))'.format(INTERPRETER_OBJECT, range)

            if len(args) == 2:
                return "{} if {} else 0".format(args[1].emit(ast, sheet_name=sheet_name), args[0].emit(ast, sheet_name=sheet_name))

            elif len(args) == 3:
                return "({} if {} else {})".format(args[1].emit(ast, sheet_name=sheet_name), args[0].emit(ast, sheet_name=sheet_name), args[2].emit(ast, sheet_name=sheet_name))

            else:
                raise Exception("if with {} arguments not supported".format(len(args)))

        elif fun == "array":
            my_str = '['
            if len(args) == 1:
                # only one row
                my_str += args[0].emit(ast, sheet_name=sheet_name)

            else:
                # multiple rows
                my_str += ",".join(['[' + n.emit(ast, sheet_name=sheet_name) + ']' for n in args])

            my_str += ']'

            return my_str

        elif fun == "arrayrow":
            #simply create a list
            return ",".join([n.emit(ast, sheet_name=sheet_name) for n in args])

        elif fun == "and":
            return "all([" + ",".join([n.emit(ast, sheet_name=sheet_name) for n in args]) + "])"

        elif fun == "or":
            return "any([" + ",".join([n.emit(ast, sheet_name=sheet_name) for n in args]) + "])"

        elif fun == "index":
            if pointer or self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                return 'index(' + ",".join([n.emit(ast, sheet_name=sheet_name) for n in args]) + ")"

            else:
                return 'self.eval_ref(index(%s))' % (",".join([n.emit(ast, sheet_name=sheet_name) for n in args]))

        elif fun == "offset":
            if pointer or self.parent(ast) is None or self.parent(ast).tvalue == ':':
                return 'offset(' + ",".join([n.emit(ast, sheet_name=sheet_name) for n in args]) + ")"

            else:
                return 'self.eval_ref(offset(%s))' % (",".join([n.emit(ast, sheet_name=sheet_name) for n in args]))
        else:
            # map to the correct name
            if fun.upper() in SUPPORTED_FUNCTIONS:

                f = SUPPORTED_FUNCTIONS[fun.upper()]
                return f + "(" + ",".join( [n.emit(ast, sheet_name=sheet_name) for n in args] ) + ")"

            else:
                message = "Function {} is not supported in xlcalcualtor.".format(fun.upper())
                logging.error(message)

    def __repr__(self):
        return "tvalue: {} ttype: {} tsubtype: {}".format(self.tvalue, self.ttype, self.tsubtype)

    def __str__(self):
        return self.__repr__()
