
import logging
from dataclasses import dataclass, field
import json
import gzip
from collections import deque

from networkx import DiGraph
import matplotlib.pyplot as plt
from jsonpickle import encode, decode

from ..koala_types import XLCell
from ..koala_types import XLFormula
from ..koala_types import XLRange
from ..koala_types import RangeNode
from ..koala_types import OperandNode
from ..koala_types import OperatorNode
from ..koala_types import FunctionNode
from ..read_excel import f_token
from ..read_excel import ExcelParser


# def init_graph():
#     """Default factory to initialise Formula.ranges."""
#     return DiGraph()


def init_dict():
    """Default factory to initialise Formula.ranges."""

    return {}


class Operator(object):
    """Small wrapper class to manage operators during shunting yard"""

    def __init__(self, value, precedence, associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity


@dataclass
class Model():

    # graph: DiGraph = field(init=False, default_factory=init_graph, compare=True, hash=True, repr=True)
    cells: dict = field(init=False, default_factory=init_dict, compare=True, hash=True, repr=True)
    defined_names: dict = field(init=False, default_factory=init_dict, compare=True, hash=True, repr=True)
    formulae: dict = field(init=False, default_factory=init_dict, compare=True, hash=True, repr=True)
    ranges: dict = field(init=False, default_factory=init_dict, compare=True, hash=True, repr=True)


    # def draw_graph(self):
    #     """Draw a graphical representation of the graph."""
    #
    #     drawing = nx.draw(self.graph)
    #     plt.show()

    def set_cell_value(self, address, value):
        """Sets a new value for a specified cell."""
        if isinstance(address, str):
            if address in self.cells:
                self.cells[address].value = value
            else:
                self.cells[address] = XLCell(address, value)

        elif isinstance(address, XLCell):
            if address.address in self.cells:
                self.cells[address.address].value = value
            else:
                self.cells[address.address] = XLCell(address.address, value)

        else:
            raise Exception("I can't set the cell value for an address of type {}, I need XLCell or a string".format(type(address)))


    def get_cell_value(self, address):
        if isinstance(address, str):
            if address in self.cells:
                return self.cells[address].value
            else:
                logging.debug("Trying to get value for cell {} but that cell doesn't exist".format(address))
                return 0

        elif isinstance(address, XLCell):
            if address.address in self.cells:
                return self.cells[address.address].value
            else:
                logging.debug("Trying to get value for cell {} but that cell doesn't exist".format(address.address))
                return 0

        else:
            raise Exception("I can't get the cell value for an address of type {}, I need XLCell or a string".format(type(address)))


    def persist_to_json_file(self, fname):
        """Writes the state to disk. Doesn't write the graph directly, but persist all the things that provide the ability to re-create the graph."""

        output = {
            'cells' : self.cells,
            'defined_names' : self.defined_names,
            'formulae' : self.formulae,
            'ranges' : self.ranges,
        }

        if fname.split('.')[-1:][0].upper() in ['GZIP', 'GZ']:
            outfile = gzip.GzipFile(fname,'wb')
            outfile.write( str.encode( encode(output, keys=True) ) )
        else:
            outfile = open(fname, "w")
            outfile.write( encode(output, keys=True) )

        outfile.close()


    def construct_from_json_file(self, fname):
        """Constructs a graph from a state persisted to disk."""

        if fname.split('.')[-1:][0].upper() in ['GZIP', 'GZ']:
            infile = gzip.GzipFile(fname,'rb')
        else:
            infile = open(fname, "rb")

        json_bytes = infile.read()
        infile.close()
        data = decode(json_bytes, keys=True, classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        # for cell_key in self.cells.keys():
        #     if self.cells[cell_key].formula is not None:
        #         self.formulae[cell_key] = self.cells[cell_key].formula

        # for formula in self.formulae:
        #     for range in self.formulae[formula].ranges:
        #         if ":" in range:
        #             self.ranges[range] = XLRange(range, range)
        #         else:
        #             if range not in self.defined_names:
        #                 self.ranges[range] = XLCell(range)
        # logging.debug("Summary of the cells converted to a network graph:\r\n%s" % nx.info(self.graph))


    # def translate(self, outputs = [], inputs = []):
    #     """Translates a Microsoft Excel cell structure into a model representation."""
    #
    #     # We need the defined names in the graph so we can "back-link"/associate the cells to them
    #     for name in self.defined_names:
    #         self.graph.add_node(self.defined_names[name])
    #
    #     # add all the cells to the graph
    #     for cell in self.cells.values():
    #         self.graph.add_node(cell)
    #
    #         # define the "back-link"/association between the cell and the defined name
    #         if cell.defined_names != []:
    #             for name in cell.defined_names:
    #                 self.graph.add_edge(cell, self.defined_names[name])
    #
    #         # Associate cells if they are involved with a calculation (formula)
    #         if cell.formula is not None:
    #             has_infix_operator = False
    #             for token in cell.formula.tokens:
    #                 if token.ttype in ['operator-infix']:
    #                     has_infix_operator = True
    #
    #             if has_infix_operator:
    #                 for range in cell.formula.ranges:
    #                     if range in self.cells.keys():
    #                         self.graph.add_edge(cell, self.cells[range])
    #
    #                     else:
    #                         if ':' in range:
    #                             message = "Will need to resolve the range. {} referenced from cell {}".format(range, cell.address)
    #                             logging.error( message )
    #                             raise Exception( message )
    #
    #                         else:
    #                             potential_cell = "{}!{}".format(cell.sheet, range)
    #                             if potential_cell in self.cells.keys():
    #                                 self.graph.add_edge(cell, potential_cell)
    #
    #                             else:
    #                                 message = "I can't seem to find this cell: {}".format(potential_cell)
    #                                 logging.error( message )
    #                                 raise Exception( message )


    def build_code(self):
        """Define the Python code for all cells in the dict of cells."""

        for cell in self.cells:
            if self.cells[cell].formula is not None:
                sheet_name = self.cells[cell].sheet
                tokenized_formula = self.shunting_yard(self.cells[cell].formula.formula, self.defined_names.keys(), sheet_name=sheet_name, ref=None, tokenize_range=False)
                ast, root = self.build_ast(tokenized_formula)
                self.cells[cell].formula.python_code = root.emit(ast, sheet_name=sheet_name)


    def shunting_yard(self, expression, named_ranges, sheet_name=None, ref=None, tokenize_range=False):
        """
        Tokenize an excel formula expression into reverse polish notation

        Core algorithm taken from wikipedia with varargs extensions from
        http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-algorithm-to-allow-variable-numbers-of-arguments-to-functions/


        The ref is the cell address which is passed down to the actual compiled python code.
        Range basic operations signature require this reference, so it has to be written during OperatorNode.emit()
        https://github.com/iOiurson/koala/blob/master/koala/ast/graph.py#L292.

        This is needed because Excel range basic operations (+, -, * ...) are applied on matching cells.

        Example:
        Cell C2 has the following formula 'A1:A3 + B1:B3'.
        The output will actually be A2 + B2, because the formula is relative to cell C2.
        """

        #remove leading =
        if expression.startswith('='):
            expression = expression[1:]

        excel_parser = ExcelParser(tokenize_range = tokenize_range);
        excel_parser.parse(expression, sheet_name=sheet_name)

        # insert tokens for '(' and ')', to make things clearer below
        tokens = []
        for token in excel_parser.tokens.items:
            if token.ttype == "function" and token.tsubtype == "start":
                token.tsubtype = ""
                tokens.append(token)
                tokens.append(f_token('(','arglist','start'))

            elif token.ttype == "function" and token.tsubtype == "stop":
                tokens.append(f_token(')','arglist','stop'))

            elif token.ttype == "subexpression" and token.tsubtype == "start":
                token.tvalue = '('
                tokens.append(token)

            elif token.ttype == "subexpression" and token.tsubtype == "stop":
                token.tvalue = ')'
                tokens.append(token)

            elif token.ttype == "operand" and token.tsubtype == "range" and token.tvalue in named_ranges:
                token.tsubtype = "named_range"
                tokens.append(token)

            else:
                tokens.append(token)

        #http://office.microsoft.com/en-us/excel-help/calculation-operators-and-precedence-HP010078886.aspx
        operators = {}
        operators[':'] = Operator(':',8,'left')
        operators[''] = Operator(' ',8,'left')
        operators[','] = Operator(',',8,'left')
        operators['u-'] = Operator('u-',7,'left') #unary negation
        operators['%'] = Operator('%',6,'left')
        operators['^'] = Operator('^',5,'left')
        operators['*'] = Operator('*',4,'left')
        operators['/'] = Operator('/',4,'left')
        operators['+'] = Operator('+',3,'left')
        operators['-'] = Operator('-',3,'left')
        operators['&'] = Operator('&',2,'left')
        operators['='] = Operator('=',1,'left')
        operators['<'] = Operator('<',1,'left')
        operators['>'] = Operator('>',1,'left')
        operators['<='] = Operator('<=',1,'left')
        operators['>='] = Operator('>=',1,'left')
        operators['<>'] = Operator('<>',1,'left')

        output = deque()
        stack = []
        were_values = []
        arg_count = []

        new_tokens = []

        # reconstruct expressions with ':' and replace the corresponding tokens by the reconstructed expression
        if not tokenize_range:
            for index, token in enumerate(tokens):
                new_tokens.append(token)

                if type(token.tvalue) == str:

                    if token.tvalue.startswith(':'): # example -> :OFFSET( or simply :A10
                        depth = 0
                        expr = ''
                        rev = reversed(tokens[:index])

                        for reversed_token in rev: # going backwards, 'stop' starts, 'start' stops
                            if reversed_token.tsubtype == 'stop':
                                depth += 1

                            elif depth > 0 and reversed_token.tsubtype == 'start':
                                depth -= 1

                            expr = reversed_token.tvalue + expr

                            new_tokens.pop()

                            if depth == 0:
                                new_tokens.pop() # these 2 lines are needed to remove INDEX()
                                new_tokens.pop()
                                expr = rev.next().tvalue + expr
                                break

                        expr += token.tvalue

                        depth = 0

                        if token.tvalue[1:] in ['OFFSET', 'INDEX']:
                            for t in tokens[(index + 1):]:
                                if t.tsubtype == 'start':
                                    depth += 1

                                elif depth > 0 and t.tsubtype == 'stop':
                                    depth -= 1

                                expr += t.tvalue
                                tokens.remove(t)

                                if depth == 0:
                                    break

                        new_tokens.append( f_token(expr, 'operand', 'pointer') )

                    elif ':OFFSET' in token.tvalue or ':INDEX' in token.tvalue: # example -> A1:OFFSET(
                        depth = 0
                        expr = ''
                        expr += token.tvalue

                        for t in tokens[(index + 1):]:
                            if t.tsubtype == 'start':
                                depth += 1

                            elif t.tsubtype == 'stop':
                                depth -= 1

                            expr += t.tvalue
                            tokens.remove(t)

                            if depth == 0:
                                new_tokens.pop()
                                break

                        new_tokens.append(f_token(expr, 'operand', 'pointer'))

        tokens = new_tokens if new_tokens else tokens

        for token in tokens:
            if token.ttype == "operand":
                output.append(self.create_node(token, sheet_name=sheet_name, ref=ref))

                if were_values:
                    were_values.pop()
                    were_values.append(True)

            elif token.ttype == "function":
                stack.append(token)
                arg_count.append(0)

                if were_values:
                    were_values.pop()
                    were_values.append(True)

                were_values.append(False)

            elif token.ttype == "argument":

                while stack and (stack[-1].tsubtype != "start"):
                    output.append(self.create_node(stack.pop(), sheet_name=sheet_name, ref=ref))

                if were_values.pop(): arg_count[-1] += 1
                were_values.append(False)

                if not len(stack):
                    message = "Mismatched or misplaced parentheses"
                    logging.error(message)
                    raise Exception(message)

            elif token.ttype.startswith('operator'):

                if token.ttype.endswith('-prefix') and token.tvalue =="-":
                    o1 = operators['u-']

                else:
                    o1 = operators[token.tvalue]

                while stack and stack[-1].ttype.startswith('operator'):
                    if stack[-1].ttype.endswith('-prefix') and stack[-1].tvalue =="-":
                        o2 = operators['u-']

                    else:
                        o2 = operators[stack[-1].tvalue]

                    if ( (o1.associativity == "left" and o1.precedence <= o2.precedence) or (o1.associativity == "right" and o1.precedence < o2.precedence) ):
                        output.append(self.create_node(stack.pop(), sheet_name=sheet_name, ref=ref))

                    else:
                        break

                stack.append(token)

            elif token.tsubtype == "start":
                stack.append(token)

            elif token.tsubtype == "stop":
                while stack and stack[-1].tsubtype != "start":
                    output.append(self.create_node(stack.pop(), sheet_name=sheet_name, ref=ref))

                if not stack:
                    raise Exception("Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].ttype == "function":
                    f = self.create_node(stack.pop(), sheet_name=sheet_name, ref=ref)
                    a = arg_count.pop()
                    w = were_values.pop()
                    if w: a += 1
                    f.num_args = a
                    output.append(f)

        while stack:
            if (stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop"):
                raise Exception("Mismatched or misplaced parentheses")

            output.append(self.create_node(stack.pop(), sheet_name=sheet_name, ref=ref))

        # convert to list
        return [x for x in output]


    def build_ast(self, expression):
        """Build an AST from an Excel formula expression in reverse polish notation."""

        #use a directed graph to store the tree
        ast_graph = DiGraph()
        stack = []

        for n in expression:
            # Since the graph does not maintain the order of adding nodes/edges
            # add an extra attribute 'pos' so we can always sort to the correct order
            if isinstance(n, OperatorNode):
                if n.ttype == "operator-infix":
                    arg2 = stack.pop()
                    arg1 = stack.pop()
                    # Hack to write the name of sheet in 2argument address
                    if(n.tvalue == ':'):
                        if '!' in arg1.tvalue and arg2.ttype == 'operand' and '!' not in arg2.tvalue:
                            arg2.tvalue = arg1.tvalue.split('!')[0] + '!' + arg2.tvalue

                    ast_graph.add_node(arg1, pos = 0)
                    ast_graph.add_node(arg2, pos = 1)
                    ast_graph.add_edge(arg1, n)
                    ast_graph.add_edge(arg2, n)

                else:
                    arg1 = stack.pop()
                    ast_graph.add_node(arg1, pos = 1)
                    ast_graph.add_edge(arg1, n)

            elif isinstance(n, FunctionNode):
                args = []
                for _ in range(n.num_args):
                    try:
                        args.append(stack.pop())

                    except:
                        raise Exception()

                args.reverse()
                for i, a in enumerate(args):
                    ast_graph.add_node(a, pos = i)
                    ast_graph.add_edge(a, n)
            else:
                ast_graph.add_node(n, pos=0)

            stack.append(n)

        return ast_graph, stack.pop()


    def create_node(self, t, sheet_name=None, ref=None):
        """Simple factory function"""

        if t.ttype == "operand":
            if t.tsubtype in ["range"]:
                return RangeNode(t, ref)

            if t.tsubtype in ["pointer"]:
                print("FOUND A POINTER!!!", t.tvalue)
                return RangeNode(t, ref)

            elif t.tsubtype in ["named_range"] :
                # we need to attempt resolving defined names here
                # so we can persist formulas without defined names
                if t.tvalue in self.defined_names:
                    name_definition = self.defined_names[t.tvalue]

                    if isinstance(name_definition, XLCell):
                        t.tvalue = name_definition.address

                    elif isinstance(name_definition, XLRange):
                        t.tvalue = name_definition.name
                        message = "{} is a range, which is not yet supported".format(name_definition)
                        logging.error(message)
                        raise Exception(message)

                    else:
                        message = "{} is of type {}, which is not yet supported".format(name_definition, type(name_definition))
                        logging.error(message)
                        raise Exception(message)

                return RangeNode(t, ref)

            else:
                return OperandNode(t)

        elif t.ttype == "function":
            return FunctionNode(t, ref)

        elif t.ttype.startswith("operator"):
            return OperatorNode(t, ref)

        else:
            return ASTNode(t)


    def __eq__(self, other):

        cells_comparison = []
        for self_cell in self.cells:
            cells_comparison.append( self.cells[self_cell] == other.cells[self_cell] )

        defined_names_comparison = []
        for self_defined_names in self.defined_names:
            defined_names_comparison.append( self.defined_names[self_defined_names] == other.defined_names[self_defined_names] )

        return self.__class__ == other.__class__ and all(cells_comparison) and all(defined_names_comparison)
