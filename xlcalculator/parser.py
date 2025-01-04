from . import ast_nodes, tokenizer


class Operator(object):
    """Small wrapper class to manage operators during shunting yard"""

    def __init__(self, value, precedence, associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity


# TODO: why is this like the third time that the operators are defined?

# http://office.microsoft.com/en-us/excel-help/
#    calculation-operators-and-precedence-HP010078886.aspx
OPERATORS = {
    ":": Operator(":", 8, "left"),
    "": Operator(" ", 8, "left"),
    ",": Operator(",", 8, "left"),
    "u-": Operator("u-", 7, "right"),  # unary negation
    "%": Operator("%", 6, "left"),
    "^": Operator("^", 5, "left"),
    "*": Operator("*", 4, "left"),
    "/": Operator("/", 4, "left"),
    "+": Operator("+", 3, "left"),
    "-": Operator("-", 3, "left"),
    "&": Operator("&", 2, "left"),
    "=": Operator("=", 1, "left"),
    "<": Operator("<", 1, "left"),
    ">": Operator(">", 1, "left"),
    "<=": Operator("<=", 1, "left"),
    ">=": Operator(">=", 1, "left"),
    "<>": Operator("<>", 1, "left"),
}


class FormulaParser:
    """Excel Formula Parser"""

    # TODO: Changed the default for named ranges to something that dons't crash
    # TODO: deal with unused variables
    def parse(self, formula: str, named_ranges=None, tokenize_range=False):
        """Parse formula into evaluable AST."""
        # 1. Parse the formula into syntactic tokens.
        tokens = self.tokenize(formula)
        # 2. Organize tokens into reverse polish notation.
        nodes = self.shunting_yard(tokens, named_ranges)
        # 3. Construct the AST by connecting the AST nodes.
        ast = self.build_ast(nodes)
        return ast

    def tokenize(
        self, formula: str, tokenize_range: bool = False
    ) -> list[tokenizer.f_token]:
        # Remove leading "=" sign.
        if formula.startswith("="):
            formula = formula[1:]

        excel_parser = tokenizer.ExcelParser(tokenize_range=tokenize_range)
        return excel_parser.parse(formula).items

    def shunting_yard(
        self,
        raw_tokens: list[tokenizer.f_token],
        named_ranges: list[str],
        tokenize_range: bool = False,
    ) -> list[ast_nodes.ASTNode]:
        """Reorganize tokens into proper reverse polish notation.

        Core algorithm taken from wikipedia with varargs extensions from

        http://www.kallisti.net.nz/blog/2008/02/
            extension-to-the-shunting-yard-algorithm-to-allow-
            variable-numbers-of-arguments-to-functions/


        The ref is the cell address which is passed down to the actual
        compiled python code.  Range basic operations signature require this
        reference, so it has to be written during OperatorNode.emit()

          https://github.com/iOiurson/koala/blob/master/koala/ast/graph.py#L292.

        This is needed because Excel range basic operations (+, -, * ...) are
        applied on matching cells.

        Example:

        Cell C2 has the following formula 'A1:A3 + B1:B3'.  The output will
        actually be A2 + B2, because the formula is relative to cell C2.
        """
        # insert tokens for '(' and ')', to make things clearer below
        tokens = []
        for token in raw_tokens:
            if token.ttype == "function" and token.tsubtype == "start":
                token.tsubtype = ""
                tokens.append(token)
                tokens.append(tokenizer.f_token("(", "arglist", "start"))

            elif token.ttype == "function" and token.tsubtype == "stop":
                tokens.append(tokenizer.f_token(")", "arglist", "stop"))

            elif token.ttype == "subexpression" and token.tsubtype == "start":
                token.tvalue = "("
                tokens.append(token)

            elif token.ttype == "subexpression" and token.tsubtype == "stop":
                token.tvalue = ")"
                tokens.append(token)

            elif (token.ttype == "operand" and token.tsubtype == "range" and token.tvalue in named_ranges):
                # Resolve the named range once and for all.
                token.tvalue = named_ranges[token.tvalue]
                tokens.append(token)

            else:
                tokens.append(token)

        output = []
        stack = []
        were_values = []
        arg_count = []

        new_tokens = []

        # reconstruct expressions with ':' and replace the corresponding
        # tokens by the reconstructed expression
        if not tokenize_range:
            for index, token in enumerate(tokens):
                new_tokens.append(token)

                if type(token.tvalue) is str:
                    # example -> :OFFSET( or simply :A10
                    if token.tvalue.startswith(":"):
                        depth = 0
                        expr = ""
                        rev = reversed(tokens[:index])

                        # going backwards, 'stop' starts, 'start' stops
                        for reversed_token in rev:
                            if reversed_token.tsubtype == "stop":
                                depth += 1

                            elif depth > 0 and reversed_token.tsubtype == "start":
                                depth -= 1

                            expr = reversed_token.tvalue + expr

                            new_tokens.pop()

                            if depth == 0:
                                # these 2 lines are needed to remove INDEX()
                                new_tokens.pop()
                                new_tokens.pop()
                                expr = rev.next().tvalue + expr
                                break

                        expr += token.tvalue

                        depth = 0

                        if token.tvalue[1:] in ["OFFSET", "INDEX"]:
                            for t in tokens[(index + 1):]:
                                if t.tsubtype == "start":
                                    depth += 1

                                elif depth > 0 and t.tsubtype == "stop":
                                    depth -= 1

                                expr += t.tvalue
                                tokens.remove(t)

                                if depth == 0:
                                    break

                        new_tokens.append(tokenizer.f_token(expr, "operand", "pointer"))

                    # example -> A1:OFFSET(
                    elif ":OFFSET" in token.tvalue or ":INDEX" in token.tvalue:
                        depth = 0
                        expr = ""
                        expr += token.tvalue

                        for t in tokens[(index + 1):]:
                            if t.tsubtype == "start":
                                depth += 1

                            elif t.tsubtype == "stop":
                                depth -= 1

                            expr += t.tvalue
                            tokens.remove(t)

                            if depth == 0:
                                new_tokens.pop()
                                break

                        new_tokens.append(tokenizer.f_token(expr, "operand", "pointer"))

        tokens = new_tokens if new_tokens else tokens

        for token in tokens:
            if token.ttype == "operand":
                output.append(self.create_node(token))

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
                    output.append(self.create_node(stack.pop()))

                if were_values.pop():
                    arg_count[-1] += 1
                were_values.append(False)

                if not len(stack):
                    raise ValueError("Mismatched or misplaced parentheses")

            elif token.ttype.startswith("operator"):
                if token.ttype.endswith("-prefix") and token.tvalue == "-":
                    o1 = OPERATORS["u-"]

                else:
                    o1 = OPERATORS[token.tvalue]

                while stack and stack[-1].ttype.startswith("operator"):
                    if stack[-1].ttype.endswith("-prefix") and stack[-1].tvalue == "-":
                        o2 = OPERATORS["u-"]

                    else:
                        o2 = OPERATORS[stack[-1].tvalue]

                    if (
                        o1.associativity == "left" and o1.precedence <= o2.precedence
                    ) or (
                        o1.associativity == "right" and o1.precedence < o2.precedence
                    ):
                        output.append(self.create_node(stack.pop()))

                    else:
                        break

                stack.append(token)

            elif token.tsubtype == "start":
                stack.append(token)

            elif token.tsubtype == "stop":
                while stack and stack[-1].tsubtype != "start":
                    output.append(self.create_node(stack.pop()))

                if not stack:
                    raise SyntaxError("Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].ttype == "function":
                    f = self.create_node(stack.pop())
                    a = arg_count.pop()
                    w = were_values.pop()
                    if w:
                        a += 1
                    f.num_args = a
                    output.append(f)

        while stack:
            if stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop":
                raise SyntaxError("Mismatched or misplaced parentheses")

            output.append(self.create_node(stack.pop()))

        # convert to list
        return [x for x in output]

    def create_node(self, token: tokenizer.f_token) -> ast_nodes.ASTNode:
        if token.ttype == "operand":
            if token.tsubtype in ["range", "pointer"]:
                return ast_nodes.RangeNode(token)
            else:
                return ast_nodes.OperandNode(token)

        elif token.ttype == "function":
            return ast_nodes.FunctionNode(token)

        elif token.ttype.startswith("operator"):
            return ast_nodes.OperatorNode(token)

        else:
            raise ValueError("Unknown token type: " + token.ttype)

    def build_ast(self, nodes: list[ast_nodes.ASTNode]) -> ast_nodes.ASTNode:
        """Update AST nodes to build a proper parse tree.

        XXX: There is really no need for this. The shunting yeard algorithm
        should jsut take care of it.
        """
        stack: list[ast_nodes.ASTNode] = []

        for node in nodes:
            if isinstance(node, ast_nodes.OperatorNode):
                if node.ttype == "operator-infix":
                    # Stack has arguments in reverse order.
                    node.right = stack.pop()
                    node.left = stack.pop()
                else:
                    node.right = stack.pop()

            elif isinstance(node, ast_nodes.FunctionNode):
                args = []
                for _ in range(node.num_args):
                    args.append(stack.pop())
                # Stack has arguments in reverse order.
                node.args = list(reversed(args))

            stack.append(node)

        return stack.pop()
