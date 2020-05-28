import unittest

from xlcalculator.model import parser


class FormulaParserTest(unittest.TestCase):

    def parse(self, formula, named_ranges=None):
        if named_ranges is None:
            named_ranges = {
                'first': 'A1',
                'two_by_two': 'A1:B2'
            }
        ast = parser.FormulaParser().parse(formula, named_ranges)
        return ast

    def test_parse(self):
        self.assertEqual(str(self.parse('=A1+1')), '(A1) + (1)')

    def test_parse_without_equal(self):
        self.assertEqual(str(self.parse('A1+1')), '(A1) + (1)')

    def test_parse_with_named_range(self):
        # Named ranges get immediately resolved.
        self.assertEqual(str(self.parse('first+1')), '(A1) + (1)')

    def test_parse_function(self):
        self.assertEqual(str(self.parse('SUM(A1:B2)')), 'SUM(A1:B2)')

    def test_parse_function_with_arg_expr(self):
        self.assertEqual(
            str(self.parse('SUM(1, 2, (3+4))')), 'SUM(1, 2, (3) + (4))')

    def test_parse_nested_function(self):
        self.assertEqual(
            str(self.parse('MOD(SUM(A1:B2), 2)')),
            'MOD(SUM(A1:B2), 2)')

    def test_parse_with_parens(self):
        self.assertEqual(str(self.parse('2*(A1+1)')), '(2) * ((A1) + (1))')

    def test_parse_prefix_op(self):
        self.assertEqual(str(self.parse('-(A1+1)')), '- ((A1) + (1))')

    def test_parse_percent(self):
        # Should be a proper postfix op, but tokenizer takes care of it right
        # away.
        self.assertEqual(str(self.parse('(A1+1)%')), '((A1) + (1)) * (0.01)')

    def test_parse_open_range_start(self):
        self.assertEqual(str(self.parse(':B2')), ':B2')

    def test_parse_open_range_end(self):
        self.assertEqual(str(self.parse('A1:')), 'A1:')

    @unittest.skip(
        'Not properly supported during shunting. Does not remove func node.')
    def test_parse_with_offser(self):
        self.assertEqual(str(self.parse(':OFFSET(A1, 1, 1)')), ':B2')

    def test_parse_with_parens_mismatched(self):
        with self.assertRaises(SyntaxError):
            self.parse('2*(A1+1')
        with self.assertRaises(IndexError):
            self.parse('2*)')
