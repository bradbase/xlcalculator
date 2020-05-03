
import logging
from functools import lru_cache
from datetime import datetime

from pandas import DataFrame
from numpy import ndarray
from xlfunctions.exceptions import ExcelError
from xlfunctions import *

from ..xlcalculator_types import XLCell
from ..xlcalculator_types import XLRange


class Evaluator():
    """Traverses and evaluates a given model."""

    def __init__(self, model):

        self.model = model
        self.recursed_cells = set()
        self.cache_count = 0
        self.cells_to_evaluate = {}


    @staticmethod
    def recurse_evaluate(cells, data, cells_to_evaluate, recursed_cells, recurse_depth=0):
        """Sometimes (maybe most of the time) we need to 'chase' the dependency tree to eval the correct value of a cell."""

        if len(data) == 0:
            return

        # Base case
        if len(data) == 1 and isinstance(data[0], str):
            if data[0] in cells and cells[data[0]].formula is None:
                if recurse_depth in cells_to_evaluate:
                    cells_to_evaluate[recurse_depth].add( data[0] )
                else:
                    for level in cells_to_evaluate:
                        if data[0] in cells_to_evaluate[level]:
                            cells_to_evaluate[level].remove(data[0])
                    cells_to_evaluate[recurse_depth] = set(data)
                    recursed_cells.add(data[0])

        # Recursive cases
        elif len(data) == 1 and isinstance(data[0], str) and cells[data[0]].formula is not None:
            Evaluator.recurse_evaluate(cells, list( cells[data[0]].formula.associated_cells ), cells_to_evaluate, recursed_cells, recurse_depth=recurse_depth+1)
            if recurse_depth in cells_to_evaluate:
                cells_to_evaluate[recurse_depth].add( data[0] )
                recursed_cells.add(data[0])

            else:
                for level in cells_to_evaluate:
                    if data[0] in cells_to_evaluate[level]:
                        cells_to_evaluate[level].remove(data[0])
                cells_to_evaluate[recurse_depth] = set(data)
                recursed_cells.add(data[0])

        elif len(data) == 1 and isinstance(data[0], list):
            Evaluator.recurse_evaluate(cells, data, cells_to_evaluate, recursed_cells, recurse_depth=recurse_depth+1)

        else:
            mid = len(data) // 2
            first_half = data[:mid]
            second_half = data[mid:]
            Evaluator.recurse_evaluate(cells, first_half, cells_to_evaluate, recursed_cells, recurse_depth=recurse_depth)
            Evaluator.recurse_evaluate(cells, second_half, cells_to_evaluate, recursed_cells, recurse_depth=recurse_depth)


    @lru_cache(maxsize=None)
    def evaluate(self, cell_address, clear_cache=False):
        """Evaluates Python code as defined in formula.python_code"""

        defined_names = self.model.defined_names
        cells = self.model.cells
        ranges = self.model.cells

        if clear_cache:
            self.evaluate.cache_clear()
            self.eval_ref.cache_clear()

        try:
            # Although defined names have been resolved in Model.create_node()
            # we need to attempt to resolve defined names as we might have been
            # given one in argument cell_address.
            if cell_address in defined_names:
                name_definition = defined_names[cell_address]
                if isinstance(name_definition, XLCell):
                    cell_address = name_definition.address

                elif isinstance(name_definition, XLRange):
                    message = "I can't resolve {} to a cell. It's a range and they aren't supported yet.".format(cell)
                    logging.error(message)
                    raise Exception(message)

                elif isinstance(name_definition, XLFormula):
                    message = "I can't resolve {} to a cell. It's a formula and they aren't supported as a cell reference.".format(cell)
                    logging.error(message)
                    raise Exception(message)

            elif cell_address in ranges:
                cell = ranges[cell_address]

            else:
                cell = cells[cell_address]

        except:
            logging.error('Empty cell at {}'.format(cell_address))
            return ExcelError('#NULL', 'Cell {} is empty'.format(cell_address))

        # no formula, or no evaluation means fixed value but could be defined name
        if cells[cell_address].formula is None or cells[cell_address].formula.evaluate == False:
            logging.debug("\r\nCell {} has no formula \r\nbut its value is {}".format(cells[cell_address].address, cells[cell_address].value))
            return cells[cell_address].value

        try:
            if cells[cell_address].formula.python_code != None:
                skip = False
                for level in self.cells_to_evaluate:
                    if cell_address in self.cells_to_evaluate[level]:
                        skip = True

                if not skip:
                    recursed_cells = Evaluator.recurse_evaluate(cells, list( cells[cell_address].formula.associated_cells ), self.cells_to_evaluate, self.recursed_cells)
                    cells_to_evaluate_keys = list( self.cells_to_evaluate.keys() )
                    sorted(cells_to_evaluate_keys)
                    cells_to_evaluate_keys.reverse()
                    for item in cells_to_evaluate_keys:
                        for recursed_cell_address in self.cells_to_evaluate[item]:
                            active_cell = cells[recursed_cell_address]
                            if active_cell.formula is not None:
                                cells[recursed_cell_address].value = eval(active_cell.formula.python_code)

                value = eval(cells[cell_address].formula.python_code)
                if isinstance(value, Evaluator): # this should mean that vv is the result of RangeCore.apply_all, but with only one value inside
                    cells[cell_address].value = value.values[0]

                else:
                    if isinstance(value, ndarray):
                        cells[cell_address].value = value if len(value) != 0 else None
                    else:
                        if isinstance(value, DataFrame):
                            cells[cell_address].value = value if not value.empty else None
                        else:
                            cells[cell_address].value = value if value != '' else None

            else:
                cells[cell_address].value = 0

            cells[cell_address].need_update = False

        except Exception as exception:
            if str(exception).startswith("Problem evalling"):
                raise exception

            else:
                raise Exception("Problem evalling: {} for {}, {}".format(exception, cells[cell_address].address, cells[cell_address].formula.python_code))

        logging.debug("\r\nCell {} has a formula, {} \r\nwhich has been translated to Python as {} which evaluates to {}\r\n".format(cells[cell_address].address, cells[cell_address].formula.formula, cells[cell_address].formula.python_code, cells[cell_address].value))

        return cells[cell_address].value


    def eval_ref(self, address):
        """"""
        if address in self.model.cells:
            return self.model.cells[address].value
            # return self.model.cells[address]

        elif address in self.model.defined_names:
            # this is problematic as a defined name could be a cell, range or formula
            # TODO: support defined name to be cell, range and formula
            return DataFrame([self.model.defined_names[address]])

        elif address in self.model.ranges:
            range_cells = []
            for range_column in self.model.ranges[address].cells:
                row = []

                for cell_address in range_column:
                    row.append( self.model.cells[cell_address].value )

                range_cells.append(row)

            self.model.ranges[address].value = DataFrame(range_cells)
            return self.model.ranges[address].value


    @staticmethod
    def apply(func, first, second, ref=None):
        """"""
        # TODO: need to support ranges
        return Evaluator.apply_all(func, first, second)


    @staticmethod
    def apply_one(func, first, second, ref=None):
        """"""
        #TODO: This can't be called ATM, but needs to be - once we fully support ranges.
        function = SUPPORTED_OPERATORS[func]

        if ref is None:
            first_value = first
            second_value = second

        else:
            first_value = self.find_associated_value(ref, first)
            second_value = self.find_associated_value(ref, second)

        return function(first_value, second_value)


    @staticmethod
    def apply_all(func, first, second, ref=None):
        """"""

        function = SUPPORTED_OPERATORS[func]

        if isinstance(first, XLRange) and isinstance(second, XLRange):
            if first.length != second.length:
                return ExcelError('#VALUE!', 'apply_all must have 2 Ranges of identical length')

            vals = [function(
                x.value if isinstance(x, CellBase) else x,
                y.value if isinstance(x, CellBase) else y
            ) for x, y in zip(first.cells, second.cells)]

            return (first.addresses, vals, first.nrows, first.ncols)

        elif isinstance(first, XLRange):
            vals = [function(
                x.value if isinstance(x, CellBase) else x,
                second
            ) for x in first.cells]

            return (first.addresses, vals, first.nrows, first.ncols)

        elif isinstance(second, XLRange):
            vals = [function(
                first,
                x.value if isinstance(x, CellBase) else x
            ) for x in second.cells]

            return (second.addresses, vals, second.nrows, second.ncols)

        else:
            return function(first, second)


    @staticmethod
    def find_associated_cell(ref, range):
        """This function retrieves the cell associated to ref in a Range
        For instance, in the range [A1, B1, C1], the cell associated to B2 is B1
        This is useful to mimic the way Excel works.
        """

        if ref is not None:
            row, col = ref

            if (range.length) == 0:  # if a Range is empty, it means normally that all its cells are empty
                return None

            elif range.type == "vertical":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[
                        range.order.index((row, range.origin[1]))
                    ]

                else:
                    return None

            elif range.type == "horizontal":
                if (range.origin[0], col) in range.order:
                    return range.addresses[
                        range.order.index((range.origin[0], col))
                    ]

                else:
                    return None

            elif range.type == "scalar":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[
                        range.order.index((row, range.origin[1]))
                    ]

                elif (range.origin[0], col) in range.order:
                    return range.addresses[
                        range.order.index((range.origin[0], col))
                    ]

                elif (row, col) in range.order:
                    return range.addresses[
                        range.order.index((row, col))
                    ]

                else:
                    return None

            else:
                return None

        else:
            return None


    @staticmethod
    def find_associated_value(ref, item):
        """retrieves the value and not the Cell."""

        row, col = ref

        if isinstance(item, Evaluator):
            try:
                if (item.length) == 0:  # if a Range is empty, it means normally that all its cells are empty
                    item_value = 0

                elif item.type == "vertical":
                    if item.__cellmap is not None:
                        item_value = item[(row, item.origin[1])].value if (row, item.origin[1]) in item else None

                    else:
                        item_value = item[(row, item.origin[1])] if (row, item.origin[1]) in item else None

                elif item.type == "horizontal":

                    if item.__cellmap is not None:
                        try:

                            item_value = item[(item.origin[0], col)].value if (item.origin[0], col) in item else None

                        except:
                            raise Exception

                    else:
                        item_value = item[(item.origin[0], col)] if (item.origin[0], col) in item else None

                else:
                    return ExcelError('#VALUE!', 'cannot use find_associated_value on %s' % item.type)

            except ExcelError as e:
                raise Exception('First argument of Range operation is not valid: ' + e.value)

        elif item is None:
            item_value = 0

        else:
            item_value = item

        return item_value


    @staticmethod
    def check_value(a):
        """"""

        if isinstance(a, ExcelError):
            return a

        elif isinstance(a, str) and a in ErrorCodes:
            return ExcelError(a)

        try:  # This is to avoid None or Exception returned by Range operations
            if float(a) or isinstance(a, (str)):
                return a

            else:
                return 0

        except:
            if a == 'True':
                return True

            elif a == 'False':
                return False

            else:
                return 0


    @staticmethod
    def add(lhs, rhs):
        """"""

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:
            return Evaluator.check_value(lhs) + Evaluator.check_value(rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def substract(lhs, rhs):
        """"""

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:

            return Evaluator.check_value(lhs) - Evaluator.check_value(rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def minus(a, b=None):
        """"""

        # b is not used, but needed in the signature. Maybe could be better
        try:
            if isinstance(a, XLCell):
                a = a.value

            return a * -1

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def multiply(lhs, rhs):
        """"""

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:

            return Evaluator.check_value( lhs ) * Evaluator.check_value( rhs )

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def divide(numerator, denominator):
        """"""

        if isinstance(numerator, XLCell):
            numerator = numerator.value

        if isinstance(denominator, XLCell):
            denominator = denominator.value

        if denominator == 0:
            return ExcelError('#DIV/0!', e)

        return numerator / denominator


    @staticmethod
    def is_strictly_superior(lhs, rhs):

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:
            return Evaluator.check_value(lhs) > Evaluator.check_value(rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_strictly_inferior(lhs, rhs):

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:
            return Evaluator.check_value(lhs) < Evaluator.check_value(rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_superior_or_equal(lhs, rhs):

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:
            lhs = Evaluator.check_value(lhs)
            rhs = Evaluator.check_value(rhs)

            return a > b or Evaluator.is_almost_equal(lhs, rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_inferior_or_equal(lhs, rhs):

        if isinstance(lhs, XLCell):
            lhs = lhs.value

        if isinstance(rhs, XLCell):
            rhs = rhs.value

        try:
            lhs = Evaluator.check_value(lhs)
            rhs = Evaluator.check_value(rhs)

            return a < b or Evaluator.is_almost_equal(lhs, rhs)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_equal(a, b):
        """"""

        try:
            if not isinstance(a, (str)):
                a = Evaluator.check_value(a)

            if not isinstance(b, (str)):
                b = Evaluator.check_value(b)

            return Evaluator.is_almost_equal(a, b, precision=0.00001)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_not_equal(a, b):
        """"""

        try:
            if not isinstance(a, (str)):
                a = Evaluator.check_value(a)

            if not isinstance(a, (str)):
                b = Evaluator.check_value(b)

            return a != b

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_number(s): # http://stackoverflow.com/questions/354038/how-do-i-check-if-a-string-is-a-number-float-in-python
        try:
            float(s)
            return True

        except:
            return False


    @staticmethod
    def is_almost_equal(a, b, precision = 0.0001):
        if Evaluator.is_number(a) and Evaluator.is_number(b):
            return abs(float(a) - float(b)) <= precision

        elif (a is None or a == 'None') and (b is None or b == 'None'):
            return True

        else: # booleans or strings
            return str(a) == str(b)


    def set_cell_value(self, address, value, clear_cache=False):
        """Sets the value of a cell in the model."""
        self.model.set_cell_value(address, value)
        if clear_cache:
            Evaluator.evaluate.cache_clear()


    def get_cell_value(self, address):
        """Gets the value of a cell in the model."""
        return self.model.get_cell_value(address)



SUPPORTED_OPERATORS = {
    "minus": Evaluator.minus,
    "add": Evaluator.add,
    "substract": Evaluator.substract,
    "multiply": Evaluator.multiply,
    "divide": Evaluator.divide,
    "is_equal": Evaluator.is_equal,
    "is_not_equal": Evaluator.is_not_equal,
    "is_strictly_superior": Evaluator.is_strictly_superior,
    "is_strictly_inferior": Evaluator.is_strictly_inferior,
    "is_superior_or_equal": Evaluator.is_superior_or_equal,
    "is_inferior_or_equal": Evaluator.is_inferior_or_equal,
}
