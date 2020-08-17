import sys
from functools import lru_cache

from xlcalculator.xlfunctions import xl, func_xltypes

from . import ast_nodes, xltypes


class EvaluatorContext(ast_nodes.EvalContext):

    def __init__(self, evaluator, ref):
        super().__init__(evaluator.namespace, ref)
        self.evaluator = evaluator

    @property
    def cells(self):
        return self.evaluator.model.cells

    @property
    def ranges(self):
        return self.evaluator.model.ranges

    @lru_cache(maxsize=None)
    def eval_cell(self, addr):
        # Check for a cycle.
        if addr in self.seen:
            raise RuntimeError(
                f'Cycle detected for {addr}:\n- ' + '\n- '.join(self.seen))
        self.seen.append(addr)

        return self.evaluator.evaluate(addr, self)


class Evaluator:
    """Traverses and evaluates a given model."""

    def __init__(self, model, namespace=None):
        self.model = model
        self.namespace = namespace \
            if namespace is not None else xl.FUNCTIONS.copy()
        self.cache_count = 0

    def _get_context(self, ref):
        return EvaluatorContext(self, ref)

    def resolve_names(self, addr):
        # Although defined names have been resolved in Model.create_node()
        # we need to attempt to resolve defined names as we might have been
        # given one in argument addr.
        if addr not in self.model.defined_names:
            return addr

        defn = self.model.defined_names[addr]

        if isinstance(defn, xltypes.XLCell):
            return defn.address

        if isinstance(defn, xltypes.XLRange):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"range and they aren't supported yet.")

        if isinstance(defn, xltypes.XLFormula):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"formula and they aren't supported as a cell "
                f"reference.")

    def evaluate(self, addr, context=None):
        # 1. Resolve the address to a cell.
        addr = self.resolve_names(addr)
        if addr not in self.model.cells:
            # Blank cell that has no stored value in the model.
            return func_xltypes.BLANK
        cell = self.model.cells[addr]

        # 2. If there is no formula, we simply return the cell value.
        if (cell.formula is None or cell.formula.evaluate is False):
            return func_xltypes.ExcelType.cast_from_native(
                self.model.cells[addr].value)

        # 3. Prepare the execution environment and evaluate the formula.
        #    (Note: Range nodes will automatically evaluate all their
        #           dependencies.)
        context = context if context is not None else self._get_context(addr)
        try:
            value = cell.formula.ast.eval(context)
        except Exception as err:
            raise RuntimeError(
                f"Problem evaluating cell {addr} formula "
                f"{cell.formula.formula}: {repr(err)}"
            ).with_traceback(sys.exc_info()[2])

        # 4. Update the cell value.
        #    Note for later: If an array is returned, we should distribute the
        #    values to the respective cell (known as spilling).
        cell.value = value
        cell.need_update = False

        return value

    def set_cell_value(self, address, value):
        """Sets the value of a cell in the model."""
        self.model.set_cell_value(address, value)

    def get_cell_value(self, address):
        """Gets the value of a cell in the model."""
        return self.model.get_cell_value(address)
