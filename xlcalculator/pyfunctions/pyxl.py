
from copy import deepcopy

from ..xlfunctions import xl

PYFUNCTIONS = deepcopy(xl.FUNCTIONS)


def registerpy(name=None):
    """Decorator to register a function in the PYFUNCTIONS namespace."""

    def registerFunction(func):
        PYFUNCTIONS.register(func, name)
        return func

    return registerFunction
