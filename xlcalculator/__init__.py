from .model import ModelCompiler, Model  # noqa
from .evaluator import Evaluator  # noqa


from .xlfunctions.xl import FUNCTIONS, register  # noqa: F401
from .xlfunctions.xlerrors import *  # noqa: F401, F403
from .xlfunctions.func_xltypes import *  # noqa: F401, F403

# Make sure to register all functions
from .xlfunctions import (  # noqa: F401
    date,
    financial,
    information,
    logical,
    lookup,
    math,
    operator,
    statistics,
    text
)


name = "xlcalculator"
