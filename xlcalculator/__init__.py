name = "xlcalcualtor"
from .model import ModelCompiler
from .model import Model
from .evaluator import Evaluator

# Apply xlcalculator monkeypatches as soon as possible. This will ensure that
# all dependent modules that might import monkeypatched functions directly
# will get patched versions.
from xlcalculator import patch
patch.apply_all()
