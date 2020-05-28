from .model import ModelCompiler, Model  # noqa
from .evaluator import Evaluator  # noqa

# Apply xlcalculator monkeypatches as soon as possible. This will ensure that
# all dependent modules that might import monkeypatched functions directly
# will get patched versions.
from xlcalculator import patch
patch.apply_all()

name = "xlcalculator"
