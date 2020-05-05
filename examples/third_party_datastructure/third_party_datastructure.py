
input_dict = {
    "B4": 0.95,
    "B2": 1000,
    "B19": 0.001,
    "B20": 4,
    # B21
    "B22": 1,
    "B23": 2,
    "B24": 3,
    "B25": "=B2*B4",
    "B26": 5,
    "B27": 6,
    "B28": "=B19*B20*B22",
    "C22": "=SUM(B22:B28)",
  }

from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator

compiler = ModelCompiler()
my_model = compiler.read_and_parse_dict(input_dict)
evaluator = Evaluator(my_model)

for formula in my_model.formulae:
    print("Formula", formula, "evaluates to", evaluator.evaluate(formula))

# cells need a sheet and Sheet1 is default.
evaluator.set_cell_value("Sheet1!B22", 100)
print("Formula B28 now evaluates to", evaluator.evaluate("Sheet1!B28"))
print("Formula C22 now evaluates to", evaluator.evaluate("Sheet1!C22"))
