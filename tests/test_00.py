
import logging
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Model
from koala_xlcalculator import Evaluator

logging.basicConfig(level=logging.DEBUG)

json_file_name = 'test_00_graph.json'

filename = 'test_00.xlsm'
compiler = ModelCompiler()
new_model = compiler.parse_excel_file(filename)
new_model.persist_to_json_file(json_file_name)

reconstituted_model = Model()
reconstituted_model.construct_from_json_file(json_file_name)
# reconstituted_model.draw_graph()
reconstituted_model.build_code()

evaluator = Evaluator(reconstituted_model)
val1 = evaluator.evaluate('First!A2')
print("value 'evaluated' for First!A2 without a formula:", val1)
val2 = evaluator.evaluate('Seventh!C1')
print("value 'evaluated' for Seventh!C1 with a formula:", val2)
val3 = evaluator.evaluate('Ninth!B1')
print("value 'evaluated' for Ninth!B1 with a defined name:", val3)
val4 = evaluator.evaluate('Hundred')
print("value 'evaluated' for Hundred with a defined name:", val4)
val5 = evaluator.evaluate('Tenth!C1')
print("value 'evaluated' for Tenth!C1 with a defined name:", val5)
val6 = evaluator.evaluate('Tenth!C2')
print("value 'evaluated' for Tenth!C2 with a defined name:", val6)
val7 = evaluator.evaluate('Tenth!C3')
print("value 'evaluated' for Tenth!C3 with a defined name:", val7)
