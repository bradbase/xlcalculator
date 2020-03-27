
# Koala Excel Calculator

In short Koala_xlcalculator is an attempted re-write of the [koala2](https://github.com/vallettea/koala) library with intent to modernize the code, increase maintainability and keep compatibility with external software solutions (self-indulgently one is - [FlyingKoala](https://github.com/bradbase/flyingkoala) ).

Koala2's heritage has been "spreadsheet replacement" in order to get performant calculation. In essence replacing the calculation engine in Excel with one which works faster. And, all credit to the civic hacker crew at [we are ants](https://weareants.fr/#!/koala-the-faster-excel), they absolutely succeeded. In various iterations of the project since it appears as though the koala2 has tried moving toward the idea of something like openpyxl with a "calculate" button.

This is not the direction I've taken.

Koala2 had a few features which enabled me to use koala2 quite differently. I admit these features may not have been core to the wider audience. The features I am specifically interested in are the "advanced" features of setting output cells on load and saving/loading the state of the Python representation of the worksheet. With setting the output cell on load it was possible to load only one formula, and/or the related parent formulas and the input cells involved with the network of formulas. In essence focusing the resulting "loaded Excel" to far less than the entire workbook. And, once in Python, it could be persisted (saved) and re-loaded. This created a mathematical model defined in the language of Excel formulas but executable ("evaluated") and iterable in Python without the need for a person to translate it.

This use case provides the ability to view Excel workbooks as definitions of a calculation (a model or part thereof). Where the model is an abstraction that just so happens to be expressed in the form of an Excel spreadsheet. So, in that sense, koala2 wasn't simply replacing the Excel calculation engine but was more a toolkit which uses Excel as an interface for defining parts or whole mathematical systems. I saw that as very useful.

These features I so covet have broken after some (much needed) code cleanup and I've struggled to re-implement them. With that, I have decided to use some parts of koala2 and implement a library which will have the features I'm interested in. I'm happy if this implementation (or something similar) becomes adopted by the koala2 project (I'd prefer to not split a universe) but, equally, I accept my goals may not overlap enough with that project.

Moving forward, this project, if used in one way, would achieve the purposes of replacing the Excel calculation engine or even be analogous to openpyxl with a "go" button. But would also offer significantly more for those who are inextricably bound to Excel but need more from their math. Which is what FlyingKoala is intended to facilitate.

We have some very basic functionality working. That said, the project is still evolving and has some way to go before it can claim to be feature compatible with koala2.

koala_xlcalculator currently supports:
* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Evaluating
  * Individual cells
    * no mathematical operation,
    * address like Sheet1!A1
    * essentially just returns the value of the cell
  * Defined Name (a "named cell" or range)
    * Returns the value of a cell referenced by the name
  * Ranges
    * Evaluation of a function which has had a range passed to it
  * Operands (+, -, /, \*, ==, <>, <=, >=)
    * on cells only
  * Functions
    * AVERAGE
    * CHOOSE
    * CONCAT
    * COUNT
    * COUNTA
    * DATE
    * MAX
    * MID
    * MIN
    * POWER
    * ROUND
    * ROUNDDOWN
    * ROUNDUP
    * SUM
  * Set cell value
  * Get cell value

# Run tests
From the root koala_xlcalculator directory
```python
python -m unittest discover -p "*_test.py"
```

# Run Example
From the examples/common_use_case directory
```python
python use_case_01.py
```

# How to add Excel functions
Excel function support can be easily added to koala_xlcalculator.

Do the git things.. fork, clone, branch. checkout the new branch and then;
- Write a class for the function in function_library. Use existing supported function classes as template examples.
- Add the function name and related class to excel_lib.py SUPPORTED_FUNCTIONS dict
- Write a test for it in tests\\function_library. Use existing tests as template examples. Ensure you include an associated .xlsx file and implement one or more evaluate test methods. Often a great place for example test ideas is found on the Microsoft Office Excel help page for that function.
- Update the README.md to state that function is supported.
- Put your code, tests and doco forward as a pull request.

# TODO
- [] Fix all functions in the function_library so that they work.
- [] Set up a travis continuous integration service
- [] Improve testing
- [] BUGS:
  - Somewhere between the archive and the tokens we can loose parameters from a function (probs in managing the stack / RPN stuff).
    - =CONCAT("SPAM", " ", A1:B2, "SPAM", " ") gets interpreted as =CONCAT(A1:B2, "SPAM", " ")
  - If there's a gap in cells in a formula, the gap cells error (maybe they don't get read into the model?)
    - Where A1, B1, E1, and F1 have values, =COUNTA(A1:F1) errors on evalling C1
  - Using integers in the function DATE, the integers don't get parsed correctly. Potentially another example of the first bug in this list.
    - =DATE(1, 1, 1) gets interpreted as =DATE(1)
  - If you delete the sheets which are associated with a defined name, file reading breaks.
  - Ranges aren't being tokenized or eval properly. Example found in the function CHOOSE
  - Reading some dates causes a tokenizing problem. eg; =DATE(2024,1,1)
  - ExcelError evaluating Evaluator.apply("divide",4,5,None)
  - function POWER evaluates incorrectly. 2401077.2220695755 != 2401077.2220695773
  - Problem evalling: #VALUE! Evaluator.apply_one("minus", 1.475, None, None)
  - #NUM! raises an ExcelError which cascades. A #NUM! error is a legitimate value for a cell.
