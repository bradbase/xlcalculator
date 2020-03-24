
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
    * SUM
    * AVERAGE
  * Set cell value

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

# TODO
- [] Fix all functions in the function_library so that they work. Currently only SUM and AVERAGE are being maintained.
- [] Set up a travis continuous integration service
- [] Improve testing
