# xlwings

## Python Principles
- Style Guide (PEP-8): https://www.python.org/dev/peps/pep-0008/#introduction
- Zen of Python (PEP-20): https://www.python.org/dev/peps/pep-0020/
- `import antigravity`

## Introduction
- Excel is great for visualization, but large datasets are cumbersome
- xlwings is a free and open-source add-in for Python
- xlwings provides access to Python’s libraries
- Different from a dll
- Docs are here: https://docs.xlwings.org/en/stable/api.html

What you will need for this session:
- Pycharm
- https://github.com/devalbo/xlwings-xlamples

## Installation
At the command line:
- `pip install xlwings --upgrade`
- `xlwings addin install`

In Excel:
- Alt+F11 to bring up VBA Editor
- **Tools->References** - Make sure xlwings is ticked

## The Basics
- Using Python to control Excel - [demo.py](demo.py)

## Performance matters?
[loop.py](loop.py)

## Invert control
![This is how you invert](https://i.imgur.com/cnkUFvH.jpg)
##### From Excel, use Python as a VBA substitute or replacement
  - Ensure Developer tab is enabled in Excel: https://docs.xlwings.org/en/stable/addin.html#installation
  - Execute Python functionality from VBA code
  - Expose Python methods to Excel as UDF
  - [myxl.py](myxl.py) and [myxl.xlsm](myxl.xlsm)
##### Integrating Data with Other Applications
  - Pandas - https://pandas.pydata.org/
  - See [data_exports.py](data_exports.py) - call macro from [data_exports.xlsm](data_exports.xlsm)
##### Back and forth with Python calling macros 
  - `xlwings quickstart macro_example`
  - [macro_example.py](macro_example.py)
  