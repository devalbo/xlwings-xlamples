# https://docs.xlwings.org/en/stable/api.html

# First, get it working. Next, get it working right. If it matters, get it working fast.

import xlwings as xw


# open existing workbook
def open_workbook(name):
    ex = xw.Book(name)
    return ex


# list sheets in workbook
def list_sheets(wb):
    for s in wb.sheets:
        print(s.name)


# add a sheet to a workbook
def add_a_sheet(wb, sheet_name):
    new_sheet = wb.sheets.add(sheet_name)
    return new_sheet


# add data to a worksheet
def add_data_to_sheet(sheet, data, start="A1"):
    sheet.range(start).value = data


# get value for a cell in a worksheet
def get_cell_value_from_sheet(sheet, cell):
    return sheet.range(cell).value


# get table's worth of data from a starting point in a worksheet
def get_expanded_value_from_sheet(sheet, start_cell="A1"):
    return sheet.range(start_cell).expand().value


# set a specific cell's color
def set_cell_color(sheet, cell, color):
    sheet.range(cell).color = color


# clear a specific cell's color
def clear_cell_color(sheet, cell):
    sheet.range(cell).color = None


def set_formula(sheet):
    sheet['B1'].value = '=A1*3'


def apply_formula(sheet, range="B1"):
    sheet[range].formula = '=A1*3'




if __name__ == "__main__":
    MY_SHEET_NAME = "PythonGenerated"

    # _wb = open_workbook("example.xlsx")
    # list_sheets(_wb)
    #
    # # _wb.close()
    # #
    # _sheet = add_a_sheet(_wb, MY_SHEET_NAME)
    # #
    # MY_DATA = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    # add_data_to_sheet(_sheet, MY_DATA)
    # #
    # _the_data = get_cell_value_from_sheet(_sheet, "A1")
    # print(_the_data)
    # #
    # _table_data = get_expanded_value_from_sheet(_sheet)
    # print(_table_data)
    #
    # set_cell_color(_sheet, "C3", (255, 0, 0))
    #
    # # clear_cell_color(_sheet, "C3")
    # #
    # import time
    # time.sleep(6.5)
    # clear_cell_color(_sheet, "C3")
    # # #
    # # # _wb.save()
    # #
    # _wb.save("done.xlsx")

    # Interactive
    # import xlwings
    # b = xlwings.Book()
    # Column A: down column, increase values by 1
    # import demo
    # demo.apply_formula(b.sheets[0], "B1:B10")
    # importlib.reload(demo)
