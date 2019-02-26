# import xlwings as xw
#
#
# def hello_xlwings():
#     wb = xw.Book.caller()
#     wb.sheets[0].range("A1").value = "Hello xlwings!"
#
#
# @xw.func
# def hello(name):
#     return "hello {0}".format(name)

import xlwings as xw
import time


def call_vba_macros():
    workbook = xw.Book.caller()
    macro = workbook.macro('ProgressBar')

    for i in range(0, 10):
        percentage_completion = i * 10
        macro(str(percentage_completion) + '%')

        time.sleep(1)

    macro_complete = workbook.macro('Complete')
    macro_complete()

    macro('Ready')
