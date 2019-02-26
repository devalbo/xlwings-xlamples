import xlwings as xw
import time


def call_vba_macros():
    workbook = xw.Book.caller()
    macro = workbook.macro('ProgressBar')

    for i in range(0, 11):
        percentage_completion = i * 10
        macro(str(percentage_completion) + '%')

        time.sleep(1)

    macro_complete = workbook.macro('Complete')
    macro_complete()

    macro('Ready')


"""
xlwings quickstart macro_example

Sub CallVBAMacros()
    RunPython("import macro_example; macro_example.call_vba_macros()")
End Sub

Sub ProgressBar(msg As String)
    Application.StatusBar = msg
End Sub

Sub Complete()
    MsgBox("Congratulations! You are familiar with XLWings now!")
End Sub

"""