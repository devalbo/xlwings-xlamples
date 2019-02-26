import xlwings as xw
import pandas as pd
import numpy as np
from scipy import stats


def read_data_into_pandas():
    sheet = xw.Book.caller().sheets['Main']

    raw_data = sheet.range('c4').expand().value
    data_dataframe = pd.DataFrame(data=raw_data[1:], columns=raw_data[0])

    differenced_doodads = np.diff(data_dataframe[raw_data[0][1]])
    differenced_thingamabobs = np.diff(data_dataframe[raw_data[0][2]])

    try:
        xw.sheets.add('Regression', after=xw.sheets[-1])
    except ValueError:
        xw.sheets['Regression'].clear_contents()

    xw.sheets['Regression'].range('B2').value = 'Slope'
    xw.sheets['Regression'].range('B3').value = 'Intercept'
    xw.sheets['Regression'].range('B4').value = 'Correlation Coefficient'
    xw.sheets['Regression'].range('B5').value = 'P Value'
    xw.sheets['Regression'].range('B6').value = 'Std Error'

    # The list comprehension here is used to transpose the results.
    xw.sheets['Regression'].range('C2').value = [
        [x] for x in stats.linregress(differenced_doodads, differenced_thingamabobs)]


if __name__ == "__main__":
    read_data_into_pandas()
