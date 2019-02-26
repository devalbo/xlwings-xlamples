from string import ascii_uppercase
import demo
import xlwings
import datetime

print(ascii_uppercase)

RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLUE = (0, 0, 255)
BLACK = (0, 0, 0)
YELLOW = (255, 255, 0)

colors = [RED, GREEN, BLUE, BLACK, YELLOW]

# "python iterate over alphabet" -> https://stackoverflow.com/questions/17182656/how-do-i-iterate-through-the-alphabet-in-python-please

GRID_HEIGHT = 100

wb = xlwings.Book()
sheet = wb.sheets[0]


def color_by_cells():
    start_time = datetime.datetime.now()

    i = 0
    for color in colors:
        col = ascii_uppercase[i]
        print(col)

        for y in range(1, GRID_HEIGHT + 1):
            cell = col + str(y)
            # print(cell)
            demo.set_cell_color(sheet, cell, color)
        i += 1

    end_time = datetime.datetime.now()
    duration = end_time - start_time
    print(f"DURATION: {duration}")


def color_by_column():
    start_time = datetime.datetime.now()

    # see http://book.pythontips.com/en/latest/enumerate.html
    for index, color in enumerate(colors):
        col = ascii_uppercase[index]

        column_range = f"{col}1:{col}{GRID_HEIGHT + 1}"
        print(column_range)

        sheet.range(column_range).color = color

    end_time = datetime.datetime.now()
    duration = end_time - start_time
    print(f"DURATION: {duration}")


if __name__ == "__main__":
    # color_by_cells()
    # color_by_column()

    pass
