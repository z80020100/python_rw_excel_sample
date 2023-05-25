
import string

import pandas as pd

CELL_NAME_ERROR = "cell name is not valid."
ROW_NAME_ERROR = "row name is not valid."
RANGE_NAME_ERROR = "range name is not valid."


class Workbook:

    def __init__(self, filename):
        self.filename = filename
        self.workbook = pd.read_excel(filename, header=None, sheet_name=None)

    def get_sheets(self):
        return self.workbook.keys()

    def get_cell(self, sheet_name, cell_name):
        row_index, column_index = Workbook.convert_cell_name_to_index(
            cell_name)
        sheet = self.workbook[sheet_name]
        return sheet.iloc[row_index, column_index]

    # Use Excel range format
    def get_range(self, sheet_name, range_name):
        cutting_index = 0
        for i, c in enumerate(range_name):
            if c == ":":
                cutting_index = i
                break
        if cutting_index == 0:
            raise ValueError(RANGE_NAME_ERROR)
        # Split range name to start cell and end cell
        start_cell_name = range_name[:cutting_index]
        end_cell_name = range_name[cutting_index + 1:]
        start_row_index, start_column_index = Workbook.convert_cell_name_to_index(
            start_cell_name)
        end_row_index, end_column_index = Workbook.convert_cell_name_to_index(
            end_cell_name)
        sheet = self.workbook[sheet_name]
        return sheet.iloc[start_row_index:end_row_index + 1,
                          start_column_index:end_column_index + 1]

    @staticmethod
    def convert_cell_name_to_index(cell_name):
        upper_cell_name = cell_name.upper()
        cutting_index = 0
        for i, c in enumerate(upper_cell_name):
            if c not in string.ascii_letters:
                cutting_index = i
                break
        if cutting_index == 0:
            raise ValueError(CELL_NAME_ERROR)
        # Split cell name to column and row
        column = upper_cell_name[:cutting_index]
        row = upper_cell_name[cutting_index:]
        # Convert column name to index
        column_index = 0
        for i in range(len(column)):
            column_index += ((ord(column[i]) - ord("A")) + 26 * i)
        # Convert row name to index
        row_index = 0
        try:
            row_index = int(row) - 1
        except ValueError:
            raise ValueError(ROW_NAME_ERROR)
        if row_index < 0:
            raise ValueError(ROW_NAME_ERROR)
        return row_index, column_index
