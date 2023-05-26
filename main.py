#!/usr/bin/env python3

from workbook import Workbook


TARGET_EXCEL_FILE = "sample.xlsm"
TARGET_EXCEL_SHEET = "Sheet1"
TARGET_EXCEL_CELL = "AZ3"
TARGET_EXCEL_RANGE = "A1:A3"


def main():
    wb = Workbook(TARGET_EXCEL_FILE)
    sheets = wb.get_sheets()
    # Get sheets
    print(sheets)
    # Read cell
    AZ3 = wb.get_cell(TARGET_EXCEL_SHEET, TARGET_EXCEL_CELL)
    print(AZ3)
    # Read range
    range = wb.get_range(TARGET_EXCEL_SHEET, TARGET_EXCEL_RANGE)
    # Convert range to list
    column_A = range.T[0]
    print(column_A)


if __name__ == "__main__":
    main()
