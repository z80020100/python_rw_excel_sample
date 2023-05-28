#!/usr/bin/env python3

from workbook import Workbook


TARGET_EXCEL_FILE = "sample.xlsm"
TARGET_EXCEL_SHEET = "Sheet1"
TARGET_READ_EXCEL_CELL = "AZ3"
TARGET_READ_EXCEL_RANGE = "A1:A3"
TARGET_WRITE_EXCEL_CELL = "A4"
TARGET_WRITE_EXCEL_RANGE = "A5:A7"


def main():
    # Sample for reading
    wb = Workbook(TARGET_EXCEL_FILE)
    sheets = wb.get_sheets()
    # Get sheets
    print(sheets)
    # Read cell
    AZ3 = wb.get_cell(TARGET_EXCEL_SHEET, TARGET_READ_EXCEL_CELL)
    print(AZ3)
    # Read range
    range = wb.get_range(TARGET_EXCEL_SHEET, TARGET_READ_EXCEL_RANGE)
    # Convert range to list
    column_A = range.T[0]
    print(column_A)

    # Sample for writing
    '''
    If you don't set save_formula_to_value to True,
    you can't read cell value including formula again
    (but Excel can) until you save the workbook via Excel.
    '''
    wb = Workbook(TARGET_EXCEL_FILE, writable=True, save_formula_to_value=True)
    # Write cell
    wb.set_cell(TARGET_EXCEL_SHEET, TARGET_WRITE_EXCEL_CELL, 100)
    # Write range
    wb.set_range(TARGET_EXCEL_SHEET, TARGET_WRITE_EXCEL_RANGE, 1000)
    # Save workbook
    wb.save()


if __name__ == "__main__":
    main()
