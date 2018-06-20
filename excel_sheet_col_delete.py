# -*- coding: utf-8 -*-
"""
Created on Mon Jun 18 16:02:35 2018
Takes and Excel workbook as input and checks every column in every worksheet in that workbook for a
preassigned list of column names. If found the program deletes that column.

TODO: fix worker thread issue if given large file
@author: Ian
"""

import sys
import os
import win32com.client as win32
import pythoncom


def main(file: string):
    list_of_bad_cols = ['CREATION_USERID', 'LAST_UPD_USERID', 'LAST_UPD_TMSTMP', 'CREATION_TMSTMP']
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = 0
    file_path = os.path.join(os.getcwd(), file)
    wb = excel.Workbooks.Open(file_path)

    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Worksheets(i)
        # This bit is magic to get total number of cols in a given worksheet
        # DO NOT TOUCH
        xl_to_left = -4159
        col_count = ws.Cells(1, ws.Columns.Count).End(xl_to_left).Column

        # Since we are deleting columns (which shifts the column index left) the index of the columns
        # is not constant. Thus we must only increase the starting index if we do not delete a column.
        # This can possibly lead to an infinite loop. PLEASE BE CAREFUL MODIFYING
        # TODO: see if there's a better way to write this code
        col_idx = 1
        while col_idx < col_count + 1:
            col_heading = ws.Cells(1, col_idx).Value
            if col_heading and any(x in col_heading.strip() for x in list_of_bad_cols):
                ws.Columns(col_idx).Delete()
                col_count -= 1
                continue
            else:
                col_idx += 1
    # TODO: see if I actually need all four of these lines to quit Excel
    wb.Close(True)
    excel.Application.Quit()
    pythoncom.CoUninitialize()
    del excel


if __name__ == "__main__":
    main(sys.argv[1])
