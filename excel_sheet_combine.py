# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 17:49:52 2018

This script will take a list of excel workbook file names and move every 
sheet in them into a new excel file called 'your_combined_files.xlsx'. 
If more than one sheet shares the same name across workkbooks the script 
will append (n) to the last one inserted.
EXAMPLE: foo will become foo(2), the next will become foo(3)

This method of moving sheets from mutiple workbooks into one wookbook is only 
more useful than Pandas if you require formatting to be kept across woorkbooks.
@author: Ian
"""

import sys
import win32com.client as win32
import os

list_of_files = []

def main(list_of_files):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    """
    This sets excel to appear on the screen as it reallocates wooksheets.
    This code helps to clear out any errors that might be hard for a user 
    to see if excel processed in the background.
    """
    excel.Visible = 1
    """ Creates a new workbook """
    wb = excel.Workbooks.Add()
        
    for file in list_of_files:
        """ Sets entire file name to be current working directory """
        file = os.path.join(os.getcwd(), file)
        """ opens each file """
        w = excel.Workbooks.Open(file)
        """ 
        Uses Sheets.Count to establish a total number of sheets in 
        a workbook. The range of (1, count + 1) is neccessary to avoid 
        an error from Sheet[0]
        """
        for i in range(1, w.Sheets.Count + 1):
            """ copies each sheet to the location in the new excel file """
            w.Sheets(i).Copy(wb.Sheets(i))
    wb.SaveAs(os.path.join(os.getcwd(), 'your_combined_files.xlsx'))
    excel.Application.Quit()

if __name__ == "__main__":  
    num_files = len(sys.argv) - 1
    for i in range(0, num_files):
        list_of_files.insert(i, sys.argv[i + 1]) 
    main(list_of_files)