# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 17:49:52 2018

This script will take a list of excel workbook file names and move every 
sheet in them into a new excel file called 'your_combined_files.xlsx'. 

This method of moving sheets from mutiple workbooks into one workbook is only
more useful than Pandas if you require formatting to be kept across workbooks.
@author: Ian
"""

import sys
import win32com.client as win32
import os

list_of_files = []


def main(list_of_files):
    """
    TODO: Figure out to get this to check if user has Excel open and not proceed if so
    if(win32.GetActiveObject("Excel.Application") == True):
        print('You must close all open Excel Applications first')
    else:
    """
    list_of_sheets_in_file = []
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    """
    This sets excel to appear on the screen as it reallocates worksheets.
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
        a workbook. The range of (1, count + 1) is necessary to avoid 
        an error from Sheet[0]
        """
        for index in range(1, w.Sheets.Count + 1):
            """ copies each sheet to the location in the new excel file """
            if w.Sheets(i).Name not in list_of_sheets_in_the_end:
                w.Sheets(i).Copy(wb.Sheets(i))
                list_of_sheets_in_file.append(w.Sheets(i).Name)

    """ 
    These three lines will remove the default (and blank) Sheet1 that is
    automatically now when the new Excel file is created. Normally the user is asked 
    before sheets are deleted. .DisplayAlerts removes this. 
    """
    excel.DisplayAlerts = False
    wb.Sheets["Sheet1"].Delete()
    excel.DisplayAlerts = True
    """ 
    Saves and then closes all Excel applications that are open. This will currently include all other 
    Excel workbooks that a user might have open. 
    """
    wb.SaveAs(os.path.join(os.getcwd(), 'your_combined_files.xlsx'))
    excel.Application.Quit()
    del excel


if __name__ == "__main__":
    num_files = len(sys.argv) - 1
    for i in range(0, num_files):
        list_of_files.insert(i, sys.argv[i + 1])
    main(list_of_files)
