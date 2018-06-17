Excel Workbook Sheet Combiner

Requirments:
  - Python 3.5 or greater
  - pypiwin32
      - This can be installed using: pip install pypiwin32

To run:
  - Download excel_sheet_combine.py to computer
  - Move to folder with Excel workbooks you would like to combine
  - Run excel_sheet_combine.py from cmd line in that directory, being sure to give it the names of the workbooks you want the script to       combine. 
  
 I think I commented this enough that it can be modified as needed since the documentation for win32com in python is sparse.
  
 Please see below for when to use/when not to use:
  
This script will take a list of excel workbook file names and move every 
sheet in them into a new excel file called 'your_combined_files.xlsx'. Sheets in different workbooks will not be copied over, though this can be changed pretty easily in the script. 

This method of moving sheets from mutiple workbooks into one wookbook is only 
more useful than Pandas if you require formatting to be kept across woorkbooks.

This program will likely not working on any OS other than Windows.
