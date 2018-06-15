# excel_sheet_combiner

To run:
  - Download excel_sheet_combine to computer
  - Move to folder with Excel workbooks you would like to combine
  - Run from cmd line in that directory, being sure to give it the names of the workbooks you want the script to combine. 
  
 I think I commented this enough that it can be modified as needed since the documentation doesn't seem to exist for win32com in python.
  
 Please see below for when to use/when not to use:
  
This script will take a list of excel workbook file names and move every 
sheet in them into a new excel file called 'your_combined_files.xlsx'. 
If more than one sheet shares the same name across workkbooks the script 
will append (n) to the last one inserted.
EXAMPLE: foo will become foo(2), the next will become foo(3)

This method of moving sheets from mutiple workbooks into one wookbook is only 
more useful than Pandas if you require formatting to be kept across woorkbooks.
