'Microsoft Excel Automation Basics
':: Create and edit an Excel File.
'---------------------------------

'create the excel object
	Set Excel = CreateObject("Excel.Application") 

'view the excel program and file, set to false to hide the whole process
	'Excel.Visible = True 

'add a new workbook
	Set objWorkbook = Excel.Workbooks.Add 

	Excel.range("E2") = "water"

'set a cell value at row 1 column 1
	objExcel.Cells(1,1).Value = "Time in Min"

'change a cell value
	objExcel.Cells(2,1).Value = "2"

'save the new excel file (make sure to change the location) 'xls for 2003 or earlier
	objWorkbook.SaveAs "VBScript\Report\Report-" &Day(Now)& "-" &Month(Now)& "-" &Year(Now)& ".xlsx" 

'exit the excel program
	Excel.Quit

'release objects
	Set Excel = Nothing
	Set objWorkbook = Nothing
