'Microsoft Excel Automation Basics
':: Open and edit an Excel File.
'---------------------------------

'create the excel object
	Set objExcel = CreateObject("Excel.Application") 

'view the excel program and file, set to false to hide the whole process
	objExcel.Visible = True 

'open an excel file (make sure to change the location) .xls for 2003 or earlier
	Set objWorkbook = objExcel.Workbooks.Open("VBScript\Report\Report-" &Day(Now)& "-" &Month(Now)& "-" &Year(Now)& ".xlsx")

'set a cell value at row 3 column 5
	objExcel.Cells(1,1).Value = "Time in Min"

'change a cell value
	objExcel.Cells(9,1).Value = "9"
	
'delete a cell value
	'objExcel.Cells(3,5).Value = ""

'get a cell value and set it to a variable
	r3c5 = objExcel.Cells(3,5).Value

'save the existing excel file. use SaveAs to save it as something else
	objWorkbook.Save

'close the workbook
	objWorkbook.Close 

'exit the excel program
	objExcel.Quit

'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing