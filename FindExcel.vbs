Option Explicit
REM We use "Option Explicit" to help us check for coding mistakes

REM the Excel Application
Dim objExcel
REM the path to the excel file
Dim excelPath
REM how many worksheets are in the current excel file
Dim worksheetCount
Dim counter
REM the worksheet we are currently getting data from
Dim currentWorkSheet
REM the number of columns in the current worksheet that have data in them
Dim usedColumnsCount
REM the number of rows in the current worksheet that have data in them
Dim usedRowsCount
Dim row
Dim column
REM the topmost row in the current worksheet that has data in it
Dim top
REM the leftmost row in the current worksheet that has data in it
Dim left
Dim Cells
REM the current row and column of the current worksheet we are reading
Dim curCol
Dim curRow
REM the value of the current row and column of the current worksheet we are reading
Dim word
REM result contains the input data
Dim Message, result
Dim Title, Text1, Text2

Dim FoundCell

' Define dialog box variables.
Message = "Please enter IP Address:"           
Title = "Read Data"
Text1 = "User input canceled"
Text2 = "You entered:" & vbCrLf

' Ready to use the InputBox function
' InputBox(prompt, title, default, xpos, ypos)
' prompt:    The text shown in the dialog box
' title:     The title of the dialog box
' default:   Default value shown in the text box
' xpos/ypos: Upper left position of the dialog box 
' If a parameter is omitted, VBScript uses a default value.

result = InputBox(Message, Title, "", 100, 100)

' Evaluate the user input.
If result = "" Then    ' Canceled by the user
    WScript.Echo Text1
Else

	REM where is the Excel file located?
	excelPath = "D:\ExcelTest\addressDetails.xlsx"

	WScript.Echo "Reading Data from " & excelPath

	REM Create an invisible version of Excel
	Set objExcel = CreateObject("Excel.Application")

	REM don't display any messages about documents needing to be converted
	REM from  old Excel file formats
	objExcel.DisplayAlerts = 0

	REM open the excel document as read-only
	REM open (path, confirmconversions, readonly)
	objExcel.Workbooks.open excelPath, false, true

	REM How many worksheets are in this Excel documents
	workSheetCount = objExcel.Worksheets.Count

	WScript.Echo "We have " & workSheetCount & " worksheets"

	REM Loop through each worksheet
	For counter = 1 to workSheetCount
		WScript.Echo "-----------------------------------------------"
		WScript.Echo "Reading data from worksheet " & counter & vbCRLF

		Set currentWorkSheet = objExcel.ActiveWorkbook.Worksheets(counter)
		REM how many columns are used in the current worksheet
		REM usedColumnsCount = currentWorkSheet.UsedRange.Columns.Count
		REM how many rows are used in the current worksheet
		REM usedRowsCount = currentWorkSheet.UsedRange.Rows.Count

		REM What is the topmost row in the spreadsheet that has data in it
		REM top = currentWorksheet.UsedRange.Row
		REM What is the leftmost column in the spreadsheet that has data in it , lookat:=xlWhole
		REM left = currentWorksheet.UsedRange.Column
		Set Cells = currentWorksheet.Cells
		Set FoundCell = currentWorkSheet.Range("B:B").Find(result)
		If Not FoundCell Is Nothing Then
			MsgBox (result & " found in row: " & FoundCell.Row)
			REM get the value that is in the next cell in same row
			word = Cells(FoundCell.Row, FoundCell.Column + 1).Value
			MsgBox (result & " IP Address value is: " & word )
		Else
			MsgBox (result & " not found")
		End If
	
		REM We are done with the current worksheet, release the memory
		Set currentWorkSheet = Nothing
	Next
	
	objExcel.Workbooks(1).Close
	objExcel.Quit

	Set currentWorkSheet = Nothing
	REM We are done with the Excel object, release it from memory
	Set objExcel = Nothing
	
End If
