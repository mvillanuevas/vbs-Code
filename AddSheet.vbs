Sub AddSheet(Parameters)
  
	Dim arr,strZEP1CompletePath
	arr=split(Parameters,"#")
	strZEP1CompletePath=arr(0)
	strSheet=arr(1)

  
	'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")
  
	objExcel.Application.Visible = False
	objExcel.Application.DisplayAlerts = False
	objExcel.Application.ScreenUpdating = False

	Set objWorkbook = objExcel.Workbooks.Open(strZEP1CompletePath)
	Set objWorksheet = objWorkbook.Sheets.Add
	objWorksheet.Name = strSheet
	
	' Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	' Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub

Call AddSheet("C:\DSDandCaseFillRate\Input\ZEP1.xlsx#Sheet2")