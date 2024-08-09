Sub macro(Parameters)
  
	Dim arr,strZEP1CompletePath,strBaseCompletePath
	arr=split(Parameters,"#")
	strZEP1CompletePath=arr(0)
	strBaseCompletePath=arr(1)

  
	'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkbook1 = objExcel.Workbooks.Open(strZEP1CompletePath)
	Set objWorksheet1 = objWorkbook1.Worksheets("Sheet2")	
	Set objWorksheet3 =objWorkbook1.Worksheets("Sheet3")

	Set objWorkbook2 = objExcel.Workbooks.Open(strBaseCompletePath)
	Set objWorksheet2 = objWorkbook2.Worksheets("limpia")
  
	objExcel.Application.Visible = True
	objExcel.Application.DisplayAlerts = False
	objExcel.Application.ScreenUpdating = False

	' objExcel.Application.Wait(Now + TimeValue("00:00:10")) 
	'objExcel.Application.Wait(Now + TimeValue("00:00:05")) 
  

	' objWorksheet2.Activate
  
    lastRow = objWorksheet2.Cells(objWorksheet2.Rows.Count,1).End(-4162).Row


	objWorksheet2.Range("A1:AD" & lastRow).Copy objWorksheet1.Range("A1")

   'Saves and closes the Excel file'
	objWorkbook2.Save
	objWorkbook2.Close SaveChanges = True

	' objWorksheet1.Activate
	objWorksheet1.Columns("J:J").Insert
	objWorksheet1.Range("J1").FormulaR1C1 = "D"

	'Autofit en la columna J
	lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,1).End(-4162).Row
	objWorksheet1.Range("J2").FormulaR1C1 = "=RC[-2]-RC[-1]"
	objWorksheet1.Range("J2").AutoFill(objWorksheet1.Range("J2:J" & lastRow))

	' objWorksheet1.Sheets("Sheet2").Select
    objWorksheet1.Sheets.Add
    ' objWorksheet1.Sheets("Sheet2").Select
	
	' objWorksheet1.Activate
    objWorksheet1.Range("A1:AE1").AutoFilter
    objWorksheet1.Range("J1").AutoFilter 10, Array("<>0")

	objWorksheet1.Activate
	objWorksheet1.Range("A1:AE" & lastRow).Copy objWorkbook1.Worksheets("Sheet3").Range("A1")
       
	' objWorksheet3.Activate
	objWorksheet3.Range("B:D").Delete
	objWorksheet3.Columns("C:C").Insert
	objWorksheet3.Range("C1").FormulaR1C1 = "CDS"



	'Saves and closes the Excel file'
	objWorkbook1.Save
	objWorkbook1.Close SaveChanges = True
  
    objExcel.Quit

End Sub