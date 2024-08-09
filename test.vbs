Sub CopyCreateWorkbook(ParametersTwo)

  Dim arr,strExcelName,strNewWb,strSheetName
  arr=split(ParametersTwo,"#")
  strExcelName=arr(0)
  strSheetName=arr(1)
  strNewWb=arr(2)

  
 
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

  Set objWorkbook = objExcel.Workbooks.Open(strExcelName)
  Set objWorkbook2 = objExcel.Workbooks.Add()

  
  'Activate Worksheet from workbook 1
  ' objWorkbook.Sheets(strSheetName).Activate
  Set objWorksheet = objWorkbook.Worksheets(strSheetName)
  ' llastrow = ObjExcel.Activesheet.UsedRange.Rows.Count
  llastrow = objWorksheet.Cells(objWorksheet.Rows.Count,6).End(-4162).Row

  'Copy data
  objWorksheet.Range("A12:L" & llastrow).Copy
  'Activate workbook 2
  ' objWorkbook2.Sheets(1).Activate 
  Set objWorksheet2 = objWorkbook2.Worksheets(1)
  'Paste data
  objWorksheet2.Range("A1").PasteSpecial 12
  objWorksheet2.Range("A1").PasteSpecial 8

   ' objWorkbook.Sheets(1).Activate
  ' objExcel.Application.CutCopyMode = False
    ' objWorkbook2.Sheets(1).Activate
    objExcel.Application.CutCopyMode = False
 
                      
            
    'Saves and closes the Excel file'
   objWorkbook.Save
   objWorkbook.Close SaveChanges = True
      
    'Saves and closes the Excel file'
    objWorkbook2.SaveAs(strNewWb)
    'objWorkbook2.Save
    objWorkbook2.Close SaveChanges = True
    
End Sub
call CopyCreateWorkbook