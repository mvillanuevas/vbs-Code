'C:\CuentasPorCobrarVnzla\Output/BCO MAYO BS 2023.xlsx#2400#156167354#CENTRO DISTRIBUCIONES FRANCIS#04/02/2024#04/02/2024#BOT#BANESCO#500000002

Sub macro'(SetValues) 
  
  Dim arr,sIGTFilePath,sExcelRow,sCodClient,sClient,sEmailSendDate,sCurrentDate,sValueBot,sTab,sDocumentNumber
  
  SetValues="C:\CuentasPorCobrarVnzla\Output/BCO MAYO BS 2023.xlsx#2400#156167354#CENTRO DISTRIBUCIONES#04/02/2024#04/02/2024#BOT#BANCO#500000002"
  
  arr=split(SetValues,"#")
  sIGTFilePath=arr(0)
  sExcelRow=arr(1)
  sCodClient=arr(2)
  sClient=arr(3)
  sEmailSendDate=arr(4)
  sCurrentDate=arr(5)
  sValueBot=arr(6)
  sTab=arr(7)
  sDocumentNumber=arr(8)
  
 
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkbook = objExcel.Workbooks.Open(sIGTFilePath)

    objExcel.Application.Visible = True
	
	nWorkSheets = objWorkbook.Worksheets.Count
	'Itera sobre cada hoja del libro
	For i = 1 To nWorkSheets
		If InStr(objWorkbook.Worksheets(i).Name, Trim(sTab)) <> 0 Then
			sTab = objWorkbook.Worksheets(i).Name
			Exit For
		End If				
	Next

	objWorkbook.Sheets(sTab).Activate  

  
	objExcel.Range("F" & sExcelRow) = sCodClient
  
	objExcel.Range("G" & sExcelRow) = sClient      

	objExcel.Range("H" & sExcelRow) = sEmailSendDate

	objExcel.Range("I" & sExcelRow) = sCurrentDate
  
	objExcel.Range("J" & sExcelRow) = sDocumentNumber

	objExcel.Range("K" & sExcelRow) = sValueBot
  
                      
      'Saves and closes the Excel file'
   objWorkbook.Save
   objWorkbook.Close SaveChanges = True

End Sub
Call macro