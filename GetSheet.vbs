Function GetSheet(Parameters) 
  
  arr=split(Parameters,"#")
  sBCO=arr(0)
  sTab=arr(1)
  
 
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

    Set objWorkbook = objExcel.Workbooks.Open(sBCO)
	bSheet = False
 
	nWorkSheets = objWorkbook.Worksheets.Count
	'Itera sobre cada hoja del libro
	For i = 1 To nWorkSheets
		If InStr(objWorkbook.Worksheets(i).Name, Trim(sTab)) <> 0 Then
			sTab = i
			bSheet = True
			Exit For
		End If				
	Next
	
	If bSheet = True Then
		GetSheet = sTab
	Else
		GetSheet = "error"
	End If
	' Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	' Quita la instancia del objeto Excel
	objExcel.Quit
	
End Function

MsgBox GetSheet("C:\CuentasPorCobrarVnzla\Output\BCO MAYO BS 2024.xlsx#VZLANO")