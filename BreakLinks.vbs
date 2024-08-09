Sub BreakLinks
	WorkbookPath = "C:\Users\CSF5266\Downloads\BS - OC. DISTRIBUCIONES.xlsx"
	
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel sin actualizar links
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath,0)
	
	'Elimina links del libro
	If Not IsEmpty(objWorkbook.LinkSources(1)) Then
		For Each ext_Link In objWorkbook.LinkSources(1)
			objWorkbook.BreakLink ext_Link,1
		Next
	End If
	
	'Guarda y cierra libro
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub

Call BreakLinks