Sub TasaPromedio'(Parameters)
	Parameters = "C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de MONTI SIN LIMITES $.xlsx"
	
	arr=split(Parameters,"#")
	sReportePath=arr(0)
	'Opens the Excel file'
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

	Set objWorkbook = objExcel.Workbooks.Open(sReportePath)
	Set objWorksheetR = objWorkbook.Worksheets("Resumen de Cobranza")
	
	MsgBox objWorksheetR.Range("C2").value
	
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	objExcel.Quit

End Sub

Call TasaPromedio