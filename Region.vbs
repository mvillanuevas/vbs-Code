Sub Region'(Parameters)

	arr = split(Parameters,"#")
	WorkbookPathData = arr(0)
	
	WorkbookPathMaestro = "C:\CuentasPorCobrarVnzla\Output\Maestro Clientes.xlsx"
	cClient = "156162624"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbookMaestro = objExcel.Workbooks.Open(WorkbookPathMaestro,,True)
	Set objWorksheetMaestro = objWorkbookMaestro.Worksheets("MAESTRO")
	
	Const xlPart = 2
	Const xlValues = -4163
	
	Set mRangeMaestro = objWorksheetMaestro.Range("A:A")
	
	Dim mFind :	Set mFind = mRangeMaestro.Find(cClient,,xlValues,xlPart)
	
	If Not mFind Is Nothing Then
		sRegion = objWorksheetMaestro.Cells(mFind.Row, mFind.Column + 6)
	End If
	
	MsgBox Day(Now) & "/" & Month(Now) & "/" & Year(Now)
	
	
End  Sub

Call Region