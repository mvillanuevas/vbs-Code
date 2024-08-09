Sub BancosBCO(Parameters)
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)
	wSheet = arr(1)
	amount = arr(2)
	fecha = arr(3)
	referencia = arr(4)
	codcliente = arr(5)
	cliente = arr(6)
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
	Set objWorksheet = objWorkbook.Worksheets(CInt(wSheet))
	
	Const xlValues = -4163
  	Const xlPart = 2
	
	lastRow = objWorksheet.Cells(objWorksheet.Rows.Count,1).End(-4162).Row + 1
	
  	Set mRange = objWorksheet.Range("E:E")
	
	Dim mFind :	Set mFind = mRange.Find(amount,,xlValues,xlPart)
	MsgBox mFind.Row
	If mFind Is Nothing Then
		objWorksheet.Cells(lastRow, 1).value = fecha
		objWorksheet.Cells(lastRow, 2).value = referencia
		objWorksheet.Cells(lastRow, 5).value = amount
		objWorksheet.Cells(lastRow, 9).value = codcliente
		objWorksheet.Cells(lastRow, 10).value = cliente
	End If
	
	'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	'Quita la instancia del objeto Excel
	objExcel.Quit
End Sub

Parameters = "C:\CuentasPorCobrarVnzla\Output/BCO FEBRERO BS 2024.xlsx#3#100,000.00#2/20/2024#25563858988#100124137#"
Call BancosBCO(Parameters)