Sub FormatSheet'(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	' arr = split(Parameters,"#")
	' WorkbookPath = arr(0)
	
	WorkbookPath = "C:\CuentasPorCobrarVnzla\Output\BCO ABRIL BS 2023.xlsx"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
	'Obtine el numero de hojas en el libro
	nWorkSheets = objWorkbook.Worksheets.Count
	
	'Itera sobre cada hoja del libro
	For i = 1 To nWorkSheets
	
		'Crea objeto de la hoja en curso
		Set objWorksheet = objWorkbook.Worksheets(i)
		'Valida si tiene filtros, si es asi los quita
		If objWorksheet.AutoFilterMode Then
			objWorksheet.AutoFilterMode = False
		End If
		'Genera autofit de las columnas
		objWorksheet.Cells.EntireColumn.AutoFit
		'Muestra columnas ocultas
		objWorksheet.Cells.EntireColumn.Hidden = False
		
	Next
	
	'Guarda y cierra libro
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub

Call FormatSheet