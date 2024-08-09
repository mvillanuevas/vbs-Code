Sub ReporteLog(Parameters)
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)
	fecha = arr(1)
	asunto = arr(2)
	cfecha = arr(3)
	remitente = arr(4)
	estatus = arr(5)
	comentarios = arr(6)
	mes = arr(7)
	
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

	nWorkSheets = objWorkbook.Worksheets.Count
	rSheet = "Reporte_" & mes & "_" & Year(Now)
	bSheet = false
	'Itera sobre cada hoja del libro
	For i = 1 To nWorkSheets
		If InStr(objWorkbook.Worksheets(i).Name, Trim(rSheet)) <> 0 Then
			bSheet = true
			Exit For
		End If				
	Next

	If bSheet = false Then
		Set objWorkSheet = objWorkbook.Sheets.Add
		objWorkSheet.Name = rSheet

		With objWorkSheet
			.Cells(1,1) = "Test1"
			.Cells(1,1).Font.Blod = True
			.Range("A1:F1") = Array("Fecha","CorreoAsunto","CorreoFecha","Remitente","Estatus","Comentarios")
			.Range("A1:F1").Font.Bold = True
			.Range("A1:F1").Interior.Color = RGB(153, 102, 255)
		End With
	End If
	
	Set objWorksheet = objWorkbook.Worksheets(rSheet)
	
	lastRow = objWorksheet.Cells(objWorksheet.Rows.Count,1).End(-4162).Row + 1
	
	If estatus = "1" Then
		estatus = "Procesado"
	Else
		estatus = "Pendiente"
	End If
	
	objWorksheet.Cells(lastRow,1).value = fecha
	objWorksheet.Cells(lastRow,2).value = asunto
	objWorksheet.Cells(lastRow,3).value = cfecha
	objWorksheet.Cells(lastRow,4).value = remitente
	objWorksheet.Cells(lastRow,5).value = estatus
	objWorksheet.Cells(lastRow,6).value = comentarios
	
		'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	'Quita la instancia del objeto Excel
	objExcel.Quit
End Sub

