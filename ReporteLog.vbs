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
'Parameters = "C:\CuentasPorCobrarVnzla\Output/Reporte_Log.xlsx#2024-07-26#Fw: SOPORTE DE PAGO#26/07/2024 08:35:31#Miguel.Fandino@mdlz.com#0#Error on task: ST_DollarsTransaction; error: SP_SAPTransaction_F-28_EndProcess error:Incorrect field path entered. Verify correct field path and try again.; line:74 line: 148#MAYO"
Parameters = "C:\CuentasPorCobrarVnzla\Output/Reporte_Log.xlsx#2024-07-29#SOPORTE DE PAGO_BS_DISTRIBUIDORA FULL PROGRESO,_005-367243#29/07/2024 18:55:09#Manuel.Serrano@mdlz.com#1##JULIO"

Call ReporteLog(Parameters)
