Sub Quota(Parameters)
	
	Parameters = "C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\Control Cuota MT.xlsx#C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\PowerBI.csv#C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\CDR - Customer data report.xlsx#4/19/2024"
	
	arr=split(Parameters,"#")
	sTemplateFile=arr(0)
	sPowerBIFile=arr(1)
	sCDRFile=arr(2)
	sDiaControl=arr(3)
	
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = True
	MsgBox "00"
	' Set objWorkbookBI = objExcel.Workbooks.Open(sPowerBIFile)
	' Set objWorksheetBI = objWorkbookBI.Worksheets("PowerBI")
	
	Set objWorkbookTemplate = objExcel.Workbooks.Open(sTemplateFile)
	Set objWorksheetS = objWorkbookTemplate.Worksheets("Status")
	Set objWorksheetR = objWorkbookTemplate.Worksheets("Redistribucion")
	Set objWorksheetC = objWorkbookTemplate.Worksheets("CDR (sharepoint)")
	Set objWorksheetD = objWorkbookTemplate.Worksheets("DinamicaStatus")
	
	' ' **** Llenar hoja Status ****
	' llastrow = objWorksheetBI.Cells(objWorksheetBI.Rows.Count,2).End(-4162).Row
	
	' Const xlValues = -4163
  	' Const xlPart = 2
	
	' Set mRangeTemp = objWorksheetBI.Rows(1)
	' Dim tColumn
	' tColumn = 98
	' ' Listado de columnas BI
	' lBIColumns = Array("custom.status0","cust_name1","g_level2","a_level4","loc_code","po_date","po_num","ship_to","order_num","order_item","sku_name","sku_code","qty")
		
	' For each lColumn in lBIColumns
		' ' Busca las columna en el reporte de BI
		' Dim mFindC : Set mFindC = mRangeTemp.Find(lColumn,,xlValues,xlPart)
		' If Not mFindC Is Nothing Then
			' ' Si encuentra la columna, copia y pega apartir de la columna B en el Template
			' nRow = mFindC.Row
			' nColumn = Split(mFindC.Address,"$")(1)
			' sColumn = Chr(tColumn)			
			' objWorksheetBI.Range(nColumn & "2:" & nColumn & llastrow).Copy objWorksheetS.Range(sColumn & "2:" & sColumn & llastrow)	
			' tColumn = tColumn + 1			
		' End If
	' Next
	
	' objWorkbookBI.Close SaveChanges = False
	' ' AutoFill de formulas
	' objWorksheetS.Range("A2").AutoFill(objWorksheetS.Range("A2:A" & llastrow))
	' objWorksheetS.Range("O2:T2").AutoFill(objWorksheetS.Range("O2:T" & llastrow))
	' objWorksheetS.Columns("G").Replace " 00:00:00+00:00", ""
	
	' ' **** Llenar hoja Redistribucion ****
	' ' Filtrar por criterios
	' objWorksheetS.Range("B1").AutoFilter 2, "1 CREDITS", 2, "2 WAITING FOR GATP"
	' objWorksheetS.Range("G1").AutoFilter 7, sDiaControl
	' ' Copiar y pegar celdas visibles
	' objWorksheetS.Range("B2:N" & llastrow).SpecialCells(12).Copy objWorksheetR.Range("B2:N" & llastrow)
	
	' ' **** Llenar hoja CDR ****
	' Set objWorkbookCDR = objExcel.Workbooks.Open(sCDRFile)
	' Set objWorksheetCDR = objWorkbookCDR.Worksheets("CDR")
	
	' llastrow = objWorksheetCDR.Cells(objWorksheetCDR.Rows.Count,2).End(-4162).Row
	
	' objWorksheetCDR.Range("A3:BI" & llastrow).Copy objWorksheetC.Range("A2:BI" & llastrow)
	MsgBox ""
	
	llastrow = objWorksheetS.Cells(objWorksheetS.Rows.Count,2).End(-4162).Row
	
	Const xlR1C1= -4150
	SourceAddress = "Status!" & objWorksheetS.Range("A1:T" & llastrow).Address(xlR1C1)

	Const xlDatabase = 1
	
	objWorkbookTemplate.Sheets("DinamicaStatus").Activate
	
	objExcel.ActiveSheet.PivotTables("TablaDinámica1").ChangePivotCache _
	objWorkbookTemplate.PivotCaches.Create(xlDatabase, SourceAddress, 6)
	
	objExcel.ActiveSheet.PivotTables("TablaDinámica2").ChangePivotCache _
	objWorkbookTemplate.PivotCaches.Create(xlDatabase, SourceAddress, 6)
	
	' Guardar Libro
	objWorkbookTemplate.Save
	objWorkbookTemplate.Close SaveChanges = True
	' Quitar objeto Excel de memoria
	objExcel.Quit
	MsgBox "0"
End Sub

Parameters = "C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\Control Cuota MT.xlsx#C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\PowerBI.csv#C:\Users\CSF5266\Downloads\OneDrive_2024-04-19\SCRIPTS QUOTA\CDR - Customer data report.xlsx#4/19/2024"
	
Call Quota(Parameters)