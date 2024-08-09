Sub BCO(Parameters)

	arr = split(Parameters,"#")
	WorkbookPathBCO = arr(0)
	CodCliente = arr(1)
	Cliente = arr(2)
	FReporte = arr(3)
	CodSAP = arr(4)
  	BReference = arr(5)
  	BCOSheet=arr(6)
	
	' WorkbookPathBCO = "C:\CuentasPorCobrarVnzla\Output\BCO ABRIL BS 2023.xlsx"
	' BCOSheet = "MERCANTIL"
	' CodCliente = "1234"
	' Cliente = "12345"
	' FReporte = "1234"
	' CodSAP = "1234"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbookBCO = objExcel.Workbooks.Open(WorkbookPathBCO)
	
	nWorkSheets = objWorkbookBCO.Worksheets.Count
	
	'Itera sobre cada hoja del libro
	For i = 1 To nWorkSheets
		If InStr(objWorkbookBCO.Worksheets(i).Name, Trim(BCOSheet)) <> 0 Then
			BCOSheet = objWorkbookBCO.Worksheets(i).Name
			Exit For
		End If				
	Next
	
	Set objWorksheetBCO = objWorkbookBCO.Worksheets(BCOSheet)
	
	Const xlPart = 2
	Const xlWhole = 1
	Const xlValues = -4163
	
	Set mRangeBCO = objWorksheetBCO.Rows(1)
  	Set bRangeBCO = objWorksheetBCO.Range("B:B")
	
  	Dim bFind :	Set bFind = bRangeBCO.Find(BReference,,xlValues,xlPart)
	Dim rFind :	Set rFind = mRangeBCO.Find("REFERENCIA",,xlValues,xlPart)
	Dim ccFind : Set ccFind = mRangeBCO.Find("CODIGO CLIENTE",,xlValues,xlPart)
	Dim cFind :	Set cFind = mRangeBCO.Find("CLIENTE",,xlValues,xlWhole)
	Dim fFind :	Set fFind = mRangeBCO.Find("FECHA REPORTE VENTAS",,xlValues,xlPart)
	Dim frFind : Set frFind = mRangeBCO.Find("FECHA REGISTRO COBRANZA",,xlValues,xlPart)
	Dim sFind : Set sFind = mRangeBCO.Find("SAP",,xlValues,xlPart)
	Dim aFind : Set aFind = mRangeBCO.Find("ANALISTA",,xlValues,xlPart)
	

	' -------------------------------------------------------------------------------------------------------
  	If Not bFind Is Nothing Then
    	rlast = bFind.Row
  	Else
		rlast = 2
	End If
	
	If Not ccFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,ccFind.Column).value = CodCliente
	End If
	
	If Not cFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,cFind.Column).value = Cliente
	End If
	
	If Not fFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,fFind.Column).value = FReporte
	End If
	
	If Not frFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,frFind.Column).value = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
	End If
	
	If Not sFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,sFind.Column).value = CodSAP
	End If
	
	If Not aFind Is Nothing Then
		objWorksheetBCO.Cells(rlast,aFind.Column).value = "Robot"
	End If
	
	objWorkbookBCO.Save
	objWorkbookBCO.Close SaveChanges = True
	
	' Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub

Parameters = "C:\Users\CSF5266\Downloads\CuentasPorCobrarVnzla2\Output\BCO JUNIO BS 2024.xlsx#12345#YO###123456#BANESCO"

Call BCO(Parameters)