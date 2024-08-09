Sub NCAbierto(Parameters)
	arr = split(Parameters,"#")
	WorkbookPathData = arr(0)
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbookModule = objExcel.Workbooks.Open(WorkbookPathData,,True)
	Set objWorksheetModule = objWorkbookModule.Worksheets("Reporte del Pago")
	
	Const xlPart = 2
	Const xlValues = -4163	
	Dim m
	m = 0
	tmp = "0"
	
	Set mRangeModule = objWorksheetModule.Cells
	
	Dim dFind :	Set dFind = mRangeModule.Find("Nota de Cr",,xlValues,xlPart)
	
	If Not dFind Is Nothing Then
		firstAddress = dFind.Address
		Do
			m = objWorksheetModule.Cells(dFind.Row, dFind.Column + 4).value			
			tmp = tmp & "|" & m
			
			Set dFind =  mRangeModule.FindNext(dFind)
		Loop While Not dFind Is Nothing And dFind.Address <> firstAddress
	End If
	
	If tmp = "0" Then
		MontoAbierto = tmp
	Else
		MontoAbierto = Mid(tmp,3)
	End If
	Msgbox MontoAbierto
	objWorkbookModule.Save
	objWorkbookModule.Close SaveChanges = True
		
	objExcel.Quit
	
End Sub

Parameters="C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de CENTRO DE DISTRIBUCIONES FRANCIS MIXTO 005-340258-005-340259.xlsx"
Call NCAbierto(Parameters)