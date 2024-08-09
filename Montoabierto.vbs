Sub MontoAbierto(Parameters)
	arr = split(Parameters,"#")
	WorkbookPathData = arr(0)
  	TypeProcess  = arr(1)
	
	' Parameters = "C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de Cobranzas - URIEL.xlsx"
	
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
  	Set objWorksheetResumen = objWorkbookModule.Worksheets("Resumen de Cobranza")
	
	' On Error Resume Next
	
	Const xlPart = 2
	Const xlValues = -4163
	
	Dim m
	m = 0
  	tmp = 0
	
	Set mRangeModule = objWorksheetModule.Cells
	
	Dim dFind :	Set dFind = mRangeModule.Find("Monto Abierto",,xlValues,xlPart)
	
	If Not dFind Is Nothing Then
		firstAddress = dFind.Address
		Do
			m = objWorksheetModule.Cells(dFind.Row, dFind.Column + 8).value
			If TypeProcess = "Mixto" Then
				m = m / objWorksheetResumen.Cells(2,3).value
			End If
			
			Dim output, regex1, regex2
			Set regex1 = New RegExp
			Set regex2 = New RegExp
			regex1.Pattern = "(\d{3})"
			regex1.Global = True
			regex2.Pattern = ",$"
			output = regex1.Replace(StrReverse(CStr(Round(m,2))), "$1,")
			output = StrReverse(regex2.Replace(output, ""))
			
			tmp = tmp & "|" & output
			
			Set dFind =  mRangeModule.FindNext(dFind)
		Loop While Not dFind Is Nothing And dFind.Address <> firstAddress
	End If
	
    If tmp = "0" Then
		MontoAbiert = tmp
	Else
		MontoAbiert = Mid(tmp,3)
	End If
	
	MsgBox MontoAbiert
	
	objWorkbookModule.Save
	objWorkbookModule.Close SaveChanges = True
		
	objExcel.Quit
	
End Sub

Parameters= "C:\Users\CSF5266\Downloads\Formato de Reporte de Cobranza MANSION DEL CARIBE MIXTO.xlsx#"
Call MontoAbierto(Parameters)