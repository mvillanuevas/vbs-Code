Function CreditNote(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)
	
	'WorkbookPath = "C:\CuentasPorCobrarVnzla\Temp\BS - COMERCIALIZADORA CHOCOMAYOR.xlsx"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = True
	
	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath,0)
	Set objWorksheetDIF = objWorkbook.Worksheets("Resumen de Cobranza")
	Set objWorksheetPago = objWorkbook.Worksheets("Reporte del Pago")
	Set objWorksheetModulo = objWorkbook.Worksheets("Modulo de Deuda")
	
	llastrow = objWorksheetPago.Cells(objWorksheetPago.Rows.Count,6).End(-4162).Row
	bMoneda = False
	bIGTF = False
	
	For i = 11 to llastrow
		moneda = objWorksheetPago.Cells(i,2).value
		pIGTF = objWorksheetPago.Cells(i,3).value
		
		If InStr(moneda,"lares") <> 0 Then
			bMoneda = True
		End If
		
		If pIGTF = "No" Then
			bIGTF = True
		End If
		
	Next
	
	Const xlWhole = 1
	Const xlPart = 2
	Const xlValues = -4163
	Set mRangeDIF = objWorksheetDIF.Cells
	
	Dim mFind :	Set mFind = mRangeDIF.Find("dito Bs.D",,xlValues,xlPart)
	Dim mFindDIF : Set mFindDIF = mRangeDIF.Find("DIF Bs.D total",,,xlWhole)
	Dim mFindDeduda : Set mFindDeduda = mRangeDIF.Find("Deuda pendiente en USD",,,xlWhole)
	
	If bMoneda And bIGTF Then
		DIF = (objWorksheetDIF.Cells(mFindDeduda.Row,mFindDeduda.Column+1).value - objWorksheetModulo.Cells(8,10).value) * objWorksheetDIF.Cells(2,3).value
	Else
		DIF = Replace(objWorksheetDIF.Cells(mFindDIF.Row,mFindDIF.Column+1).value,"-","")
	End If
	
	NC = objWorksheetDIF.Cells(mFind.Row,mFind.Column+1).value

	NC = CStr(Round(CDbl(NC),2))
	vDif = CStr(Round(CDbl(DIF),2))
	
	Set regex1 = New RegExp
	Set regex2 = New RegExp
	regex1.Pattern = "(\d{3})"
	regex1.Global = True
	regex2.Pattern = ",$"

	output = regex1.Replace(StrReverse(NC), "$1,")
	output = StrReverse(regex2.Replace(output, ""))
	
	output2 = regex1.Replace(StrReverse(vDif), "$1,")
	output2 = StrReverse(regex2.Replace(output2, ""))
	
	If bMoneda And bIGTF Then
		CreditNote = Replace(output,"-","") & "-|" & Replace(output2,"-","") & "-"
	Else
		CreditNote = Replace(output,"-","") & "-|" & Replace(output2,"-","")
	End If

	   'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Function

Parameters="C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de Cobranzas - FERRETERIA EPA 005-357626 005-357630.xlsx"
MsgBox CreditNote(Parameters)