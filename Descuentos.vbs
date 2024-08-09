Function Descuentos(Parameters)
	' Parameters = "C:\Users\CSF5266\Desktop\Formato de Reporte de SIN LIMITES $.xlsx"
	
	arr=split(Parameters,"#")
	sReportePath=arr(0)
	'Opens the Excel file'
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

	Set objWorkbook = objExcel.Workbooks.Open(sReportePath)
	Set objWorksheetR = objWorkbook.Worksheets("Resumen de Cobranza")
	Set objWorksheetM = objWorkbook.Worksheets("Modulo de Deuda")
	Set objWorksheetPago = objWorkbook.Worksheets("Reporte del Pago")
	
	llastrow = objWorksheetPago.Cells(objWorksheetPago.Rows.Count,6).End(-4162).Row
	bMoneda = False
	bIGTF = False
	
	For i = 11 to llastrow
		moneda = objWorksheetPago.Cells(i,2).value
		pIGTF = objWorksheetPago.Cells(i,3).value
		
		If InStr(LCase(moneda),"lares") <> 0 Then
			bMoneda = True
		End If
		
		If InStr(UCase(pIGTF),"N") <> 0 Then
			bIGTF = True
		End If
		
	Next
	
	Dim Desc
	
	Const xlPart = 2
	Const xlValues = -4163
	Set mRangeDIF = objWorksheetR.Range("B:B")
	
  	Dim mDif : Set mDif = mRangeDIF.Find("DIF Bs.D total",,xlValues,xlPart)

	Dif = objWorksheetR.Range("C" & mDif.row).value
	TasaPromedio = objWorksheetR.Range("C2").value
	IGTF_USD = Round(objWorksheetM.Range("J8").value,2)
	
	lastRow = objWorksheetM.Cells(objWorksheetM.Rows.Count,3).End(-4162).Row
	Const xlWhole = 1

	Set mRangeDIF = objWorksheetM.Rows(12)
	
	Dim mFind :	Set mFind = mRangeDIF.Find("Total Notas de Cr",,xlValues,xlPart)
	If Not mFind Is Nothing Then
		lastColumn = mFind.Column - 1
	Else
		lastColumn = 39
	End If

	For i = 13 to lastRow Step 7		
		If Not IsEmpty(objWorksheetM.Cells(i, 3).value) Then
			For j = 13 to lastColumn Step 4
				If objWorksheetM.Cells(i, j).value <> 0 Then
					Desc = Desc & "|" & Round(objWorksheetM.Cells(i, j).value/TasaPromedio,2) & "-"
				End If
			Next
		End If
	Next

	For i = 13 to lastRow Step 7		
		If Not IsEmpty(objWorksheetM.Cells(i, 3).value) Then
			'For j = 10 to 53 Step 4
				If objWorksheetM.Cells(i, 10).value <> 0 Then
					Desc = Desc & "|" & Round(objWorksheetM.Cells(i, 10).value/TasaPromedio,2) & "-"
				End If
			'Next
		End If
	Next
	
	If bMoneda And bIGTF Then
		IGTF_USD = ""
		Dif = 0
	End If
	
	SaldoFavor = 0
	If Dif = 0 Then
		SaldoFavor = 0
	Else
		If Sgn(Round(Dif,2)) = -1 Then
			SaldoFavor = Replace(Round(Dif,2),"-","")
		Else
			SaldoFavor = Round(Dif,2)
		End If
	End If
	
	
	
	Tmp = Desc & "|" & IGTF_USD & "|" & SaldoFavor
	Descuentos = Mid(Tmp,2)
	
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	objExcel.Quit
	
End Function