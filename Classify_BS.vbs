Function Classify(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	arr = split(Parameters,"#")
	WorkbookPathData = arr(0)
	WorkbookPathModule = arr(1)
	amount = arr(2)
	
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbookData = objExcel.Workbooks.Open(WorkbookPathData,,True)
	Set objWorksheetData = objWorkbookData.Worksheets("Concepto MA")
	Set objWorkbookModule = objExcel.Workbooks.Open(WorkbookPathModule,,True)
	Set objWorksheetModule = objWorkbookModule.Worksheets("Modulo de Deuda")
	Set objWorksheetResumen = objWorkbookModule.Worksheets("Resumen de Cobranza")
	
  
	Const xlWhole = 1
	Const xlPart = 2
	Const xlValues = -4163
	Set mRangeModule0 = objWorksheetModule.Range("I1:BA300")
	
	Dim tFind :	Set tFind = mRangeModule0.Find("Total Notas de Cr",,xlValues,xlPart)
	columna = Split(objWorksheetModule.Cells(12,tFind.Column - 1).Address,"$")(1)
	
	Set mRangeModule = objWorksheetModule.Range("I12:" & columna & "300")
	Set mRangeData = objWorksheetData.Cells
	Set mRangeResumen = objWorksheetResumen.Range("B:B")
  
  	Set regex1 = New RegExp
	Set regex2 = New RegExp
	regex1.Pattern = "(\d{3})"
	regex1.Global = True
	regex2.Pattern = ",$"
	output = regex1.Replace(StrReverse(Int(amount)), "$1,")
	amount = StrReverse(regex2.Replace(output, ""))
	
	Dim mFind :	Set mFind = mRangeModule.Find(amount,,xlValues,xlPart)
	Dim iFind :	Set iFind = mRangeResumen.Find("n Bs.D de pagos reportados",,xlValues,xlPart)
	Dim sFind :	Set sFind = mRangeResumen.Find("DIF a favor o por cobrar Bs.D",,xlValues,xlPart)
	Dim pFind :	Set pFind = mRangeResumen.Find("DIF Bs.D total",,xlValues,xlPart)
	Dim concepto
	concepto = ""
	
	If Not iFind Is Nothing Then
		index = objWorksheetResumen.Cells(iFind.Row, iFind.Column + 1)
		index_tmp = index
		r = Round(index,2) - CDbl(Replace(amount,",",""))
		r = Round(r,2)
		If r >= -5 And r <= 5 Then
			concepto = "INDEX. 005-XXXXXX"
		end If
	End If
	
	' If Not sFind Is Nothing Then
		' saldo = objWorksheetResumen.Cells(sFind.row, sFind.Column + 1)
		' If saldo > 0 Then					
			' r = Round(saldo,2) - CDbl(Replace(amount,",",""))
			' r = Round(r,2)
			' If r >= -5 And r <= 5 Then
				' concepto = "AV. COBRO 005-XXXXXX"
			' End If
		' End If
			
		' If saldo < 0 Then
			' r = Round(saldo,2) + CDbl(Replace(amount,",",""))
			' r = Round(r,2)
			' If r >= -5 And r <= 5 Then
				' concepto = "SALDO A FAVOR 005-XXXXXX"
			' End if
		' End If			
	' End If
	
	If concepto = "" Then
		If Not pFind Is Nothing Then
			Index = objWorksheetResumen.Cells(iFind.Row, iFind.Column + 1)
			DIF_total = Abs(objWorksheetResumen.Cells(pFind.Row, pFind.Column + 1))
			SimDIF_total = Sgn(objWorksheetResumen.Cells(pFind.Row, pFind.Column + 1))
			DIF_FC = objWorksheetResumen.Cells(sFind.row, sFind.Column + 1)			
			r = Round(Abs(DIF_total),2) - CDbl(Replace(amount,",",""))
			r = Round(r,2)
			If r >= -5 And r <= 5 Then
				If DIF_total < Index And Index <> 0 And SimDIF_total = -1 Then
					concepto = "INDEX.P " & DIF_FC & " 005-XXXXXX"
				ElseIf DIF_total < Index Then
					concepto = "AV. COBRO 005-XXXXXX"
				ElseIf DIF_total > Index Then
					concepto = "AV. COBRO 005-XXXXXX"
				ElseIf DIF_total = Index And Index <> 0Then
					concepto = "INDEX. 005-XXXXXX"
				End If
			End If
			
		End If
	End If

	If concepto = "" Then
		If Not sFind Is Nothing Then
			Index = objWorksheetResumen.Cells(iFind.Row, iFind.Column + 1)
			DIF_FC = Abs(objWorksheetResumen.Cells(sFind.row, sFind.Column + 1))
			SimDIF_FC = Sgn(objWorksheetResumen.Cells(sFind.row, sFind.Column + 1))
			r = Round(Abs(DIF_FC),2) - CDbl(Replace(amount,",",""))
			r = Round(r,2)
			If r >= -5 And r <= 5 Then
				If DIF_FC > Index And SimDIF_FC = -1 Then
					concepto = "SALDO A FAVOR 005-XXXXXX"
				End If
			End If
			
		End If
	End If

	' sFolio = ""
	' If Not mFind Is Nothing Then
		' sConcept = objWorksheetModule.Cells(12,mFind.Column - 2).value
		' sFolio = objWorksheetModule.Cells(mFind.Row,3).value
		
		' For i = mFind.Row To mFind.Row + 6
			' If objWorksheetModule.Cells(i,mFind.Column - 1).value <> "" Then
				' sPercentage = objWorksheetModule.Cells(i,mFind.Column - 1).value * 100
			' End If
		' Next
		
		' Dim dFind :	Set dFind = mRangeData.Find(Trim(sConcept),,xlValues,xlWhole)
		
		' If Not dFind Is Nothing Then
			' sClassify = objWorksheetData.Cells(dFind.Row, dFind.Column + 1).value
			' If sFolio <> "" Then
					' sClassify = Replace(sClassify,"005-XXXXXX",sFolio)
			' End If	
		' End If
	' End If
	
	For k = 0 to 5
		saldo = CLng(amount) + k
		
		Set regex1 = New RegExp
		Set regex2 = New RegExp
		regex1.Pattern = "(\d{3})"
		regex1.Global = True
		regex2.Pattern = ",$"
		output = regex1.Replace(StrReverse(saldo), "$1,")
		saldo = StrReverse(regex2.Replace(output, ""))

		Set mFind = mRangeModule.Find(saldo,,xlValues,xlPart)
		
		
		sConcept = ""
		sFolio = ""
		If Not mFind Is Nothing Then
			sConcept = objWorksheetModule.Cells(12,mFind.Column - 2).value
			sFolio = objWorksheetModule.Cells(mFind.Row,3).value
			
			For i = mFind.Row To mFind.Row + 6
				If objWorksheetModule.Cells(i,mFind.Column - 1).value <> "" Then
					If IsNumeric(objWorksheetModule.Cells(i,mFind.Column - 1).value) Then
						sPercentage = objWorksheetModule.Cells(i,mFind.Column - 1).value * 100	
					End If
				End If
			Next
			
			Set dFind = mRangeData.Find(Trim(sConcept),,xlValues,xlWhole)							

			If Not dFind Is Nothing Then
				sClassify = objWorksheetData.Cells(dFind.Row, dFind.Column + 1).value
				
				If sFolio <> "" Then
					sClassify = Replace(sClassify,"005-XXXXXX",sFolio)
				End If				
			End If
		End If
	Next

	For k = 0 to 5
		saldo = CLng(amount) - k
		
		Set regex1 = New RegExp
		Set regex2 = New RegExp
		regex1.Pattern = "(\d{3})"
		regex1.Global = True
		regex2.Pattern = ",$"
		output = regex1.Replace(StrReverse(saldo), "$1,")
		saldo = StrReverse(regex2.Replace(output, ""))

		Set mFind = mRangeModule.Find(saldo,,xlValues,xlPart)
		
		sConcept = ""
		sFolio = ""
		If Not mFind Is Nothing Then
			sConcept = objWorksheetModule.Cells(12,mFind.Column - 2).value
			sFolio = objWorksheetModule.Cells(mFind.Row,3).value
			
			For i = mFind.Row To mFind.Row + 6
				If objWorksheetModule.Cells(i,mFind.Column - 1).value <> "" Then
					If IsNumeric(objWorksheetModule.Cells(i,mFind.Column - 1).value) Then
						sPercentage = objWorksheetModule.Cells(i,mFind.Column - 1).value * 100	
					End If
				End If
			Next
				
			Set dFind = mRangeData.Find(Trim(sConcept),,xlValues,xlWhole)							

			If Not dFind Is Nothing Then
				sClassify = objWorksheetData.Cells(dFind.Row, dFind.Column + 1).value
				If sFolio <> "" Then
					sClassify = Replace(sClassify,"005-XXXXXX",sFolio)
				End If				
			End If
		End If
	Next

	If concepto = "" Then
		Classify = Replace(sClassify,"__",sPercentage)
	Else
		Classify = concepto
	End If
	
	' Saves and closes the Excel file'
	objWorkbookData.Save
	objWorkbookData.Close SaveChanges = True
	objWorkbookModule.Save
	objWorkbookModule.Close SaveChanges = True
	
	' Quita la instancia del objeto Excel
	objExcel.Quit
	
End Function

Parameters = "C:\CuentasPorCobrarVnzla\Output\Datos basicos.xlsx#C:\CuentasPorCobrarVnzla\Temp\AUTOMERCADO66973.xlsx#150"

MsgBox Classify(Parameters)