Function LeyendaMA(Parameters)

	' WorkbookPathModule = "C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de Cobranzas - URIEL.xlsx"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbookModule = objExcel.Workbooks.Open(Parameters)
	Set objWorksheetModule = objWorkbookModule.Worksheets("Resumen de Cobranza")

	
	Const xlPart = 2
	Const xlValues = -4163
	
	Set mRangeModule = objWorksheetModule.Range("B:B")
	
	Dim dFind :	Set dFind = mRangeModule.Find("DPP Bs.D",,xlValues,xlPart)
	Dim pFind :	Set pFind = mRangeModule.Find("Bs.D de pagos reportados",,xlValues,xlPart)
	Dim cFind :	Set cFind = mRangeModule.Find("DIF a favor o por cobrar Bs.D",,xlValues,xlPart)
	Dim tFind :	Set tFind = mRangeModule.Find("DIF Bs.D total",,xlValues,xlPart)
	
	If Not dFind Is Nothing Then
		dpp = objWorksheetModule.Cells(dFind.Row, dFind.Column + 1).value
		If InStr(dpp,"-") = 0 Then
			dpp = "-" & dpp
		End If
		dpp = Round(CDbl(dpp),2)
		If dpp = 0 Then
			dpp = ""
		End If
	End If
	
	If Not pFind Is Nothing Then
		index = objWorksheetModule.Cells(pFind.Row, pFind.Column + 1).value
		indextmp = index
		If InStr(index,"-") <> 0 Then
			index = 0
		End If
		index = Round(CDbl(index),2)
		If index = 0 Then
			index = ""
		End If
	End If

	If Not cFind Is Nothing Then
		dif = objWorksheetModule.Cells(cFind.Row, cFind.Column + 1).value
		diftmp = dif
		If InStr(dif,"-") <> 0 Then
			dif = Replace(dif,"-","")
		Else
			dif = "-" & dif
		End If
		dif = Round(CDbl(dif),2)
		If dif = 0 Then
			dif = ""
		End If
	End If
	
	If Sgn(indextmp) = 1 And Sgn(diftmp) = 1 Then
		index = objWorksheetModule.Cells(tFind.Row, tFind.Column + 1).value
		dif = ""
		If InStr(index,"-") <> 0 Then
			index = Replace(index,"-","")
		Else
			index = "-" & index
		End If
	End If
	
	LeyendaMA = index & "|" & dif
	
	objWorkbookModule.Save
	objWorkbookModule.Close SaveChanges = True
		
	objExcel.Quit

End Function

Parameters = "C:\Users\CSF5266\Downloads\Formato de Reporte de Cobranzas - P05 2024 NACIONAL (1).xlsx"
MsgBox LeyendaMA(Parameters)