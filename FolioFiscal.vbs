Sub CreditNote'(Parameters)

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
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
	Set objWorksheetDIF = objWorkbook.Worksheets("Resumen de Cobranza")
		
	Const xlWhole = 1
	Set mRangeDIF = objWorksheetDIF.Cells
	
	Dim mFind :	Set mFind = mRangeDIF.Find("Notas de Crédito Bs.D",,,xlWhole)
	Dim mFindDIF :	Set mFindDIF = mRangeDIF.Find("DIF Bs.D total",,,xlWhole)
	Dim NC
	
	DIF = Replace(objWorksheetDIF.Cells(mFindDIF.Row,mFindDIF.Column+1).value,"-","")
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
	
	CreditNote = Replace(output,"-","") & "-|" & Replace(output2,"-","")
	
	   'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub

Call CreditNote