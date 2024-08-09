Call FindIndex
Function FindIndex(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)
	
	' WorkbookPath = "C:\Users\CSF5266\Downloads\BS - INDEX NO AUTOMERCADO LUZ 005-315080 .xlsx"
	
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
	Set objWorksheet = objWorkbook.Worksheets("Resumen de Cobranza")
	
	On Error Resume Next
	
	Const xlWhole = 2
	Const xlValues = -4163
	Set mRange = objWorksheet.Cells
	
	Dim rFind :	Set rFind = mRange.Find("Bs.D de pagos reportados",,xlValues,xlWhole)
	Dim tFind :	Set tFind = mRange.Find("DIF Bs.D total",,xlValues,xlWhole)
	Dim pFind :	Set pFind = mRange.Find("DIF a favor o por cobrar Bs.D",,xlValues,xlWhole)
	
	tmp = CStr(rFind)
	If Err.Number <> 0 Then
		Err.Clear
	Else
		rSgn = Sgn(objWorksheet.Cells(rFind.Row,rFind.Column+1).value)
	End If
	
	tmp = CStr(tFind)
	If Err.Number <> 0 Then
		Err.Clear
	Else
		tSgn = Sgn(objWorksheet.Cells(tFind.Row,tFind.Column+1).value)
	End If
	
	tmp = CStr(pFind)
	If Err.Number <> 0 Then
		Err.Clear
	Else
		pSgn = Sgn(objWorksheet.Cells(pFind.Row,pFind.Column+1).value)
	End If

	
	If rSgn = 1 And tSgn = -1 And pSgn = -1 Then
		FindIndex = "SI"
	ElseIf rSgn = 1 And tSgn = -1 And pSgn = 0 Then
		FindIndex = "SI"
	ElseIf rSgn = 1 And tSgn = 1 And pSgn = 1 Then
		FindIndex = "NO"
	ElseIf rSgn = -1 And tSgn = -1 And pSgn = -1 Then
		FindIndex = "NA"
	ElseIf rSgn = -1 And tSgn = 1 And pSgn = 1 Then
		FindIndex = "NA"
	ElseIf rSgn = 0 And tSgn = 1 And pSgn = 1 Then
		FindIndex = "NA"
	ElseIf rSgn = 0 And tSgn = -1 And pSgn = -1 Then
		FindIndex = "NA"
	ElseIf rSgn = 1 And tSgn = -1 And pSgn = 1 Then
		FindIndex = "PARCIAL"
	Else
		FindIndex = "Valor no encontrado"
	End If
	
End Function