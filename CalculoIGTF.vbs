Function CalculoIGTF(SetValues)
	
	' SetValues = "C:\CuentasPorCobrarVnzla\Output\IGTF BCO 2024.xlsx#1era Quincena Mayo#1400001425"
	
	arr=split(SetValues,"#")
	sIGTFilePath=arr(0)
	sPeriod=arr(1)
	sDocument=arr(2)
	
	'Opens the Excel file'
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

	Set objWorkbook = objExcel.Workbooks.Open(sIGTFilePath)
	Set objWorksheet = objWorkbook.Worksheets(sPeriod)
	
	Const xlValues = -4163
  	Const xlPart = 2
	Dim bFlag
	bFlag = False
	
	Set mRangeIGTF = objWorksheet.Cells

	Dim mFind :	Set mFind = mRangeIGTF.Find(Trim(sDocument),,xlValues,xlPart)
	
	If Not mFind Is Nothing Then
		rDocument = mFind.Row
		CalculoIGTF = objWorksheet.Cells(rDocument,12).value
	Else
		CalculoIGTF = "No Exist"
	End If

	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	objExcel.Quit

End Function
Call CalculoIGTF

C:\CuentasPorCobrarVnzla\Output/BCO NOVIEMBRE BS 2023.xlsx#806#100135314#AUTOMERCADO EXPRESS 2707 #04/04/2024#04/04/2024#BOT#PROVINCIAL