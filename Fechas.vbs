Function Fechas(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)
	Document = arr(1)
	
	'WorkbookPath = "C:\CuentasPorCobrarVnzla\Temp\BS - COMERCIALIZADORA CHOCOMAYOR.xlsx"
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False
	
	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
	Set objWorksheet = objWorkbook.Worksheets("Sheet1")
	
  	Const xlValues = -4163
  	Const xlPart = 2
  
  	Set mRange = objWorksheet.Range("F:F")
	
	Dim mFind :	Set mFind = mRange.Find(Trim(Document),,xlValues,xlPart)
	
	If Not mFind Is Nothing Then
		DocDate = objWorksheet.Cells(mFind.Row, mFind.Column + 4).value
		PstngDate = objWorksheet.Cells(mFind.Row, mFind.Column + 7).value
		PmntgDate = objWorksheet.Cells(mFind.Row, mFind.Column + 8).value
	End If
	
	Fechas = DocDate & "|" & PstngDate & "|" & PmntgDate
	
   'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Function

Parameters="C:\Users\CSF5266\Downloads\CuentasPorCobrarVnzla1\Output\FBL5N.xlsx#5893196744"
MsgBox Fechas(Parameters)