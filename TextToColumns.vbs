Sub f_TextToColumns(Parameters)

	'Debido a que AA solo admite pasar un parámetro de entrada,
	'se pasan todos los parámetros necesarios separados por #
	'haciendo un split de cada parámetro
	
	arr = split(Parameters,"#")
	WorkbookPath = arr(0)	' C:\Input\ZEP1.xlsx
	SheetName = arr(1)	' Sheet1
	ColumnName = arr(2) 	' B:B
	
	'Genera un objeto de tipo Excel Application
	Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = True
	'Paramatro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	
	'Abre libro Excel
	Set objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
	'Se situa en la hoja deseada
	Set objWorkSheet = objWorkbook.Worksheets(SheetName)
	
	'Aplica Text to columns en formate General
	objWorkSheet.Range(ColumnName).TextToColumns
	
	'Guarda y cierra el libro Excel
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True

	'Quita la instancia del objeto Excel
	objExcel.Quit
	
End Sub