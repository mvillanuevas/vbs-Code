Sub FormatTable(WorkBookF)

	Const xlOpenXMLWorkbook = 51
	Const xlYes = 1
	Const xlSrcRange = 1

	'Opens the Excel file'
	Set objExcel = CreateObject("Excel.Application")
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

	Set objWorkbookR = objExcel.Workbooks.open(WorkBookF)
	Set objWorksheet = objWorkbookR.worksheets(1)

	'create a new listobject from the Range with top-left=A1
	objWorksheet.ListObjects.Add(xlSrcRange, objWorksheet.Range("A1").CurrentRegion, , xlYes).Name = "TableReport"

	objWorkbookR.Save
  	objWorkbookR.Close SaveChanges = True

	objExcel.Quit
End Sub

Call FormatTable("C:\Users\CSF5266\Desktop\Book1.xlsx")