Sub macro'(Parameters)
  
	Parameters = "C:\CuentasPorCobrarVnzla\Output\Formato Indexacion.xlsx#156162624#URIEL C.A.#005-312860#5893054476#4/21/2023#87,829.91#75,715.44##4#24.5739#24.5763747083234#4/17/2023#4/17/2023#4/22/2023#C:\CuentasPorCobrarVnzla\Temp\Formato de Reporte de Cobranzas - URIEL.xlsx#C:\CuentasPorCobrarVnzla\Output\Maestro Clientes.xlsx"
  	arr=split(Parameters,"#")
	sIndexacionFilePath=arr(0)
	sCodClient=arr(1)
	sClient=arr(2)
	sInvoiceNumber_FormatoReporte=arr(3)
	sInvoiceSAPNumber=arr(4)
	sExcelDate=arr(5)
	sAmountInvoice_Bs=arr(6)
	sAmountBaseImponible=arr(7)
	sAmountNC3=arr(8)
	llastrow=arr(9)
	sEmissionRateAverage_BCV=arr(10)
	sPaymentRateAverage_BCV=arr(11)
	sEmissionDate=arr(12)
	sDeliveryDate=arr(13)
	sExpirationDate=arr(14)
	WorkbookPath=arr(15)
 	WorkbookPathMaestro=arr(16)
  
    'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")
	
	'Parámetro para indicar si se quiere visible la aplicación de Excel
	objExcel.Application.Visible = False
	'Parámetro evitar mostrar pop ups de Excel
	objExcel.Application.DisplayAlerts = False
	'Evita movimiento de pantalla
	objExcel.Application.ScreenUpdating = False

    Set objWorkbook = objExcel.Workbooks.Open(sIndexacionFilePath)
	Set objWorksheet = objWorkbook.Worksheets(1)
	Set objWorkbookR = objExcel.Workbooks.Open(WorkbookPath)
	Set objWorksheetR = objWorkbookR.Worksheets("Resumen de Cobranza")
  	Set objWorkbookMaestro = objExcel.Workbooks.Open(WorkbookPathMaestro)
	Set objWorksheetMaestro = objWorkbookMaestro.Worksheets("MAESTRO")
	
	Const xlValues = -4163
  	Const xlWhole = 2
  
  	Dim sRegion
  	sRegion = ""
	
	Set mRangeMaestro = objWorksheetMaestro.Range("A:A")
	
  	Dim mFind :	Set mFind = mRangeMaestro.Find(Trim(sCodClient),,xlValues,xlWhole)
	
	If Not mFind Is Nothing Then
		sRegion = objWorksheetMaestro.Cells(mFind.Row, mFind.Column + 6)
	End If
  	
  	objWorkbookMaestro.Save
  	objWorkbookMaestro.Close SaveChanges = True
  
	
	Set mRange = objWorksheetR.Cells
	
	Dim rFind :	Set rFind = mRange.Find("Bs.D de pagos reportados",,xlValues,xlWhole)
	Dim tFind :	Set tFind = mRange.Find("DIF Bs.D total",,xlValues,xlWhole)
	Dim pFind :	Set pFind = mRange.Find("DIF a favor o por cobrar Bs.D",,xlValues,xlWhole)
	

	If Not rFind Is Nothing Then
		rSgn = Sgn(objWorksheetR.Cells(rFind.Row,rFind.Column+1).value)
	End If

	If Not tFind Is Nothing Then
		tSgn = Sgn(objWorksheetR.Cells(tFind.Row,tFind.Column+1).value)
	End If
	
	If Not pFind Is Nothing Then
		pSgn = Sgn(objWorksheetR.Cells(pFind.Row,pFind.Column+1).value)
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
		FindIndex = ""
	End If  
  
  	objWorkbookR.Save
  	objWorkbookR.Close SaveChanges = True 

    objWorksheet.Range("A" & llastrow).value = sRegion
  	objWorksheet.Range("B" & llastrow).value = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  	objWorksheet.Range("C" & llastrow).value = sCodClient
    objWorksheet.Range("D" & llastrow).value = sClient
    objWorksheet.Range("E" & llastrow).value = sInvoiceNumber_FormatoReporte
    objWorksheet.Range("F" & llastrow).value = sInvoiceSAPNumber
    objWorksheet.Range("G" & llastrow).value = sEmissionDate
    objWorksheet.Range("H" & llastrow).value = sDeliveryDate
    objWorksheet.Range("I" & llastrow).value = sExpirationDate
    objWorksheet.Range("J" & llastrow).value = sExcelDate
    objWorksheet.Range("K" & llastrow).value = sAmountInvoice_Bs
    objWorksheet.Range("L" & llastrow).value = sAmountBaseImponible
    objWorksheet.Range("N" & llastrow).value = sAmountNC3     
    objWorksheet.Range("Q" & llastrow).value = sEmissionRateAverage_BCV     
    objWorksheet.Range("S" & llastrow).value = sPaymentRateAverage_BCV     
    objWorksheet.Range("Z" & llastrow).value = FindIndex
  
   'Saves and closes the Excel file'
   	objWorkbook.Save
  	objWorkbook.Close SaveChanges = True 
  
	objExcel.Quit
   
   
End Sub

call macro