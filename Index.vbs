Parameters = "C:\Users\CSF5266\Downloads\CuentasPorCobrarVnzla1\Output\Formato Indexacion.xlsx####19######C:\Users\CSF5266\Downloads\CuentasPorCobrarVnzla1\Temp\PRACTIMERCADO MINUTO LOS PALOS GRANDES 005-360253 - 005-360254 - 005-360255.xlsx#C:\Users\CSF5266\Downloads\CuentasPorCobrarVnzla1\Output\Maestro Clientes.xlsx#"
Call macro(Parameters)

Sub macro(Parameters)
  
  	arr=split(Parameters,"#")
	sIndexacionFilePath=arr(0)
	sCodClient=arr(1)
	sClient=arr(2)
	sExcelDate=arr(3)
	llastrow=arr(4)
	sEmissionRateAverage_BCV=arr(5)
	sPaymentRateAverage_BCV=arr(6)
	sEmissionDate=arr(7)
	sDeliveryDate=arr(8)
	sExpirationDate=arr(9)
	WorkbookPath=arr(10)
	WorkbookPathMaestro=arr(11)
	sAmountNC3=arr(12)
  
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
	Set objWorksheetM = objWorkbookR.Worksheets("Modulo de Deuda")
  	Set objWorkbookMaestro = objExcel.Workbooks.Open(WorkbookPathMaestro)
	Set objWorksheetMaestro = objWorkbookMaestro.Worksheets("MAESTRO")
	
	Const xlValues = -4163
  	Const xlWhole = 2
	Const xlPart = 2
  
  	Dim sRegion
  	sRegion = ""
	
	Set mRangeMaestro = objWorksheetMaestro.Range("A:A")
	
  	Dim mFind :	Set mFind = mRangeMaestro.Find(Trim(sCodClient),,xlValues,xlWhole)
	
	If Not mFind Is Nothing Then
		sRegion = objWorksheetMaestro.Cells(mFind.Row, mFind.Column + 3)
	End If
  	
  	objWorkbookMaestro.Save
  	objWorkbookMaestro.Close SaveChanges = True
  
	
	Set mRange = objWorksheetR.Cells
	
	Dim rFind :	Set rFind = mRange.Find("Bs.D de pagos reportados",,xlValues,xlWhole)
	Dim tFind :	Set tFind = mRange.Find("DIF Bs.D total",,xlValues,xlWhole)
	Dim pFind :	Set pFind = mRange.Find("DIF a favor o por cobrar Bs.D",,xlValues,xlWhole)
	
	Set nRage = objWorksheetM.Rows(12)
	
	Dim mModulo : Set mModulo = nRage.Find("Total Notas de Cr",,xlValues,xlPart)
	

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
	ElseIf rSgn = -1 And tSgn = -1 Then
		FindIndex = "NA"
	ElseIf rSgn = -1 And tSgn = 1 Then
		FindIndex = "NA"
	ElseIf rSgn = 0 And tSgn = 1 Then
		FindIndex = "NA"
	ElseIf rSgn = 0 And tSgn = -1 Then
		FindIndex = "NA"
	ElseIf rSgn = 0 And tSgn = 0 And pSgn = 0 Then
		FindIndex = "NA"
	ElseIf rSgn = -1 And tSgn = 0 And pSgn = -1 Then
		FindIndex = "NA"
	ElseIf rSgn = 1 And tSgn = -1 And pSgn = 1 Then
		FindIndex = "PARCIAL"
	Else
		FindIndex = ""
	End If  
	
	lastRow = objWorksheetM.Cells(objWorksheetM.Rows.Count,6).End(-4162).Row

	For i = 14 to lastRow Step 7
		sInvoiceNumber_FormatoReporte = objWorksheetM.Cells(i - 1,3).value
		sInvoiceSAPNumber = objWorksheetM.Cells(i - 1,4).value
		If Not IsEmpty(objWorksheetM.Cells(i - 1,3).value) Then
			For j = i to i + 6
				sAmountInvoice_Bs = objWorksheetM.Cells(j,6).value
				sAmountBaseImponible = objWorksheetM.Cells(j,7).value
				sAmountNC3 = objWorksheetM.Cells(j,mModulo.Column).value
				If sAmountBaseImponible <> 0 And Trim(objWorksheetM.Cells(j,3).value) = "" Then
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
					
					If FindIndex = "PARCIAL" Then
						objWorksheet.Range("AA" & llastrow).value = objWorksheetR.Cells(pFind.Row,pFind.Column+1).value
					End If
      				objWorksheet.Range("AC" & llastrow).value = "Bot"
					llastrow = llastrow + 1
				End If			
			Next
		End If
	Next
  
  	objWorkbookR.Save
  	objWorkbookR.Close SaveChanges = True 

    
  
   'Saves and closes the Excel file'
   	objWorkbook.Save
  	objWorkbook.Close SaveChanges = True 
  
	objExcel.Quit   
   
End Sub

