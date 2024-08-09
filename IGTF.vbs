Sub macro(SetValues) 
  
  	arr=split(SetValues,"#")
	sIGTFilePath=arr(0)
	sPeriod=arr(1)
	sFecha=arr(2)
	sAssigment=arr(3)
	sMonto=arr(4)
	sDocSap=arr(5)
	sIGTF=arr(6)
  	sCodClient=arr(7)
	sInvoice=arr(8)
	
	
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
	
	Dim mFind :	Set mFind = mRangeIGTF.Find(Trim(sMonto),,xlValues,xlPart)
	
	
	If Not mFind Is Nothing Then
		rMonto = mFind.Row
	End If

	' Cliente
	If Trim(sCodClient) = Trim(objWorksheet.Cells(rMonto,1).value) Then
		bFlag = True
	End If
	' Assigment
	If Trim(sAssigment) = Trim(objWorksheet.Cells(rMonto,7).value) Then
		bFlag = True
	End If
	
	If bFlag Then
		objWorksheet.Cells(rMonto,4).value = sDocSap
		objWorksheet.Cells(rMonto,14).value = sInvoice
		If sIGTF = "No Aplica" Then
			sIGTF = "NA"
		End If
		objWorksheet.Cells(rMonto,13).value = UCase(sIGTF)
    	objWorksheet.Cells(rMonto,17).value = "Robot"
	End If
	
	'Saves and closes the Excel file'
	objWorkbook.Save
	objWorkbook.Close SaveChanges = True
	
	objExcel.Quit	
  
End Sub