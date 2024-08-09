Sub macro(Parameters)
  
  	Dim arr,strZEP1CompletePath,strMatrizSustitucionCompletePath
	'Parameters = "C:\DSDandCaseFillRate\Input/ZEP1.xlsx#C:\DSDandCaseFillRate\Input\Matriz_sustituciones.xlsx#C:\DSDandCaseFillRate\Input\Ricolino_Products.xlsx"
	
	arr=split(Parameters,"#")
	strZEP1CompletePath=arr(0)
	strMatrizSustitucionCompletePath=arr(1)
	strCatalogoRicolino = arr(2)
	'Opens the Excel file'
	Set objExcel = CreateObject("Excel.Application")
  	
  	objExcel.Application.Visible = False
	objExcel.Application.DisplayAlerts = False
  	objExcel.Application.ScreenUpdating = False

	Set objWorkbook2 = objExcel.Workbooks.Open(strMatrizSustitucionCompletePath)
	Set objWorksheet2 = objWorkbook2.Worksheets("sustitucion")
	Set objWorkbook3 = objExcel.Workbooks.Open(strCatalogoRicolino)
	Set objWorkbook1 = objExcel.Workbooks.Open(strZEP1CompletePath)
	Set objWorksheet1 = objWorkbook1.Worksheets("Sheet3")
	
	objWorksheet1.Cells.EntireColumn.AutoFit
	objWorksheet2.Cells.EntireColumn.AutoFit
	objWorksheet1.Columns("C:C").TextToColumns
	objWorksheet1.Columns("C:C").NumberFormat = "@"
  
	lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,4).End(-4162).Row
	Const xlWhole = 1
	Set fRange = objWorkbook3.Sheets("Grupos para CDS").Range("A:A")
	Set mRange = objWorkbook2.Sheets("SKU").Range("H:H")
	Set sRange = objWorkbook2.Sheets("sustitucion").Cells
	
	On Error Resume Next
	
	objWorksheet1.Cells(1,30).value = "Results"
	
	For i = 2 To lastRow
	
		If InStr(objWorksheet1.Cells(i,7).value,".") > 0 and InStr(objWorksheet1.Cells(i,7).value,".") <= 1 then
		  
			objWorksheet1.Cells(i,3).value = "BC"

		ElseIf InStr(objWorksheet1.Cells(i,7).value,".") > 1 then
				  
			objWorksheet1.Cells(i,3).value = "Drenado"
					
		Else
					
			valor = objWorksheet1.Cells(i,4).value
			plant = objWorksheet1.Cells(i,10).value
			
			objWorksheet2.Cells(1,4).value = plant		
			
			Dim sFind :	Set sFind = sRange.Find(Trim(valor),,,xlWhole)
			
			sustitucion = CStr(sFind)
			If Err.Number <> 0 Then
				columna2 = "BO"
				Err.Clear
			Else
				If sustitucion = valor then
				
					k = sFind.Row
					j = sFind.Column
							
					valor2 = objWorksheet2.Cells(k,j).value  -4163
					columna = Split(objWorksheet2.Cells(k,j).Address,"$")(1)
												  
					If columna = "B" then
						
						If IsNumeric(objWorksheet2.Cells(k,8).value) then
							columna2 = objWorksheet2.Cells(k,6).value
						Else
							If IsNumeric(objWorksheet2.Cells(k,12).value) then
								columna2 = objWorksheet2.Cells(k,10).value
							Else
								If IsNumeric(objWorksheet2.Cells(k,16).value) then
									columna2 = objWorksheet2.Cells(k,14).value
								Else
									columna2 = "BO"
								End If
							End If
						End If
					End If
									
					If columna = "F" then
						
						If IsNumeric(objWorksheet2.Cells(k,4).value) then
							columna2 = objWorksheet2.Cells(k,2).value
						Else
							If IsNumeric(objWorksheet2.Cells(k,12).value) then
								columna2 = objWorksheet2.Cells(k,10).value
							Else
								If IsNumeric(objWorksheet2.Cells(k,16).value) then
									columna2 = objWorksheet2.Cells(k,14).value
								Else
									columna2 = "BO"
								End If
							End If
						End If
					End If
									
					If columna = "J" then
					
						If IsNumeric(objWorksheet2.Cells(k,4).value) then
							columna2 = objWorksheet2.Cells(k,2).value
						Else
							If IsNumeric(objWorksheet2.Cells(k,8).value) then
								columna2 = objWorksheet2.Cells(k,6).value
							Else
								If IsNumeric(objWorksheet2.Cells(k,16).value) then
									columna2 = objWorksheet2.Cells(k,14).value
								Else
									columna2 = "BO"
								End If
							End If
						End If  
					End If
									
					If columna = "N" then
					
						If IsNumeric(objWorksheet2.Cells(k,4).value) then
							columna2 = objWorksheet2.Cells(k,2).value
						Else
							If IsNumeric(objWorksheet2.Cells(k,8).value) then
								columna2 = objWorksheet2.Cells(k,6).value
							Else
								If IsNumeric(objWorksheet2.Cells(k,12).value) then
									columna2 = objWorksheet2.Cells(k,10).value
								Else
									columna2 = "BO"
								End If
							End If
						End If
					End If

				End If					
			End If
			
  			   			
    
  			objWorksheet1.Cells(i,3).value = "'" & columna2
			
			If valor2 =  "" then
				objWorksheet1.Cells(i,3).value = "BO"
			End If
			valor2 = ""
					
			Dim vFind :	Set vFind = fRange.Find(Trim(valor),,,xlWhole)
			ricolino = CStr(vFind)
			If Err.Number <> 0 Then
				Err.Clear
			Else
				objWorksheet1.Cells(i,3).value = "BO"
			End If
			
			qConfirmed = objWorksheet1.Cells(i,7).value
			Dim mFind :	Set mFind = mRange.Find(Trim(valor),,,xlWhole)
			tmp = CStr(mFind)
			If Err.Number <> 0 then
				If qConfirmed <> 0 then
					objWorksheet1.Cells(i,3).value = "'" & qConfirmed
				End if
				Err.Clear
			Else
				
				multiplo = objWorkbook2.Sheets("SKU").Range(Replace(mFind.Address,"$","")).Offset(0,2).value
				result = qConfirmed MOD multiplo
				qCDS = qConfirmed - result
				
				If qCDS <> 0 then
					objWorksheet1.Cells(i,3).value = "'" & qCDS
				End If
			End If
			
			If objWorksheet1.Cells(i,3).value = "" Then
				objWorksheet1.Cells(i,3).value = "BO"
			End If
			
			j=0
			k=0
			
		End If
	Next 


	   'Saves and closes the Excel file'
	objWorkbook2.Save
	objWorkbook2.Close SaveChanges = True

	  'Saves and closes the Excel file'
	objWorkbook1.Save
	objWorkbook1.Close SaveChanges = True
	
	objWorkbook3.Save
	objWorkbook3.Close SaveChanges = True
		
	objExcel.Quit
  
  
End Sub
