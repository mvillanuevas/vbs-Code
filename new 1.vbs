Sub macro
  
	Dim arr,strZEP1CompletePath,strMatrizSustitucionCompletePath
	
	Parameters = "C:\DSDandCaseFillRate\Input\base_xd.xlsm#C:\DSDandCaseFillRate\Input\Matriz_sustituciones.xlsx"
	
  arr=split(Parameters,"#")
  strZEP1CompletePath=arr(0)
  strMatrizSustitucionCompletePath=arr(1)
  
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkbook1 = objExcel.Workbooks.Open(strZEP1CompletePath)
  Set objWorksheet1 = objWorkbook1.Worksheets("Sheet3")
  	Set objWorkbook2 = objExcel.Workbooks.Open(strMatrizSustitucionCompletePath)
  	Set objWorksheet2 = objWorkbook2.Worksheets("sustitucion")
  
  objExcel.Application.Visible = True
  objExcel.Application.DisplayAlerts = False

objExcel.Application.Wait(Now + TimeValue("00:00:10")) 
  'objExcel.Application.Wait(Now + TimeValue("00:00:05")) 
  
  objWorkbook1.Worksheets("Sheet3").Activate
  
  lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,4).End(-4162).Row
For i=2 To lastRow
If InStr(objWorksheet1.Cells(i,7).value,".") > 0 and InStr(objWorksheet1.Cells(i,7).value,".") <= 1 then
  
  objWorksheet1.Cells(i,3).value = "BC"
  
Else
  
    lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,4).End(-4162).Row
For i=2 To lastRow
If InStr(objWorksheet1.Cells(i,7).value,".") > 1 then
  
  objWorksheet1.Cells(i,3).value = "Drenado"
  
Else


	valor = objWorksheet1.Cells(i,4).value
  	plant = objWorksheet1.Cells(i,10).value
	objWorkbook2.Worksheets("sustitucion").Activate
	
  objExcel.Cells(1,4).value = plant
  


  For j = 1 To 16'objWorksheet2.UsedRange.Columns.Count
    For k = 4 To objWorksheet2.UsedRange.Rows.Count
      		If objExcel.Cells(k,j).value = valor then
      			valor2 = objExcel.Cells(k,j).value  -4163
      			columna = Split(objExcel.Cells(k,j).Address,"$")(1)
        

      
          		If  columna = "B" then
      				'columna2 = objExcel.Cells(k,3).value 
        				If  IsNumeric(objExcel.Cells(k,8).value) then
      						columna2 = objExcel.Cells(k,6).value
        				Else
          					If IsNumeric(objExcel.Cells(k,12).value) then
            					columna2 = objExcel.Cells(k,10).value
          					Else
                      			If IsNumeric(objExcel.Cells(k,16).value) then
            						columna2 = objExcel.Cells(k,14).value
          						Else
              						columna2 = "BO"
          						End If
          					End If
        				End If
    			End If
        If  columna = "F" then
      				'columna2 = objExcel.Cells(k,7).value
        				If  IsNumeric(objExcel.Cells(k,4).value) then
      						columna2 = objExcel.Cells(k,2).value
        				Else
          					If IsNumeric(objExcel.Cells(k,12).value) then
            					columna2 = objExcel.Cells(k,10).value
          					Else
                      			If IsNumeric(objExcel.Cells(k,16).value) then
            						columna2 = objExcel.Cells(k,14).value
          						Else
              						columna2 = "BO"
          						End If
          					End If
        				End If
    			End If
        If  columna = "J" then
      				'columna2 = objExcel.Cells(k,11).value
        				If  IsNumeric(objExcel.Cells(k,4).value) then
      						columna2 = objExcel.Cells(k,2).value
        				Else
          					If IsNumeric(objExcel.Cells(k,8).value) then
            					columna2 = objExcel.Cells(k,6).value
          					Else
                      			If IsNumeric(objExcel.Cells(k,16).value) then
            						columna2 = objExcel.Cells(k,14).value
          						Else
              						columna2 = "BO"
          						End If
          					End If
        				End If  
    			End If
      
        If  columna = "N" then
      				'columna2 = objExcel.Cells(k,15).value
        				If  IsNumeric(objExcel.Cells(k,4).value) then
      						columna2 = objExcel.Cells(k,2).value
        				Else
          					If IsNumeric(objExcel.Cells(k,8).value) then
            					columna2 = objExcel.Cells(k,6).value
          					Else
                      			If IsNumeric(objExcel.Cells(k,12).value) then
            						columna2 = objExcel.Cells(k,10).value
          						Else
              						columna2 = "BO"
          						End If
          					End If
        				End If
     			End If

    		End If
        
		Next 
	Next 

    		

  

  objWorkbook1.Worksheets("Sheet3").Activate
'objWorksheet1.Cells(i,3).value = valor2
  'objWorksheet1.Cells(i,1).value = plant
  objWorksheet1.Cells(i,3).value = columna2
If valor2 =  "" then
  objWorksheet1.Cells(i,3).value = "BO"
  'objWorksheet1.Cells(i,1).value = ""
  

End If
valor2 = ""
j=0
k=0

End If
Next 

objWorkbook1.Worksheets("Sheet3").Activate

       'Saves and closes the Excel file'
   objWorkbook2.Save
   objWorkbook2.Close SaveChanges = True

      'Saves and closes the Excel file'
   objWorkbook1.Save
   objWorkbook1.Close SaveChanges = True
      
       'Saves and closes the Excel file'
   'objWorkbook2.Save
   'objWorkbook2.Close SaveChanges = True

End Sub