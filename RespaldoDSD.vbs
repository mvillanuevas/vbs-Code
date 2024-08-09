Sub macro(Parameters)
  
  Dim arr,strZEP1CompletePath,strBaseCompletePath
  arr=split(Parameters,"#")
  strZEP1CompletePath=arr(0)
  strBaseCompletePath=arr(1)

  
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkbook1 = objExcel.Workbooks.Open(strZEP1CompletePath)
  Set objWorksheet1 = objWorkbook1.Worksheets("Sheet1")	
  Set objWorkbook2 = objExcel.Workbooks.Open(strBaseCompletePath)
  Set objWorksheet2 = objWorkbook2.Worksheets("base")
  
  objExcel.Application.Visible = True
	objExcel.Application.DisplayAlerts = False
  	objExcel.Application.ScreenUpdating = False

  objExcel.Application.Wait(Now + TimeValue("00:00:10")) 
  
  objWorkbook2.Worksheets("base").Activate
  
    lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,3).End(-4162).Row
objExcel.Range("C2:AH" &  lastRow).ClearContents
  
  objWorkbook1.Worksheets("Sheet1").Activate
  
  lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,1).End(-4162).Row
objExcel.Range("A1:AD" &  lastRow).Copy
objWorkbook2.Worksheets("base").Activate

objWorkbook2.Worksheets("base").Activate

objExcel.Range("C1").PasteSpecial
objExcel.CutCopyMode = False
objExcel.Range("A1").Select

'objExcel.Application.Wait(Now + TimeValue("00:00:10"))
objWorkbook2.Worksheets("base").Activate
lastRowConcat = objWorksheet2.Cells(objWorksheet2.Rows.Count,3).End(-4162).Row + 1
objExcel.Range("AG1").value =  "concatenado"
objExcel.Range("AH1").value = "Falso"

'Formulas in column AG
objExcel.Range("AG2").FormulaR1C1 = "=RC[-30]&RC[-26]&RC[-25]&RC[-23]"
objExcel.Range("AG2").AutoFill(objWorkbook2.Worksheets("base").Range("AG2:AG" & lastRowConcat))

'Formula in column AH
objExcel.Range("AH2").FormulaR1C1 = "=RC[-1]=R[1]C[-1]"
objExcel.Range("AH2").AutoFill(objWorkbook2.Worksheets("base").Range("AH2:AH" & lastRowConcat))

'objExcel.Run "base"

'lastRow = objWorksheet2.Cells(objWorksheet2.Rows.Count,3).End(-4162).Row + 1
'  lastRow2 = objWorksheet2.Cells(objWorksheet2.Rows.Count,34).End(-4162).Row
'objWorkbook2.Worksheets("base").Range("AH" & lastRow & ":AH" & lastRow2).Delete


lastRow = objWorksheet2.Cells(objWorksheet2.Rows.Count,3).End(-4162).Row
objWorkbook2.Worksheets("base").Range("C1:AH" & lastRow).Select

'objExcel.Cells(lastRow,lastColumn).Copy
   'Saves and closes the Excel file'
   objWorkbook1.Save
   objWorkbook1.Close SaveChanges = True
  
   'Saves and closes the Excel file'
   objWorkbook2.Save
   objWorkbook2.Close SaveChanges = True
      

End Sub

Sub macro(Parameters)
  
  Dim arr,strZEP1CompletePath,strBaseCompletePath
  arr=split(Parameters,"#")
  strZEP1CompletePath=arr(0)
  strBaseCompletePath=arr(1)

  
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

  Set objWorkbook1 = objExcel.Workbooks.Open(strBaseCompletePath)
  Set objWorksheet1 = objWorkbook1.Worksheets("base")
  Set objWorksheet2 = objWorkbook1.Worksheets("limpia")
  
  objExcel.Application.Visible = True
  objExcel.Application.DisplayAlerts = False
  objExcel.Application.ScreenUpdating = False

objExcel.Application.Wait(Now + TimeValue("00:00:10")) 
  'objExcel.Application.Wait(Now + TimeValue("00:00:05")) 
  objWorksheet2.Cells.Delete 
  objWorkbook1.Worksheets("base").Activate
  
    lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,3).End(-4162).Row
objExcel.Range("AH1").Select
objExcel.Range("AH1").Autofilter 33, Array("False"), 7

objExcel.Range("C1:AF" & lastRow).Copy objWorksheet2.Range("A1")

objWorksheet2.Activate
lastRow = objWorksheet2.Cells(objWorksheet2.Rows.Count,1).End(-4162).Row

objExcel.Range("A1:AD" & lastRow).Select

  
'objExcel.Cells(lastRow,lastColumn).Copy
   'Saves and closes the Excel file'
   objWorkbook1.Save
   objWorkbook1.Close SaveChanges = True

      

End Sub

Sub macro(Parameters)
  
  Dim arr,strZEP1CompletePath,strBaseCompletePath
  arr=split(Parameters,"#")
  strZEP1CompletePath=arr(0)
  strBaseCompletePath=arr(1)

  
  'Opens the Excel file'
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkbook1 = objExcel.Workbooks.Open(strZEP1CompletePath)
  Set objWorksheet1 = objWorkbook1.Worksheets("Sheet2")	

  Set objWorkbook2 = objExcel.Workbooks.Open(strBaseCompletePath)
  Set objWorksheet2 = objWorkbook2.Worksheets("limpia")
  
  objExcel.Application.Visible = True
  objExcel.Application.DisplayAlerts = False
  objExcel.Application.ScreenUpdating = False

objExcel.Application.Wait(Now + TimeValue("00:00:10")) 
  'objExcel.Application.Wait(Now + TimeValue("00:00:05")) 
  

  objWorksheet2.Activate
  
    lastRow = objWorksheet2.Cells(objWorksheet2.Rows.Count,1).End(-4162).Row


objExcel.Range("A1:AD" & lastRow).Copy objWorksheet1.Range("A1")

   'Saves and closes the Excel file'
   objWorkbook2.Save
   objWorkbook2.Close SaveChanges = True

objWorksheet1.Activate
objExcel.Columns("J:J").Insert
objExcel.Range("J1").FormulaR1C1 = "D"

'Autofit en la columna J
 lastRow = objWorksheet1.Cells(objWorksheet1.Rows.Count,1).End(-4162).Row
objExcel.Range("J2").FormulaR1C1 = "=RC[-2]-RC[-1]"
objExcel.Range("J2").AutoFill(objWorksheet1.Range("J2:J" & lastRow))

objExcel.Sheets("Sheet2").Select
    objExcel.Sheets.Add
    objExcel.Sheets("Sheet2").Select
objWorksheet1.Activate
    objExcel.Range("A1:AE1").AutoFilter
    objExcel.Range("J1").AutoFilter 10, Array("<>0")

objWorksheet1.Activate
objExcel.Range("A1:AE" & lastRow).Copy objWorkbook1.Worksheets("Sheet3").Range("A1")
       
objWorkbook1.Worksheets("Sheet3").Activate
objExcel.Range("B:D").Delete
objExcel.Columns("C:C").Insert
objExcel.Range("C1").FormulaR1C1 = "CDS"



   'Saves and closes the Excel file'
   objWorkbook1.Save
   objWorkbook1.Close SaveChanges = True
  
      

End Sub