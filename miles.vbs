Function miles(Parameters)
  
	arr = split(Parameters,"#")
	sRet = arr(0)
	sTasa  = arr(1)
	
	input = Round(sRet/sTasa,2)
	
	Dim output, regex1, regex2
	
	Set regex1 = New RegExp
	Set regex2 = New RegExp
	regex1.Pattern = "(\d{3})"
	regex1.Global = True
	regex2.Pattern = ",$"
	output = regex1.Replace(StrReverse(input), "$1,")
	miles = StrReverse(regex2.Replace(output, ""))
	
	If InStr(miles,"-,") <> 0 Then
		miles = Replace(miles, "-,", "-")
	End If

End Function

Parameters="-3819.12#36.50990"
MsgBox miles(Parameters)