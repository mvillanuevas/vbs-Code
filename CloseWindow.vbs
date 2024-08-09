Sub CloseWindow'(sWindow)
	sWindow = "EXPORT.XLSX - Excel"
	Set oShell = CreateObject("WScript.Shell") 
	oShell.AppActivate sWindow
	oShell.SendKeys "%{F4}"
	oShell.Quit
End Sub
Call CloseWindow

' 500000000
' 500000001


' Spring@2024

'1400001421  1400001369