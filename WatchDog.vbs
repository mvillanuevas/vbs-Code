'Script que todo el tiempo se está ejecutando y espera la ventana de timeout de SAP para terminar el proceso'
 
On Error Resume Next
 
Set wshShell = CreateObject("WScript.Shell")
	
 
endWindow = False
 
Do While endWindow = False
 
	endWindow = wshShell.AppActivate("Config_DEV_DSD.xlsx - Excel")
    sapWindow = wshShell.AppActivate("Print")
              
 
	If sapWindow = True Then
				  
		wshShell.AppActivate "Print"
		wshShell.SendKeys "%{F4}"		
								 
	End if 

Loop