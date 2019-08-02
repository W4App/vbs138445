'Write by W.b.an @2019.08.01
Option Explicit
'Disable Tcpipnetbios!
Sub closeTcpipnetbios()
    Dim objWMI
    Set objWMI =GetObject("winmgmts:\\" & "." & "\root\cimv2")
    Dim Queries 
    Set Queries  = objWMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    Dim qu
    For Each qu In Queries
    qu.setTcpipnetbios(2)
    Next
    WScript.Sleep 1000
    Set Queries  = objWMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    For Each qu In Queries
        If qu.TcpipNetbiosOptions=2 Then
        WScript.Echo "TcpipNetbios Disabled!"
        End If 	
    Next
End Sub
'operate service
sub closeService(name)
	Dim objWMIService
	Dim colServiceList
	Dim obj
	Dim errReturn
	Dim strPath
	Dim Flag
	strPath="Associators of " _
    & "{Win32_Service.Name='"&name&"'} Where " _
        & "AssocClass=Win32_DependentService " & "Role=Antecedent"        	
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")   
    Set colServiceList = objWMIService.ExecQuery(strpath)    
	For Each obj in colServiceList
		If obj.state="Stopped" Then
			WScript.Echo obj.name&" is Stopped!"
			Else
			Flag=	obj.StopService()
			WScript.Sleep 1000
			Flag=	obj.Change( , , , , "Disabled") 	
		End If
	Next
	Wscript.Sleep 10000
	Set colServiceList = objWMIService.ExecQuery _
	        ("Select * from Win32_Service where Name='"&name&"'")
	For Each obj in colServiceList
		If obj.state="Stopped" Then
			WScript.Echo obj.name&" is Stopped!"
			Else
			Flag =obj.StopService()					
			WScript.Sleep 1000
			Flag =obj.Change( , , , , "Disabled") 		
		End If
Next 
'check service status
Wscript.Sleep 5000
Set colServiceList = objWMIService.ExecQuery _
	        ("Select * from Win32_Service where Name='"&name&"'")
	For Each obj in colServiceList
		If obj.state="Stopped" Then
			WScript.Echo obj.name&"  service Stopping! --- ok"
		End If
		If obj.StartMode ="Disabled" Then
			WScript.Echo obj.name&": Disable Boot!  --- ok"					
		End If  
	Next	  
end Sub
'register operation
Sub modiRigister()
	Const HKEY_CLASSES_ROOT =    &H80000000 
	Const HKEY_CURRENT_USER =    &H80000001
	Const HKEY_LOCAL_MACHINE =   &H80000002
	Const HKEY_USERS =           &H80000003
	Const HKEY_CURRENT_CONFIG =  &H80000005  
	Const REG_SZ = 1
	Const REG_EXPAND_SZ = 2
	Const REG_BINARY = 3
	Const REG_DWORD = 4
	Const REG_MULTI_SZ = 7  
	dim oReg
	Dim result
	Dim strArr
	strArr =Array("")
	set oReg =GetObject("winmgmts:{impersonationLevel=impersonate}!\\"&"."&"\root\default:StdRegProv")
	'DCOM
	oReg.SetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Ole","EnableDCOM","N"	
	'DCOM Protocols
	oReg.SetMultiStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Rpc","DCOM Protocols",strArr		
	'new key
	oReg.CreateKey HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Rpc\Internet"
	'new name/value
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\services\NetBT\Parameters","SMBDeviceEnabled",0	
End sub
WScript.Echo "modify register..."&vblf
modiRigister
WScript.Echo "modify TcpipNetbiosOptions..."&vblf
closeTcpipnetbios
'name =MSDTC
WScript.Echo "Stopping Distributed Transaction Coordinator and so on..."&vblf
closeService "MSDTC"
'name =LanmanServer
WScript.Echo "Stopping network share(file and printer)"&vblf
closeService "LanmanServer"
WScript.Echo "Done!, Reboot your compter, run 'netstat -an' check port status."&vblf
WScript.Sleep 10000

