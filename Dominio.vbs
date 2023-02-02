Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

Function getTxt(line)

	strNomeArquivo = getPath() + ".txt"

	Dim fso
	Set fso = CreateObject("Scripting.Filesystemobject")

	If fso.FileExists(strNomeArquivo) Then
	  Set obj = fso.OpenTextFile(strNomeArquivo,1,true)
	  
	  contLinha = 0
	  Do While obj.AtEndOfStream = False
		contLinha = contLinha + 1
		linha = obj.ReadLine
		   
		If contlinha = line then
		
			getTxt = linha			
			Exit Function
		
		End If  
			
	   Loop
	   
	Else

		MsgBox("Arquivo não encontrado!")

	End If

End Function

Function getRange()

	dim NIC1, Nic, StrIP, CompName

	Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

	For Each Nic in NIC1

	if Nic.IPEnabled then

	StrIP = Nic.IPAddress(i)

	Set WshNetwork = WScript.CreateObject("WScript.Network")

	CompName = WshNetwork.Computername

	Dim WMI, Configs, Config, Adapters, Adapter

	Set WMI = GetObject("winmgmts:{impersonationlevel=impersonate}root/cimv2")

	' BEGIN CALLOUT A
	Set Configs = WMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")
	' END CALLOUT A

	For Each Config In Configs
	' BEGIN CALLOUT B
	  Set Adapters = WMI.AssociatorsOf("Win32_NetworkAdapterConfiguration.Index=" & Config.Index, "Win32_NetworkAdapterSetting")
	' END CALLOUT B
	  For Each Adapter In Adapters
		'If Left(Adapter.Description, 14) = "Cisco AnyConnect VPN Virtual Miniport Adapter for Windows" Then
		  VPNIP = Config.IPAddress(0)

		'End If
	  Next
	Next	
	
	getRange = Left(StrIP, 11)
	
	Exit Function
	
	wscript.quit

	end if

	next	

End Function

Function getPath()
	
	Dim range
	range = getRange()

	Select case range
	
	'ADM Ponta Porã
	case "192.168.191."
	    getPath = "\\192.168.191.5\HostNames\PON"
	'ADM DR\HUB 
	case "192.168.192."
	    getPath = "\\192.168.192.5\HostNames\CGR"
	'ADM Dourados 
	case "192.168.193."
	    getPath = "\\192.168.193.5\HostNames\DOU"
	'ADM Corumbá
	case "192.168.194."
		getPath = "\\192.168.194.5\HostNames\COR"
	'ADM Três Lagoas
	case "192.168.195."
		getPath = "\\192.168.195.5\HostNames\TLG"
	'PED Ponta Porã
	case "192.168.21."
	    getPath = "\\192.168.22.8\HostNames\PON\Settings"
	'PED HUB 
	case "192.168.28."
	    getPath = "\\192.168.22.8\HostNames\HUB\Settings"
	'PED Dourados 
	case "192.168.23."
	    getPath = "\\192.168.22.8\HostNames\DOU\Settings"
	'PED Corumbá
	case "192.168.24."
		getPath = "\\192.168.22.8\HostNames\COR\Settings"
	'PED Três Lagoas
	case "192.168.25."
		getPath = "\\192.168.22.8\HostNames\TLG\Settings"
	'PED Turismo Gastronomia
	case "192.168.27."
		getPath = "\\192.168.22.8\HostNames\TLG\Settings"
	case else
	   
	   MsgBox "Range de IP denconhecido"
	
	End select
	
End Function

Function insereDominio()

on error resume next

'/////////////////////////////////////////////////////////////////////////////////////
'/                                    Insere no Domínio                /
'/////////////////////////////////////////////////////////////////////////////////////

Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DELETE = 4
Const WIN9X_UPGRADE = 16
Const DOMAIN_JOIN_IF_JOINED = 32
Const JOIN_UNSECURE = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET = 256
Const INSTALL_INVOCATION = 262144
 
strDomain = getTxt(2)
strPassword = getTxt(5)
strUser = getTxt(8)
 
Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName
 
Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
    strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & _
        strComputer & "'")
 
ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, _
    strPassword, strDomain & "\" & strUser, NULL, _
        JOIN_DOMAIN + ACCT_CREATE)

end function

insereDominio()