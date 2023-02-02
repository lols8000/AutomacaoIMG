If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

Function AssetTag()

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colSMBIOS = objWMIService.ExecQuery _
	("Select * from Win32_SystemEnclosure")

	For Each objSMBIOS in colSMBIOS
	
	AssetTag = objSMBIOS.SMBIOSAssetTag
	
	Next

End Function

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

Function pingaNimMim()

	Dim target
	Dim result

	target = getRange() + "5"

	Set shell = WScript.CreateObject("WScript.Shell")
	Set shellexec = shell.Exec("ping " & target) 

	result = LCase(shellexec.StdOut.ReadAll)

	If InStr(result , "resposta de") Then
	  pingaNimMim = 1
	Else
	  pingaNimMim = 0
	End If
	
End Function

Function NewName(strNewName)

	StrComputer = "."
	Dim flag
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
			For Each objComputer in colComputers
				 objComputer.Rename(UCASE(strNewName))
			Next

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
	    getPath = "\\192.168.22.8\HostNames\PON\PON"
	'PED HUB 
	case "192.168.28."
	    getPath = "\\192.168.22.8\HostNames\HUB\HUB"
	'PED Dourados 
	case "192.168.23."
	    getPath = "\\192.168.22.8\HostNames\DOU\DOU"
	'PED Corumbá
	case "192.168.24."
		getPath = "\\192.168.22.8\HostNames\COR\COR"
	'PED Três Lagoas
	case "192.168.25."
		getPath = "\\1192.168.22.8\HostNames\TLG\TLG"
	'PED Turismo e Gastronomia
	case "192.168.27."
		getPath = "\\192.168.22.8\HostNames\ETG\ETG"
	case else
	   
	   MsgBox "Range de IP denconhecido"
	
	End select
	
End Function

Function compName()
	
	If pingaNimMim() = 0 then
		
		if msgbox ("Computador Sem Rede, por favor insira o nome manualmente.",vbyesno + vbquestion,"Alerta o computador será reiniciado em 10 segundos!")=vbyes Then
			
			strNewName = InputBox("Insita o nome do Computador","Nome ")		
			NewName(strNewName)
			Set WSHShell = WScript.CreateObject("WScript.Shell")
			WshShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 10"

		else

			Exit Function

		end if	
	
	Else
	
		StrFile = getPath() + ".xlsx"
		strCompName = "DESKTOP-" + AssetTag()
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open _
			(strFile,1)
		intRow = 2
		Do Until objExcel.Cells(intRow,1).Value = ""
			
			If objExcel.Cells(intRow,1).Value = UCASE(strCompname) then
				strNewName = objExcel.Cells(intRow,2)
				NewName(strNewName)

				objExcel.Cells(intRow,3).Value = "OK" 
				objExcel.Cells(intRow,4).Value = Date()			 
				flag = 1

			End If
				
			intRow = intRow + 1

		Loop

		If flag = "" Then
		
			strNewName = getTxt(1) + "-" + AssetTag()
			objExcel.Cells(intRow, 1).Value = "DESKTOP"+ "-" + AssetTag()
			objExcel.Cells(intRow, 2).Value = getTxt(1) + "-" + AssetTag()
			objExcel.Cells(intRow, 3).Value = "OK" 
			objExcel.Cells(intRow, 4).Value = Date()
			NewName(strNewName)
		
		End If
	
	End if

	Set ActiveWorkbook = objExcel.ActiveWorkbook 
	objExcel.Application.DisplayAlerts = False
	objExcel.ActiveWorkbook.SaveAs strFile
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit

End Function

compName()