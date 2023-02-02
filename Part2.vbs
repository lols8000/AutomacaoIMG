If Not WScript.Arguments.Named.Exists("elevate") Then

  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit

End If

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

Function strCompName()

	StrComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	strCompName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

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

' Muda adiciona senha ao usuário Suporte:
' ---------------------------------------

strComputer = "."
Set colAccounts = GetObject("WinNT://" & strComputer & ",computer")
Set objUser = GetObject("WinNT://" & strComputer & "/suporte, User")
objUser.SetPassword getTxt(11)
objUser.SetInfo

WScript.Sleep 3000

' Habilitando o UAC:
' ------------------

Const HKEY_LOCAL_MACHINE = &H80000002

Set objShell = WScript.CreateObject("WScript.Shell")
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","EnableLUA",sEnableLUA

If sEnableLUA  = 0 Then	
	objShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 1, "REG_DWORD"
Else
	WScript.Quit
End If

Set objShell = Nothing

WScript.Sleep 12000

'Adicionar atalho mapeamento de pasta professor
'----------------------------------------------

Function pastaProfessor(path)

	set WshShell = WScript.CreateObject("WScript.Shell")
	set oShellLink = WshShell.CreateShortcut("C:\Users\public\desktop" & "\Compartilhamento Professor.lnk")
	oShellLink.TargetPath = path
	oShellLink.WindowStyle = 1
	oShellLink.Description = "Compartilhamento Professor"
	oShellLink.WorkingDirectory = "%SystemDrive%\Programa"
	oShellLink.Save

End Function

Function CompName()

	dim strInput, intDash

	strInput = strCompName()
	intDash = InStr(strInput, "-")

	if intDash > 0 then
		CompName = Left(strInput, intDash - 1)
	end if
		
End Function

'Encontra uma palavra no arquivo TXT dentro de um intervalo definido e devolve a linha subsequente
Function EncontraPalavraTxt() 
	
	' Declare variables
	Dim objFSO, objFile, strLine, strWord, strNextLine, intStartLine, intEndLine, intCurrentLine

	' Set the word to search for
	strWord = "->" + CompName()

	' Set the start and end line of the range
	intStartLine = 13
	intEndLine = 33

	' Create a FileSystemObject
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' Open the text file
	Set objFile = objFSO.OpenTextFile(getPath() + ".txt", 1)

	' Initialize the current line
	intCurrentLine = 1

	' Read the file line by line
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If intCurrentLine >= intStartLine And intCurrentLine <= intEndLine Then
			If InStr(strLine, strWord) Then
				strNextLine = objFile.ReadLine
				If Left(strNextLine, 1) <> "*" Then									
					EncontraPalavraTxt = strNextLine
					Exit Do
				End If
			End If
		End If
		intCurrentLine = intCurrentLine + 1
	Loop

	' Close the file
	objFile.Close

	' Release the FileSystemObject
	Set objFSO = Nothing

End Function

pastaProfessor(EncontraPalavraTxt())
