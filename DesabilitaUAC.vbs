If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

' Desabilitando o UAC:
' -------------------

Const HKEY_LOCAL_MACHINE = &H80000002

Set objShell = WScript.CreateObject("WScript.Shell")
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","EnableLUA",sEnableLUA

If sEnableLUA  = 1 Then	
	objShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 0, "REG_DWORD"
Else
	WScript.Quit
End If

Set objShell = Nothing

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScript = Wscript.ScriptFullName
objFSO.DeleteFile(strScript)