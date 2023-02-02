If Not WScript.Arguments.Named.Exists("elevate") Then

  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit

End If

Function pingaNimMim()

	Dim target
	Dim result

	target= "192.168.22.8"

	Set shell = WScript.CreateObject("WScript.Shell")
	Set shellexec = shell.Exec("ping " & target) 

	result = LCase(shellexec.StdOut.ReadAll)

	If InStr(result , "resposta de") Then
	  pingaNimMim = 1
	Else
	  pingaNimMim = 0
	End If
	
End Function

Function bIsFileDownloaded(strPath, timeout)

  Dim FSO, fileIsDownloaded
  set FSO = CreateObject("Scripting.FileSystemObject")
  fileIsDownloaded = false
  limit = DateAdd("s", timeout, Now)
  
  Do While Now < limit
  
	If FSO.FileExists(strPath) Then : fileIsDownloaded = True : Exit Do : End If
    WScript.Sleep 1000      
  
  Loop
  
  Set FSO = Nothing
  bIsFileDownloaded = fileIsDownloaded

End Function

Function deletaScript(path)	

	Const DeleteReadOnly = TRUE
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(path), DeleteReadOnly

End Function

Function reinicia()
	
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 18"

end Function

Dim ping
ping = 0

Do While ping = 0
	
	If pingaNimMim() = 1 Then		
		
		If bIsFileDownloaded("C:\Suporte\Part1.vbs", 3) Then
			
			Set WshShell = WScript.CreateObject("WScript.Shell")	
			WshShell.Run "C:\Suporte\Part1.vbs"
			WScript.Sleep 3000	
			deletaScript("C:\Suporte\Part1.vbs")	
			reinicia()
			
		ElseIf bIsFileDownloaded("C:\Suporte\Part2.vbs", 3) Then
			
			Set WshShell = WScript.CreateObject("WScript.Shell")
			WshShell.Run "C:\Suporte\Part2.vbs"
			WScript.Sleep 10000
			deletaScript("C:\Suporte\Part2.vbs")
			WshShell.Run "C:\Suporte\Dominio.vbs"
			WScript.Sleep 6000
			deletaScript("C:\Suporte\Dominio.vbs")	
			reinicia()		
		  
		Else

			Set WshShell = WScript.CreateObject("WScript.Shell")
			WshShell.Run "winword"
			WScript.Sleep 3000
			WshShell.Run "C:\Suporte\epskit_x64.exe"
			WScript.Sleep 3000
			deletaScript("C:\Users\suporte\AppData\Roaming\Microsoft\Windows\STARTM~1\Programs\Startup\chamaScript.bat")	
			WScript.Sleep 1000	
			
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			strScript = Wscript.ScriptFullName
			objFSO.DeleteFile(strScript)
			
		End If			
		
		ping = 1
		
	Else
	
		WScript.Sleep 15000		
	
	End If

Loop
