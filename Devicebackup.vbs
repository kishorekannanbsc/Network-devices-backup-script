#$language = "VBScript"
#$interface = "1.0"

Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("process")
Dim userPass
Dim projectPath

'--------------------------------------(User editable Section)------------------------------------'

'In case of password change, update below
userPass = "userpassword"

'In case of script path change, update below
projectPath = "C:\Users\userid\Desktop\Test"

'--------------------------------------------------------------------------------------------------'

Dim logPath
logPath = projectPath & "\" & "Logs"

Dim hostsFile
hostsFile = projectPath & "\" & "hosts.txt"
Dim commandFile
commandFile = projectPath & "\" & "commands.txt"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim LoginUserName 
Dim LoginUserPassword
Dim EnablePassword

Sub Main
	If ( Not fso.FolderExists(logPath) ) Then
		fso.CreateFolder(logPath) 
	End If
	If ( Not fso.FileExists(hostsFile) ) Or ( Not fso.FileExists(commandFile) ) Then
		Dim Host_File
		Set Host_File = fso.CreateTextFile(hostsFile, True)
		Host_File.WriteLine("")
		Host_File.close
		MsgBox "Device IP list missing in - " & hostsFile
		Dim Command_File
		Set Command_File = fso.CreateTextFile(commandFile, True)
		Command_File.WriteLine("")
		Command_File.close
		MsgBox "Device Commands missing in - " & commandFile
		Exit Sub
	End If
	
	Set fileIOSHosts = fso.OpenTextFile(hostsFile, 1)
		Do Until fileIOSHosts.AtEndOfStream
			host = fileIOSHosts.ReadLine
		If ( len(host) > 6 ) Then
			On Error Resume Next
			crt.session.connect(" /SSH2 /ACCEPTHOSTKEYS " & host)
			On Error Goto 0
			if not ( crt.session.connected ) then
				MsgBox host & " Connection failed... Click OK to continue next switch. (Empty log created)"
				crt.Session.LogFileName = logPath & "\" & host & "_Connection_failed_%D-%M-%Y_%hh%mm%ss.txt"
				crt.Session.Log True
				crt.Session.Log False
			else
				crt.Session.Log True
				crt.Screen.Send "ssh " & host & chr(13)
				crt.Screen.WaitForString ": "
				crt.Screen.Send userPass & chr(13)
				crt.Screen.WaitForString "#"
				crt.Screen.Send chr(13)
				Set objTab = crt.GetScriptTab()
				nRow = objTab.Screen.CurrentRow
				hostname = objTab.screen.Get(nRow, 0, nRow, objTab.Screen.CurrentColumn - 2)
				hostname = Trim(hostname)
				crt.Session.LogFileName = logPath & "\" & hostname & "_" & host &"_%Y-%M-%D--%h-%m-%s.%t.txt"
				Set fileDumpConf = fso.OpenTextFile(commandFile, 1)
				crt.Screen.Send "terminal length 0" & VbCr
				Do While fileDumpConf.AtEndOfStream <> True
					dump = fileDumpConf.ReadLine
					crt.Screen.Send (dump) & VbCr
					crt.Screen.WaitForString "#" 
					crt.Screen.Send chr(13)
				loop
				result = crt.Screen.WaitForStrings("# ", "$ ", 5)
				crt.Screen.Send "" & VbCr
				crt.Screen.WaitForString "$"
				crt.Session.Log False
				crt.Screen.Synchronous = False
			End If
		End If
	Loop
	fileIOSHosts.Close
	MsgBox "Script execution finished, check log here: " & logPath
End Sub
