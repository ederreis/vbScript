'				MAIN FUNCTION.
'	--> Ping multiple computers and log when one or more doesn't respond.
'
'	Logics:
'		
'		Configurations to setup
'		Check Scenario
'		If not running via cscript.exe Then running via cscript.exe
'
'		Do While True
'			If not ping to address then record on log time and address.
'			Wait
'		Loop
'	
'	Function:
'
'		checkLogPath -> boolean (true for folder exists || false if does not exists)
'		Ping -> boolean (true for ping valid || false for ping not valid)
'
'	Subrutine:
'
'		logme -> record the log
'		checkLogFile -> create a file if does not exists.
'

'	################### Configuration #######################
'
'	Enter the IPs or machine names on the line below separated by a colon
strMachines = array("www.yahoo.com","www.google.com","www.microsoft.com")
'
' 	Put the path on: strLogPath
' 	Puth the file name on: strLogFile
strLogPath = "c:\logs\"
strLogFile = "pinglog.txt"
'	################### End Configuration ###################

strLogAddress = strLogPath & strLogFile	'putting all together

' 	Check scenario, folder exists? files exists if not create.
if not checkLogPath(strLogPath) then WScript.Echo "You must create a folder called: " & strLogPath : Wscript.Quit
checkLogFile(strLogAddress)

'	The default application for .vbs is wscript. If you double-click on the script,
'	this little routine will capture it, and run it in a command shell with cscript.

If Right(WScript.FullName,Len(WScript.FullName) - Len(WScript.Path)) <> "\cscript.exe" Then
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objStartup = objWMIService.Get("Win32_ProcessStartup")
	Set objConfig = objStartup.SpawnInstance_
	Set objProcess = GetObject("winmgmts:root\cimv2:Win32_Process")

	objProcess.Create WScript.Path + "\cscript.exe """ + WScript.ScriptFullName + """", Null, objConfig, intProcessID

	WScript.Quit

End If

Const ForAppending = 8

Do While True 
	
	For Each machine In strMachines
		
		If not Ping(machine) Then
			
			Call logme(Time & " - "  & machine & " is not responding to ping",strLogAddress)
			
		Else
			
			WScript.Echo(Time & " + "  & machine & " is responding to ping")
			
		End If
		
	Next

	WScript.Sleep 5000

Loop



Sub logme(message,strLogAddress)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Set objTextFile = objFSO.OpenTextFile(strLogAddress, ForAppending, True)

	objtextfile.WriteLine(message)

	WScript.Echo(message)

	objTextFile.Close

End Sub

Function checkLogPath(strLogPath)
	
	Dim objFSO
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If not objFSO.FolderExists(strLogPath) Then
		
		checkLogPath = False
		
	Else
		
		checkLogPath = True
		
	End If

End Function

Sub checkLogFile(strLogAddress)
	
	Dim objFSO
	Dim objFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists(strLogAddress) Then
	
		Set objFile = objFSO.GetFile(strLogAddress)

	Else
		
		Set objFile = objFSO.CreateTextFile(strLogAddress, True)
		
	End If
	
End Sub

Function Ping(machine)
	
	Dim colPingResults
	Dim objPingResult
	Dim objWMIService
	
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set ColPingResults = objWMIService.ExecQuery("Select StatusCode from Win32_PingStatus where Address = '" & machine & "'")
	
	For Each objPingResult In colPingResults
		
		If objPingResult.StatusCode = 0 Then
			
			Ping = True
		
		Else
			
			Ping = False
		
		End If
		
		Exit For
	
	Next
	
End Function
