Dim oFso,oShell,sLine,sLogName,sLogDir,sMasterLog,oMasterLog

Set oFso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

Sub StartService(strServiceName)
	'strServiceName = "Alerter"
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
	For Each objService in colListOfServices
	objService.StartService()
	Log "Executed Start of service: "&strServiceName
	Next
End Sub

Sub Log(sLine)
	On Error Resume Next
	If Session.Property("ProductCode") = "" Then
		sLogDir = "" & sLogDir
		sLogName = "" & sLogName
		If sLogDir = "" Then sLogDir = oFso.GetParentFolderName(WScript.ScriptFullName)
		If Not oFso.FolderExists(sLogDir) Then
			oShell.Run "cmd.exe /c MD """ & sLogDir & """", 0, True
			If Not oFso.FolderExists(sLogDir) Then
				sLogDir = oShell.ExpandEnvironmentStrings("%TEMP%")
			End If
		End If
		If sLogName = "" Then sLogName = Left(WScript.ScriptName, Len(WScript.ScriptName) - 4) & "_Master_" & Replace(Replace(Replace(Now, " ", "_"), ":", "."), "/", ".") & ".log"
		sMasterLog = sLogDir & "\" & sLogName
		Set oMasterLog = oFso.OpenTextFile(sMasterLog, 8, True)
		Err.Clear
		oMasterLog.WriteLine Now & " : " & sLine
		oMasterLog.Close
		If Err.Number <> 0 Then WScript.Quit
		oMasterLog.Close
		Set oMasterLog = Nothing
	Else
		Dim oRec
		Set oRec = Session.Installer.CreateRecord(1)
		oRec.StringData(1) = Now & " : " & sLine
		Session.Message &H04000000, oRec
	End If
End Sub

StartService "EWP"