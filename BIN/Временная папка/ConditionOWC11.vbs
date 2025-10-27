Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("SYSTEM")
If WshSysEnv.Item("PROCESSOR_ARCHITECTURE") = "x86" Then
	CPF = "%CommonProgramFiles%"
Else
	CPF = "%CommonProgramFiles(x86)%"
End If
If Not fso.FileExists(WshShell.ExpandEnvironmentStrings(CPF & "\microsoft shared\Web Components\11\OWC11.dll")) Then
	msiMessageTypeWarning = &H02000000 
	Set record = Session.Installer.CreateRecord(0)
	record.StringData(0) = "Не забудьте установить веб-компоненты Office 2003!"
	Session.Message msiMessageTypeWarning, record
'	WshShell.Run("http://www.microsoft.com/downloads/details.aspx?displaylang=ru&FamilyID=7287252c-402e-4f72-97a5-e0fd290d4b76")
End If

SFN = WshShell.ExpandEnvironmentStrings("%TEMP%" & "\ConditionOWC11.vbs")
If fso.FileExists(SFN) Then fso.DeleteFile(SFN)
