' ������ ��� ����������� ������ �����������.��c � ����.��c
' ��� ������ ����� ����� 

Dim gTimerIns		' ������ �������� ����� ������� ������ OnInsert

' ��������� ������� �� ����� Intellisence � ������� Telepat_OnShowMemberList(Line, Col)
Sub ShowTooltip()
	Set Intellisence = Scripts("Intellisence")
	Intellisence.ShowTooltip(0)
End Sub

Function DelComment(line)
	DelComment = line
	PosKomment = InStr(1,line,"//")
	if PosKomment>0 Then
		if PosKomment = 1 Then
			DelComment = ""
		else
			DelComment	= Mid(line,1,PosKomment-1)
		End If
	End If
	DelComment = Trim(DelComment)
End Function

 '����� �� ������ �������� � ������������ 
Function VerifyIfInnerComment(Line, CommaPos)
	VerifyIfInnerComment = false

	PosKomment = InStr(1,line,"//")
	if PosKomment = 0 Then 
		Exit Function
	end if
	
	if PosKomment < CommaPos Then
		VerifyIfInnerComment = true
	End If                   
	
End Function ' VerifyIfInnerComment
	
'[+]metaeditor 01.05.2006, ������ �� Intellisence.vbs
Sub Telepat_OnInsert(InsertType, InsertName, Text)
	Select Case InsertType
		Case 10
			If InsertName = "�������������" Then 
				Text = "�������������(""!"");"
				gTimerIns = CfgTimer.SetTimer(1, True)
			end if
	End Select
End Sub
Sub Configurator_OnTimer(timerID)
	If timerID = gTimerIns Then
		CfgTimer.KillTimer gTimerIns
		Scripts("Intellisence").MethodsList
	End If
End Sub
'[+]_

function Telepat_OnShowMemberList(Line, Col)
	Telepat_OnShowMemberList=""
	ShowMethodList
End Function

' ������ ��� ������������� ������ Intell + Dots � "���������" ������ (Ctrl+I)
Sub ShowMethodList()
		
	set doc = CommonScripts.GetTextDocIfOpened(0)
	
	if doc is Nothing then 		Exit Sub
	If doc.LineCount = 0 Then	Exit Sub
			
	If (doc.SelStartLine<>0) And (doc.SelStartCol>0) Then
		
		if doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelEndLine,doc.SelEndCol) = "." then	
			StrBezKomment = DelComment(doc.Range(doc.SelStartLine))
			if Len(StrBezKomment)=0 Then
				ShowTooltip()
				Exit Sub
			End IF               

			 '����� �� ������ �������� � ������������ 
			if VerifyIfInnerComment(doc.Range(doc.SelStartLine), doc.SelStartCol-1) then
				Set wshShell = CreateObject("wScript.Shell")
				wshShell.sendKeys "{ESC}"
				Exit Sub
			end if
				
			'�� �����, �������� ��������� �����������
			'�� ����� - ������� �������� ������ ������� Intellisence ������ ���� ����� ����� ������ ���...
			'�� ����� if Len( Trim(doc.Range(doc.SelStartLine,doc.SelStartCol))) = 0 Then
			Set Intellisence = Scripts("Intellisence")
			Ret = Intellisence.MethodsList()
			Select Case Ret 
				Case -1: ' ������ ESC, ������������ ���������, ������ � Dots �������� ������ ���
					exit Sub
				Case  1: ' ��� ��, ����� �� ������ �������������� ����������
					'message "intel called"
					ShowTooltip()
					exit Sub
				Case Else: ' ����� �� ���� ��������� ������ ��� ��������������, �������� �������� Dots'�
			End Select
			'�� ����� End If
			if Doc.Name <> "���������� ������" Then
				SuccessfulWork = Scripts("dots").IsSuccessfulWork()		
				if SuccessfulWork Then
					'message "dots called"
					ShowTooltip()
					exit Sub
				End If
			End If
		else
			exit Sub
		End If
	End If
End Sub

Private Function ShouldHandleTelepatOnShowMemberListEvent()
	ShouldHandleTelepatOnShowMemberListEvent = false
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	intell_ini = fso.BuildPath(BinDir, "config\Intell\intell.ini")
	
	If Not fso.FileExists(intell_ini) Then Exit Function
		
	Set ini = CreateObject("OpenConf.RegistryIniFile")
	ini.SetConfig(Configurator)
	ini.UsedIniFile = true
	ini.IniFile = intell_ini
	
	If ini.Param(Null,"TELEPAT") = "��" Then
		ShouldHandleTelepatOnShowMemberListEvent = true
	End If
End Function

'
' ��������� ������������� �������
'
Private Sub Init()
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
        Message "������ " & SelfScript.Name & " �� ��������", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	  
	' ��������: ��������� ������ ����������� � ����� �������, ����� �� ���������� !

	If ShouldHandleTelepatOnShowMemberListEvent() Then
		c.AddPluginToScript SelfScript, "�������", "Telepat", Telepat
	End If	

	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

Init
