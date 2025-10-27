'===========================================================================
Sub GotoBeginOfMethod()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	ModuleText = split(doc.Text, vbCrLf)
	for i = doc.SelStartLine to 0 step -1
		sText = UCase(lTrim(ModuleText(i)))
		if Instr(sText,"���������") = 1 or Instr(sText,"�������") = 1 _
		or Instr(sText,"SUB") = 1 or Instr(sText,"FUNCTION") =1 or Instr(sText,"PRIVATE") = 1 then
			doc.MoveCaret i, 0
			Exit For
		end if
	next
End Sub ' GotoBeginOfMethod

'===========================================================================
Sub GotoEndOfMethod()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	ModuleText = split(doc.Text, vbCrLf)
	for i = doc.SelStartLine to UBound(ModuleText)
		sText = UCase(lTrim(ModuleText(i)))
		if Instr(sText,"��������������") = 1 or Instr(sText,"������������") = 1 _
		or Instr(sText,"END SUB") = 1 or Instr(sText,"END FUNCTION") =1 then
			doc.MoveCaret i, 0
			Exit For
		end if
	next
End Sub ' GotoEndOfMethod                                                   

'=========================================================================
Sub SelectProcedure()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  
  GotoBeginOfMethod()
  l1 = Doc.SelStartLine
  GotoEndOfMethod()
  l2 = Doc.SelStartLine
  
  Doc.MoveCaret l1, 0, l2+1, 0
End Sub

'========================================================================================
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
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub
'========================================================================================
Init


