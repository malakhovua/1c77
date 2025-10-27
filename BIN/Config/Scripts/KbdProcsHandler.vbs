$NAME KbdProcsHandler
'========================================================================================
'������ ���������� ������� "OnKeyPress" ������� "KbdProcs.dll"
'	MetaEditor <shotfire@inbox.ru> 
'========================================================================================
'========================================================================================
'09.09.2005
'	- ������ �����
'14.09.2005
'	- ��������� ��������� ���� ������ 
'	- ��������� ���������� ���������
'	- ����������� ����������, ��������, ������
'
'30.10.2005
'	- ������ ��������� �� ������� "�������������" � ���������, ������ �� ����������.
'
'06.02.2006
'	- �������� ����� �������������
'
'========================================================================================
'========================================================================================
'��������:
'=========
'������� OnKeyPress(ASCIIKeyCode, ByRef CancelKey, IsVirtual)
'  ���������� ��� ������� ������ ����������
'
'ASCIIKeyCode - ASCII ��� ������� �������
'
'CancelKey - ���� true �� ���������� ������ ������� �������
'
'IsVirtual - ������� ���� ��� ������� ����������� (�� ���������-��������),
'  ��������, ����� ������ "(" � ������� ���� ����� � ASCII ��� 40, 
'  �� ��� ������� ���� IsVirtual = true, � ��� "(" = false
'  
'===========================
'�������� � ������ �������:
'===========================
'GetKeyState(VirtualKeyCode) - ��������� ������� (������, ��������) 
'  ��������� ��. �������� API ������� GetKeyState
'
'===========================
'GetKeyboardLayout() - �������� ������� ��������� ����������
'  ��������� ��. �������� API ������� GetKeyboardLayout
'
'====================================
'����������� ���������� �� ��������:
'====================================
'Enabled = true/false - ���/���� ������
'
'========================================================================================
Const VK_SHIFT = 16
Const VK_CONTROL = 17
Const VK_MENU = 18 'ALT
Const VK_SCROLL = 145

Const klRU = 1049 '�������
Const klEN = 1033 '����������

Dim flDebug
Dim doc

'========================================================================================
Sub KbdProcs_OnKeyPress(ASCIIKeyCode, ByRef CancelKey, IsVirtual)
	Debug "Code: " & ASCIIKeyCode & "; Virtual: " & IsVirtual & "; CTRL: " & IsKeyPressed(VK_CONTROL) & "; ALT: " & IsKeyPressed(VK_MENU) & "; SHIFT: " & IsKeyPressed(VK_SHIFT)
	
	Select Case ASCIIKeyCode	
	Case 13: '����� Enter
		docType = GetDocType(doc)	
		Select Case docType
		Case -1:
			'���� ��� �������� ���� �� �� Shift+Enter ������� ���� ������������
			if IsKeyPressed(VK_SHIFT) then SendCommand cmdOpenConfigWnd
		
		Case docText: '��������� �������� 
			'���� ������ Shift+Enter � ������ �����������, �� ���� "//" ������������ �� ��������� ������
			'(��� � �������� - ������� ����� " �� ��������� ������ � ������ �������)
			
			if IsKeyPressed(VK_SHIFT) then
				pref = ""
				'set Matches = CommonScripts.RegExpExecute("^(\s*(//|')\s*)", getLeftPart(doc))
				Set regEx = New RegExp : regEx.Pattern = "^(\s*(//|')\s*)" : regEx.IgnoreCase = true : set Matches = regEx.Execute(getLeftPart(doc))	
				if not Matches is nothing then
					if Matches.Count > 0 then
						pref = Matches(0).SubMatches(0)
						CancelKey = true '������� ������� �������
						doc.Range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = vbCrLf & pref
						doc.MoveCaret doc.SelStartLine + 1, Len(pref)
					end if	
				end if
			end if
			
		Case docDEdit: '�����
			'�� Shift+Enter � ����� ������� ��� � ��������� �������� ����������� ��������
			if IsKeyPressed(VK_SHIFT) then
				if doc.Selection <> "" then
					if InStr(1,doc.Selection, ",") = 0 then 
						CancelKey = true
						ctrl = cInt(doc.Selection)
						strInf = doc.ctrlType(ctrl) & vbCrlf & _
								 "���������: " & doc.ctrlProp(ctrl,cpTitle) & vbCrlf & _
								 "�������������: " & doc.ctrlProp(ctrl,cpStrID) & vbCrlf & _
								 "�������: " & doc.ctrlProp(ctrl,cpFormul)
						message strInf
					end if
				end if
			end if			 
		End Select
	'Case 27:
		
	'Case 32: '������
	
	Case 221: '���������� "��" �� ������� �������
		docType = GetDocType(doc)	
		if docType = docText then
			if Right(getLeftPart(doc),1) = "�" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = """"""
				doc.MoveCaret doc.SelStartLine,doc.SelStartCol,doc.SelEndLine, doc.SelEndCol
			end if	 
		end if  
		
	Case 253: '���������� "��" �� ��������� ������� (��� ���)
		docType = GetDocType(doc)	
		if docType = docText then
			if Right(getLeftPart(doc),1) = "�" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = "''"
				doc.MoveCaret doc.SelStartLine,doc.SelStartCol,doc.SelEndLine, doc.SelEndCol

			end if	 
		end if  
		
	Case 222: '���������� " ��" �� "<>"
		docType = GetDocType(doc)	
		if docType = docText then
		    leftPart = getLeftPart(doc)
			if Right(leftPart,2) = " �" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = "<> "
				doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 2
			elseif Right(leftPart,1) = "�" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = "<>"
				doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
			end if	 
		end if
		
	Case 218: '���������� "��" �� "[]"
		docType = GetDocType(doc)	
		if docType = docText then
			if Right(getLeftPart(doc),1) = "�" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = "[]"
				doc.MoveCaret doc.SelStartLine,doc.SelStartCol,doc.SelEndLine, doc.SelEndCol
			end if	 
		end if

	'Case 59: '���������� " ;" �� $
	'	docType = GetDocType(doc)	
	'	if docType = docText then
	'		if Right(getLeftPart(doc),1) = " " then
	'			CancelKey = true
	'			doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = " $"
	'			doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
	'		end if	             
	'	end if	    
				
	'Case 46: '�����
	Case 61: '���� "=", ���������� "!=" �� "<>"
		docType = GetDocType(doc)	
		if docType = docText then
			if Right(getLeftPart(doc),1) = "!" then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = "<>"
				doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 2
			'else
			'	CancelKey = true
			'	'doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = " = "
			'	doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = " ="
			'	'if KbdProcs.GetKeyState(VK_SCROLL) = 1 then 
			'	'	message "scroll pressed"
			'	'	CommonScripts.WSH.SendKeys "{ESC}"
			'	'end if	
			'	CommonScripts.WSH.SendKeys " "
			'	'doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 3
			'	doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 2
			end if	             
		end if
		
	'Case 230: '����� "�" - ���������� "�" �� ";" ����� ������
	'	docType = GetDocType(doc)	
	'	if docType = docText then
	'		if Right(getLeftPart(doc),1) = ")" then
	'			CancelKey = true
	'			doc.Range(doc.SelStartLine,doc.SelStartCol-1,doc.SelStartLine, doc.SelStartCol) = ");"
	'			doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
	'		end if	             
	'	end if
	'Case 44: ',
	'Case 43: '+
	'	docType = GetDocType(doc)
	'	if docType = docText then
	'		if (not IsKeyPressed(VK_SHIFT)) and (not IsVirtual) and (Instr(1,GetLeftPart(doc),"//") = 0) then
	'			ClearSelection doc
	'			CancelKey = true
	'			doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = " + "
	'			doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 3
	'		end if
	'	end if
	'Case 45: '-
	'	docType = GetDocType(doc)
	'	if docType = docText then
	'		if (not IsKeyPressed(VK_SHIFT)) and (not IsVirtual) and (Instr(1,GetLeftPart(doc),"//") = 0) then
	'			ClearSelection doc
	'			CancelKey = true
	'			doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = " - "
	'			doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 3
	'		end if
	'	end if
	Case 40: '(
		if not IsVirtual then 
			docType = GetDocType(doc)	
			if docType = docText then
				if not (doc.SelStartCol = doc.SelEndCol) and (doc.SelStartLine = doc.SelEndLine) then '���� ���������
					CancelKey = true
					tText = doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine,  doc.SelEndCol)
					tText = "(" & tText & ")"
					doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine,  doc.SelEndCol) = tText
					doc.MoveCaret doc.SelStartLine,doc.SelStartCol,doc.SelEndLine,doc.SelEndCol+2 
				else
			'		doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = "("
			'		doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
			'		SendCommand 22502 '������������� ������ - ������� ��������� ��������
				end if
			end if	             	
		else '������� "����"
		end if	
	'Case 41: ')
	Case 47: '/
		docType = GetDocType(doc)
		if docType = docText then
			if IsKeyPressed(VK_SHIFT) then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = "|"
				doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
			end if
		end if
		
	Case 63: '"?" - ��������� ���������� �������� "?" �� "?(,,)" ����� �������� ���������� ���������
		docType = GetDocType(doc)
		if docType = docText then
			if GetKeyboardLayout() = klEn then
				CancelKey = true
				doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelStartLine, doc.SelStartCol) = "?"
				doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
			end if
		end if
		
	'Case 1:
	'		if IsKeyPressed(VK_SHIFT) and (IsKeyPressed(VK_CONTROL)) then message "Ctrl+Shift+A"
		
	'Case 113:
		'if IsVirtual then
		'	msgbox "F2"
		'end if	
	'Case 37: '������� ������
		'exit sub
		'if IsKeyPressed(VK_MENU) then
		'	set doc = Windows.ActiveWnd.Document
		'	if doc is nothing then exit sub
		'	if doc.Type <> 0 then exit sub
		'	if doc.Name = "CMDTabDoc::������������" then
		'		if MDWnd.ActiveTab <> 0 then exit sub
		'		strLastMDTreeSelection = Scripts("NavigationTools").strLastMDTreeSelection
		'		if strLastMDTreeSelection <> "" then
		'			CancelKey = true
		'			MDWnd.DoAction strLastMDTreeSelection, 0
		'		end if	
		'	end if	
		'end if	
	End Select
End Sub

'========================================================================================
Private Function getLeftPart(doc)
	getLeftPart = Left(doc.Range(doc.SelStartLine), doc.SelStartCol)
End Function	
                     
'========================================================================================
Private Sub ClearSelection(doc)
	if (doc.SelStartCol<>doc.SelEndCol) or (doc.SelStartLine<>doc.SelEndLine) then
		doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelEndLine, doc.SelEndCol) = ""
	end if	
End Sub
                     
'========================================================================================
Private Function IsKeyPressed(VirtualKeyCode)
	IsKeyPressed = (KbdProcs.GetKeyState(VirtualKeyCode) and 32768) <> 0
End Function

'========================================================================================
Private Function GetKeyboardLayout()
	GetKeyboardLayout = KbdProcs.GetKeyboardLayout and 65535
End Function

'========================================================================================
Private Function GetDocType(ByRef doc)
	GetDocType = -1
	set doc = Windows.ActiveWnd
	if doc is nothing then exit Function
	set doc = doc.Document		
	if doc.Type	= docWorkBook then set doc = doc.Page(doc.ActivePage)
	GetDocType = doc.Type	
End Function

'========================================================================================
Private Sub Debug(s)
	if flDebug then message s
End Sub  

'========================================================================================
Sub ShowPluginState()
	message "������ �������: " & KbdProcs.Enabled
End Sub

'========================================================================================
Sub TogglePluginState()
	KbdProcs.Enabled = not KbdProcs.Enabled
	ShowPluginState()
End Sub

'========================================================================================
Sub ToggleDebug()
	flDebug = not flDebug
end sub

'========================================================================================
Sub ShowActiveLayout()
	l = GetKeyboardLayout() 
	if l = klRu then 
		message "������� ���������: RU"
	elseif l = klEn then 
		message "������� ���������: EN"
	else
		message "������� ���������: " & l
	end if	
End Sub

'========================================================================================
Private Sub Init()
	Set p = Nothing
	On Error Resume Next
	Set p = Plugins("KbdProcs")
	On Error Goto 0
	If p Is Nothing Then
		Message "������ �������� ������� KbdProcs", mRedErr
		Message "������ " & SelfScript.Name & " �� ��������", mInformation
		Scripts.UnLoad SelfScript.Name
		Exit Sub
	End If
	SelfScript.AddNamedItem "KbdProcs", p, False
	
	if not KbdProcs.Enabled then KbdProcs.Enabled = true '���/���� ������
	
	'Set c = Nothing
	'On Error Resume Next
	'Set c = CreateObject("OpenConf.CommonServices")
	'On Error GoTo 0
	'If c Is Nothing Then
	'	Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
	'	Message "������ " & SelfScript.Name & " �� ��������", mInformation
	'	Scripts.UnLoad SelfScript.Name
	'	Exit Sub
	'End If
	'c.SetConfig(Configurator)
	'SelfScript.AddNamedItem "CommonScripts", c, False
	flDebug = false
End Sub

'========================================================================================
Init