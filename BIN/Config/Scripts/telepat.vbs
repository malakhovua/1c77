' ������, ���������� � ���� � �������� "�������"
' � �������������� ������� �� ����
' ���������� � bin\config\scripts

' ��������� ������� "������� �������".
' ��������� ��� ������� ������� ����� ���������� ��� 1�.
' ��������� �������� ����� ������, �������������� ��� �� ��������.
' name - ��� ������������ �������.
' text - ����� ������ �������. ����� �������� ���.
' cancel - ���� ������. ��� ��������� � True ������� ������� ����������.
'
Sub Telepat_OnTemplate(Name, Text, Cancel)
    Select Case Name
        Case "������� ����"
            Text = "// " & Trim(CStr(Now()))
        Case "���������� ����"
            Cancel = FillDoc(Text)
        Case "���������� ��"
            Cancel = FillTablePart(Text)
        Case "���������� ��������"
            Cancel = FillDocRegister(Text)
        Case "����� �������"
            Cancel = FillNewRefItem(Text)
        Case "����� ������"
            Cancel = FillNewRefGroup(Text)
		Case "�� ���� � ���"
			Cancel=CopyDocFromDoc(Text)
    End Select
End Sub

' ���������� ������� "����� ���� ��������"
' ��������� ����������� �������� ������� � ���� ��������.
' ��� ����� ������ ������� ������-��������� ����������� �������.
' ������ ����������� ����� ������ ������������� �� ��������� ������.
' ��������� ������ ������ ������������ �� ��������� � ������ ������.
' ����� �������� ������� ����� ������ | ����� ������� �������
' d ��� D (�� Disabled) - ����������� �����
' c ��� C (�� Checked)  - ����� � "��������"
' ����� ����� ������ | ����� ������� ������������� �������.
' � ���� ������ � ������� OnCustomMenu ������ �������� ������ ����
' ����� ������� ���� �������������
' ��� �������� ������������ ������� ��� "-"
'
Function Telepat_GetMenu()
	'message "menu"
    Telepat_GetMenu = _
        "������ ����� ������" & vbCrLf & _
        vbTab & "����������� �������|d" & vbCrLf & _
        vbTab & "���������� �����|c|command1" & vbCrLf & _
        vbTab & "������ ����� ����| |command2" & vbCrLf & _
        vbTab & "-" & vbCrLf & _
        vbTab & "��� �������" & vbCrLf & _
        "��� ���� �������" & vbCrLf & _
        "� ��� ����" & vbCrLf & _
        "-" & vbCrLf & _
        "����������� ����������|dc"
End Function

' ���������� ������� OnCustomMenu.
' ���������� ��� ������ ������������� ������ ����,
' ������������ � "GetMenu"
' Cmd - �������� (��� �������������) ���������� ������ ����.
'
Sub Telepat_OnCustomMenu(Cmd)
    Message Cmd, mNone
End Sub

' ���������� ������� ������� ������ �� ������ ����������
' ��������� �������� ����� �������.
' InsertType - ��� ������������ �������� (��������� ����)
' InsertName - ��� ������������ �������� (��� ��� ������� � ������ ����������)
' Text - ����������� �����
' �� ����������� ������ �������������� ����� "!" ���������� ����������
' ������� ����� �������. (�������� ��������� ������ ��� ������������ �������)
' ���� ��������� ������� �� �������, �� �� ��������������� � ����� ������.
' ��� ������� ������� �� ������ ���������� ������ ���������� �� ����������.
' ��� ������� ��������, ��� ������ �,���,�� ��������� �,���,��
'
Sub Telepat_OnInsert(InsertType, InsertName, Text)
	'message InsertType
    Select Case InsertType
        Case 0          ' �������������� ���������� �������� ������
			'Text = Text & " !"
        Case 1          ' ���������������� ���������� �������� ������
			'Text = Text & " !"
        Case 2          ' �������������� ����� �������� ������
		Case 3          ' ���������������� ����� �������� ������
        Case 4          ' �������� �������
			'Text = Text & " !"
        Case 5          ' �������������� ���������� ����������� ������
			'Text = Text & " !"
        Case 6          ' ���������������� ���������� ����������� ������
			'Text = Text & " !"
        Case 7          ' �������������� ����� ����������� ������
        Case 8          ' ���������������� ����� ����������� ������
        Case 9          ' ��� ����������� �������� (���� ����������������, ������������ � ��)
        Case 10         ' ����� 1�
            ' ����� ��� �� �������� ��������� ����������� �������������...
            If InsertName = "�������������" Then 
				Text = "�������������(""!"");"
			ElseIf InsertName = "�������������" Then 
				Text = "�������������();" & vbCrLf & vbTab & "���� ��������������() = 1 ����" & vbCrLf & vbTab & vbTab & "!" & vbCrLf & vbTab & "����������;"
			End If
			'Text = InsertName & "(""!"");" 
        Case 11         ' �������� �����
            If InStr(",���,��,", InsertName) > 0 Then Text = LCase(Text)
        Case 12         ' ��������� ����������
			'Text = Text & " !"
        Case 13         ' ������� ���������� �������� ������
        Case 14         ' ���������������� �����
        	If InsertName="���������_��������������" or InsertName="�������_������������" Then
        		ProcName=InputBox("������� ��� ������","�����", "��")
				
				tempText = Replace(Text,vbCrLf,vbCrLf & vbTab  & "!" & vbCrLf)
				
				tempText = "//"+String(70,"=") & vbCrLf & _ 
				Replace(tempText,"!(",ProcName & "(") & " // " & ProcName
				
				Text = tempText
        		'Text= "//"+String(70,"=") & vbCrLf & Replace(Text,"!(",ProcName & "(!") & vbTab & "// " & ProcName
        	End If
        Case 15         ' ����������� � ������
    End Select
End Sub

' ������ ���������� ���������
Function FillDoc(Text)
    FillDoc = True      ' �� ��������� ������� ������� �������
    ' �������� ������������ ������� ��� ���������
    DocKind = ""
    Set MetaDoc = SelectMetaObj(DocKind, "��������", "������� ��� ���������")
    If MetaDoc Is Nothing Then Exit Function    ' �������� �� ������

    ' ����� ������������ �������, ��� ������� ����������
    docName = InputBox("������� ��� ���������� ���������", "���������� ���������", "���" & DocKind)
    If Len(docName) = 0 Then Exit Function
    FillDoc = False ' �� ����� �������� ������� �������

    ' ��������� � ������ ��������� ��� ���������
    Text = Replace(Text, "!selDoc!", DocKind)
    ' ��������� � ������ ��������� ��� ����������
    Text = Replace(Text, "!selVar!", docName)
    
    ' ������ �������� ������ ������������ �����������
    Set Head = MetaDoc.Childs("�������������")
    Set Table = MetaDoc.Childs("����������������������")
    Set CommonRekv = MetaData.TaskDef.Childs("����������������������")
    ' ��������� ���������� ����� � ���������� �������
    ' � ������ ������������
    ReDim Lines(2 + CommonRekv.Count + Head.Count + Table.Count)
    
    docName = docName & "."
   
    iCurLine = 0        ' ������ ����������� ������
    ' �������� ����� ��������� ���������
    FillNamesLines "// ����� ���������", docName, CommonRekv, Lines, iCurLine
    ' �������� ��������� �����
    FillNamesLines "// ��������� �����", docName, Head, Lines, iCurLine
    ' �������� ��������� ��
    FillNamesLines "// ��������� ��������� �����", docName, Table, Lines, iCurLine
    ' ��������� � ������ ���������� �����
    Text = Text & Join(Lines, vbCrLf)
End Function

' ��������� ������� ���������� �� ���������
' ��������� ������ ���� ��������.�������������� = ��������.��������������;
Function FillTablePart(Text)
    FillTablePart = True    ' �� ��������� ������� ������� �������
    ' �������� ������������ ������� ��� ���������
    DocKind = ""
    Set MetaDoc = SelectMetaObj(DocKind, "��������", "������� ��� ���������")
    If MetaDoc Is Nothing Then Exit Function
    ' ������� MetaArray c ��������� ���������� �� ���������� ���������
    Set Table = MetaDoc.Childs("����������������������")
    iCountLines = Table.Count
    If iCountLines = 0 Then
        MsgBox "� ��������� """ & DocKind & """ ��� ���������� ��", , "�������"
        Exit Function
    End If
    
    ' �������� � ������������ ��� ���������� ���������
    selFrom = InputBox("�������, ������ ���������� ��������", "���������� ��", "��")
    If Len(selFrom) = 0 Then Exit Function
    ' �������� � ������������ ��� ���������� ���������
    selTo = InputBox("�������, ���� ���������� ��������", "���������� ��", "���")
    If Len(selTo) = 0 Then Exit Function
    
    FillTablePart = False   ' ������������ ��� ������. �������� ������� �������
    ' ��������� � ������ ��� ���������� ���������
    Text = Replace(Text, "!selFrom!", selFrom)
    ' �������� ������������ ����� ������� ����������� ��
    ReDim Lines(iCountLines - 1)
    selTo = vbTab & selTo & "."
    iMaxNameLen = MaxNameLen(Table, 0) + Len(selTo)
    selFrom = " = " & selFrom & "."
    For i = 0 To iCountLines - 1
        Lines(i) = selTo & Table(i).Name
        Lines(i) = Lines(i) & Space(iMaxNameLen - Len(Lines(i))) & selFrom & Table(i).Name & ";"
    Next
    ' ��������� � ������ �������������� ������������ �����
    Text = Replace(Text, "!dynamic_part!", Join(Lines, vbCrLf))
End Function

' �������� ������� "���������� ��������"
Function FillDocRegister(Text)
    FillDocRegister = True  ' �� ��������� ������� ������� �������
    ' �������� ������������ ������� ��� ��������
    RegName = ""
    Set MetaReg = SelectMetaObj(RegName, "�������", "������� ��� ��������")
    If MetaReg Is Nothing Then Exit Function
    
    ' ��������, ������ ���������� ������
    Set TextDoc = Windows.ActiveWnd.Document
    If TextDoc = docWorkBook Then Set TextDoc = TextDoc.Page(1)
    docVarName = ""
    If TextDoc.Kind <> "Transact" Then
        ' ���� ������ ���������� �� �� ������ ���������,
        ' ���� �������� ��� ���������� ���������
        docVarName = InputBox("������� ��� ���������� ���������", "���������� ���������", "����")
        If Len(docVarName) > 0 Then docVarName = docVarName & "."
    End If
    ' �������� ������������ ������� ��� ���������� ��� ����������� ��������
    regVarName = InputBox("������� ��� ���������� ��������", "���������� ���������", "���" & RegName)
    If Len(regVarName) = 0 Then Exit Function
    ' ������������ ��� ������, �������� ������� �������
    FillDocRegister = False
    ' ��������� � ������ ��������� ����������
    Text = Replace(Text, "!regVar!", regVarName)
    Text = Replace(Text, "!docVar!", docVarName)
    Text = Replace(Text, "!regName!", RegName)
    ' ������� ������ � ��������
    Set Fields = MetaReg.Childs("���������")
    Set Resurs = MetaReg.Childs("������")
    Set Rekv = MetaReg.Childs("��������")

    ' ���������� ������ ��� �����
    ReDim Lines(2 + Fields.Count + Resurs.Count + Rekv.Count)

    regVarName = vbTab & regVarName & "."
    
    iCurLine = 0        ' ������ ����������� ������
    ' �������� ���������
    FillNamesLines vbTab & "// ��������� �������� " & RegName, regVarName, Fields, Lines, iCurLine
    ' �������� ���������
    FillNamesLines vbTab & "// ��������� �������� " & RegName, regVarName, Rekv, Lines, iCurLine
    ' �������� �������
    FillNamesLines vbTab & "// ������� �������� " & RegName, regVarName, Resurs, Lines, iCurLine
    ' ��������� � ������ ������������ �����
    Text = Replace(Text, "!dynamic_part!", Join(Lines, vbCrLf))
End Function

Function CopyDocFromDoc(Text)
	CopyDocFromDoc=True
	docKindTo=""
	Set MetaTo=SelectMetaObj(docKindTo,"��������","������� ��� ������������ ���������")
	if MetaTo Is Nothing Then Exit Function

	docKindFrom=""
	Set MetaFrom=SelectMetaObj(docKindFrom,"��������","������� ��� ������������ ���������")
	if MetaFrom Is Nothing Then Exit Function

	docTo="���" & docKindTo
	docFrom="���" & docKindFrom
	if docKindFrom=docKindTo Then
		docTo=docTo & "1"
		docFrom=docFrom & "2"
	End If
	docTo=InputBox("���������� ������������ ����","�� ���� � ���",docTo)
	if len(docTo)=0 then exit function
	docFrom=InputBox("���������� ������������ ����","�� ���� � ���",docFrom)
	if len(docFrom)=0 then exit function
	if docTo=docFrom Then
		MsgBox "�������� � �������� ���������",,"�������"
		exit function
	end if
		
	CopyDocFromDoc=False
	Text=Replace(Text,"!docTo!",docTo)
	Text=Replace(Text,"!docFrom!",docFrom)
	docTo=docTo & "."
	docFrom=docFrom & "."
	' �������� ����� ���������
    Set CommonRekv = MetaData.TaskDef.Childs("����������������������")
    FillHead=""
    ReDim Lines(CommonRekv.Count)
    Lines(0)="// ����� ��������� ���������"
    iMaxNameLen=0
    for i=0 to CommonRekv.Count-1
    	Name=CommonRekv(i).Name
    	Lines(i+1)=docTo & Name & "! = "  & docFrom & Name & ";"
    	if len(Name)>iMaxNameLen Then iMaxNameLen=Len(Name)
    next
    iMaxNameLen=iMaxNameLen+Len(docTo)+1
    for i=1 to CommonRekv.Count
    	line=Lines(i)
    	l=InStr(line,"!")
    	Lines(i)=Replace(line,"!",Space(iMaxNameLen-l))
    next
    FillHead=Join(Lines,vbCrLf) & vbCrLf
    ' �������� ����������� ��������� �����
    Set HeadTo = MetaTo.Childs("�������������")
    Set HeadFrom = MetaFrom.Childs("�������������")
    ReDim Lines(HeadTo.Count)
    Lines(0)="// ��������� �����"
    iMaxNameLen=0
    On Error Resume Next
    for i=0 to HeadTo.Count-1
    	Name=HeadTo(i).Name
    	RightPart=""
    	if len(Name)>iMaxNameLen Then iMaxNameLen=Len(Name)
    	Lines(i+1)=docTo & Name & "! = "
    	Set rekv=Nothing
    	Set rekv=HeadFrom(CStr(Name))
    	if not rekv is nothing then
    		TypeTo=HeadTo(i).Props("���")
    		KindTo=HeadTo(i).Props("���")
    		TypeFrom=rekv.Props("���")
    		KindFrom=rekv.Props("���")
    		if TypeTo=TypeFrom and (KindFrom=KindTo or KindTo="") Then RightPart = docFrom & Name
    	End If
    	Lines(i+1)=Lines(i+1) & RightPart & ";"
	next
    iMaxNameLen=iMaxNameLen+Len(docTo)+1
    for i=1 to HeadTo.Count
    	line=Lines(i)
    	l=InStr(line,"!")
    	Lines(i)=Replace(line,"!",Space(iMaxNameLen-l))
    next
    Text=Replace(Text,"!fill_head!",FillHead & Join(Lines,vbCrLf))
    ' �������� ��������� ��
    Set TableTo = MetaTo.Childs("����������������������")
    Set TableFrom = MetaFrom.Childs("����������������������")
    docTo=vbTab & docTo
    ReDim Lines(TableTo.Count)
    Lines(0)=vbTab & "// ��������� ��������� �����"
    iMaxNameLen=0
    On Error Resume Next
    for i=0 to TableTo.Count-1
    	Name=TableTo(i).Name
    	RightPart=""
    	if len(Name)>iMaxNameLen Then iMaxNameLen=Len(Name)
    	Lines(i+1)=docTo & Name & "! = "
    	Set rekv=Nothing
    	Set rekv=TableFrom(CStr(Name))
    	if not rekv is nothing then
    		TypeTo=TableTo(i).Props("���")
    		KindTo=TableTo(i).Props("���")
    		TypeFrom=rekv.Props("���")
    		KindFrom=rekv.Props("���")
    		if TypeTo=TypeFrom and (KindFrom=KindTo or KindTo="") Then RightPart = docFrom & Name
    	End If
    	Lines(i+1)=Lines(i+1) & RightPart & ";"
	next
    iMaxNameLen=iMaxNameLen+Len(docTo)+1
    for i=1 to TableTo.Count
    	line=Lines(i)
    	l=InStr(line,"!")
    	Lines(i)=Replace(line,"!",Space(iMaxNameLen-l))
    next
    Text=Replace(Text,"!fill_table!",Join(Lines,vbCrLf))
End Function



' ���������� ������� "����� ������� �����������"
Function FillNewRefItem(Text)
    FillNewRefItem = FillNewReference(Text, False)
End Function
' ���������� ������� "����� ������ �����������"
Function FillNewRefGroup(Text)
    FillNewRefGroup = FillNewReference(Text, True)
End Function

' ��������� ������� "����� �������/������ �����������"
Function FillNewReference(Text, IsGroup)
    FillNewReference = True ' �������� ������� �������
    refKind = ""
    Set MetaRef = SelectMetaObj(refKind, "����������", "������� ��� �����������")
    If MetaRef Is Nothing Then Exit Function
    LevelsCount = CLng(MetaRef.Props("�����������������"))
    If IsGroup = True And LevelsCount = 1 Then
        MsgBox "� ����������� """ & refKind & """ �� ����� ���� �����!", , "�������"
        Exit Function
    End If
    varName = InputBox("������� ��� ����������", "����� �������/������", "���" & refKind)
    If Len(varName) = 0 Then Exit Function
    FillNewReference = False    ' ��������� ������� �������
    
    Text = Replace(Text, "!varName!", varName)
    Text = Replace(Text, "!refKind!", refKind)
    varName = varName & "."
    
    ' ������� ��������� �����������
    Set Fields = MetaRef.Childs("��������")
    iMaxNameLen = 0
    iLineCount = Fields.Count
    ReDim Lines(iLineCount)
    Lines(0) = "// ��������� �����������"
    iCurLine = 1
    ' �������, ���� �� �������� ���
    If CLng(MetaRef.Props("���������")) > 0 Then
        iLineCount = iLineCount + 1
        ReDim Preserve Lines(iLineCount)
        Lines(1) = varName & "���!"
        iCurLine = 2
        iMaxNameLen = 3
    End If
    ' �������, ���� �� �������� ������������
    If CLng(MetaRef.Props("�����������������")) > 0 Then
        iLineCount = iLineCount + 1
        ReDim Preserve Lines(iLineCount)
        Lines(iCurLine) = varName & "������������!"
        iCurLine = iCurLine + 1
        iMaxNameLen = 12
    End If
    ' �������� ������ ���� ���������� � �� �������������
    HavePeriodicVal = False
    For i = 0 To Fields.Count - 1
        Set Rekv = Fields(i)
        UseIn = Rekv.Props("�������������")
        If (UseIn = "���������" And IsGroup = False) _
            Or (UseIn = "�����������" And IsGroup = True) Then
            iLineCount = iLineCount - 1
        Else
            If Rekv.Props("�������������") = "1" Then HavePeriodicVal = True
            Name = Rekv.Name
            Lines(iCurLine) = varName & Name & "!"
            If Len(Name) > iMaxNameLen Then iMaxNameLen = Len(Name)
            Descr = Rekv.Descr
            If Len(Descr) > 0 And Descr <> Name Then
                Lines(iCurLine) = Lines(iCurLine) & " // " & Descr
            End If
            iCurLine = iCurLine + 1
        End If
    Next
    ' �������� ������ ������������ ������� =
    iMaxNameLen = iMaxNameLen + Len(varName) + 1
    ReDim Preserve Lines(iLineCount)
    For i = 1 To iLineCount
        Line = Lines(i)
        l = InStr(Line, "!")
        Lines(i) = Replace(Line, "!", Space(iMaxNameLen - l) & " = ;")
    Next
    Text = Replace(Text, "!dynamic_part!", Join(Lines, vbCrLf))
    
    ' �������, ����������� �� ����������, ����� �� �� ����������.
    ' ���������, ������� � ��� �������
    ParentPeriodic = ""
    If HavePeriodicVal = True Then ParentPeriodic = varName & _
        "����������������(�����������());" & vbCrLf
    If Len(MetaRef.Props("��������")) > 0 Then ParentPeriodic = _
        ParentPeriodic & varName & "���������������������(...); //" & MetaRef.Props("��������") & vbCrLf
    If LevelsCount > 1 Then ParentPeriodic = ParentPeriodic & _
        varName & "��������������������(...);" & vbCrLf
    Text = Replace(Text, "!parent_periodic!", ParentPeriodic)
End Function
    
' ������� ������ ���������� 1� ��������� ����� ����������
Function SelectMetaObj(objName, objType, Title)
    ' ��� ������ ������������� ������� �������� ConvertTemplate
    ' ������ ����� ������������ ��������� �����������
    ' ������ ������� ����������� 1�-�������
    ' � ���������� ������������ �����
    ' � ������ ������ ����������� ����� ����� �� �����������
    ' ������� �������� ��� ������ ������� ����������,
    ' �������� <?"������� ��������",��������>
    objName = Telepat.ConvertTemplate("<?""" & Title & """," & objType & ">")
    If Len(objName) = 0 Then Set SelectMetaObj = Nothing: Exit Function
    ' ������ �������� ���������� ��� ���������� �������
    Set SelectMetaObj = MetaData.TaskDef.Childs(CStr(objType))(CStr(objName))
End Function

' ������� ���������� ������������ ����� ��������������
' � ������� ����������. ������������ ��� ������������ ������ =
Function MaxNameLen(MetaArray, OldMaxLen)
    MaxNameLen = OldMaxLen
    For i = 0 To MetaArray.Count - 1
        If Len(MetaArray(i).Name) > MaxNameLen Then MaxNameLen = Len(MetaArray(i).Name)
    Next
End Function

' ������� ��������� ������ �����, ������� ���������,
' �������� ����
' LeftPart ������������ = ; // ����������� ���������
' � ������������� ������ =
Sub FillNamesLines(Title, LeftPart, MetaRekvArray, Lines(), iCurLine)
    iMaxNameLen = MaxNameLen(MetaRekvArray, 0) + Len(LeftPart)
    Lines(iCurLine) = Title
    iCurLine = iCurLine + 1
    For i = 0 To MetaRekvArray.Count - 1
        Line = LeftPart & MetaRekvArray(i).Name
        Line = Line & Space(iMaxNameLen - Len(Line)) & " = ;"
        Descr = MetaRekvArray(i).Descr
        If Len(Descr) > 0 Then
            If Descr <> MetaRekvArray(i).Name Then
                Line = Line & " // " & Descr
            End If
        End If
        Lines(iCurLine) = Line
        iCurLine = iCurLine + 1
    Next
End Sub


Sub Configurator_OnMsgBox(Text, Style, DefAnswer, Answer)
'message "Configurator_OnMsgBox" & Text 
End Sub

' ������������� �������. param - ������ ��������,
' ����� �� ������� � �������
'
Sub Init(param)
	Set t = Nothing
    On Error Resume Next
	Set t = Plugins("�������")  ' �������� ������
	On Error Goto 0
    If Not t Is Nothing Then    ' ���� "�������" ��������
        ' ����������� ������ � �������� �������
        SelfScript.AddNamedItem "Telepat", t, False
        ' ������ ������ �������� ��� ������ Telepat
        ' �������� ��� ���������
        ' ���������������� ������ ������
        Telepat.Components = 1 + 2 + 4                 ' ������������ ����������. 1-��������, 2 - �����������, 4 - ������
        Telepat.Language = 2                        ' ������������ �����. 1- ����������, 2 - �������
        Telepat.UseStdMethodDlg = False                 ' ������������ ����������� ������ "������ ������"
        Telepat.NoOrderMethodDlg = False                ' �� ����������� ������ � ������� "������ ������"
        Telepat.FilterMethodDlg = True                  ' ����������� ������ � ������� "������ ������"
        Telepat.AutoParamInfo = True                    ' ������������� ��������� � ����������
        Telepat.ParamInfoAddMethDescr = True            ' � ��������� � ���������� �������� �������� ������
        Telepat.ParamInfoAddParamDescr = True           ' � ��������� � ���������� �������� �������� ���������
        Telepat.AutoActiveCountSymbols = 2              ' ���������� �������� � �������������� ��� ��������������
        Telepat.DisableTemplateInRemString = 1 + 2      ' ��������� �������. 1-� ������������, 2-� �������
        Telepat.AddTemplate = False                      ' ��������� ������� � ������ ����������
    Else
        ' ������ �� ��������. �������� � ������
        Scripts.Unload SelfScript.Name
    End If
End Sub

' ��� �������� ������� �������������� ���
Init 0