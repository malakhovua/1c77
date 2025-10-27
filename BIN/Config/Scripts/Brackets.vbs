'������,  �����������  ��������� ���������� ����� � ��������� ��������� 
'������������, � � ������� ������, ��������� MetaEditor-��, ���� �������� 
'����� �� �����.
'
'������: $Revision: 1.2 $
'�����: trdm � MetaEditor (c) 2005 �.
'
'������: ���� �����:"���� ������� = ������������� �����"
'        
'�� ������ ��������� ���������� "�������������" � �������.
'����������� ������ �� ������������� � �������� ������, ��� ��������� 
'������ "�����������": "."; (.); ��������(.); � ������ ����. �������� "." 
'� ��� ����� ����� � �������� ).
'
'����������: 
'- ����� � ������������ �������� ������ ������� ������ ����������� ��������� 
'��� �������� �����, �����������; 
'- "�" �������� �������� ������� ����� ������� ����� ����������� "�����������".
'
'������ ������� "�����������":
'""".""", 
'"(.)", 
'"��������(.)", 
'"������������(.)", 
'"��������������(.) = 1",
'"?(.,�,)",
'"�����(.)",
'"�����(.)", 
'"������(.)", 
'"���(.,�)",
'"����(.,�)",
'"����(,�)",
'"�����(.,�)",
'"�����������(.,�,)",
'"�����������������(.,�)",
'"������������������(.)",
'"����(.)", 
'"����(.)",
'"OemToAnsi(.)",
'"AnsiToOem(.)",
'"����(.)",
'"�������(.)", 
'"��������������(.)"
'... ���������� �� ������ ������ �� ��������������� �� ���� ����...

'31.10.2005
'  ������ �������� ���������, �� ������� ������������� � �������� ������
'  ���������� ����_�����_���������, ���_����, ����_����...

Dim glBracket
Dim BracketsDict
Dim tTextOfKomment
Dim tTextOfKommentWithDade

NameDelimiters = " .,;:|#=+-*/%?<>\()[]{}!~@$^&'""" & vbTab
glBracket = """.""" 
tTextOfKomment = ""
tTextOfKommentWithDade = ""

'========================================================================================
Function GetCurrentWordBorders(Line, CursorPos, Delimiters)
	Dim Col1, Col2, TextLen
	
	GetCurrentWordBorders = ""
	TextLen = Len(Line)
	Col1 = CursorPos
	
	do while Col1 > 0
		If InStr(Delimiters, Mid(Line, Col1, 1)) > 0 Then
			Col1 = Col1 + 1
			Exit Do
		End If
		Col1 = Col1 - 1
	loop
	
	If Col1 = 0 Then Col1 = 1
	
	Col2 = Col1
	do while Col2 <= TextLen
		If InStr(Delimiters, Mid(Line, Col2, 1)) > 0 Then
			Col2 = Col2 - 1
			Exit Do
		End If
		Col2 = Col2 + 1
	loop
	
	If Col2 > TextLen Then Col2 = TextLen
	GetCurrentWordBorders = Array(Col1, Col2 - Col1 + 1)
	
End Function 'GetCurrentWordBorders 

'========================================================================================
Sub AddBracket()
    Dim flMultiLine
	
	set doc = CommonScripts.GetTextDocIfOpened(0)
	If doc Is Nothing Then Exit Sub
  
	flMultiLine = false
	locBracket = glBracket
	dY = 0
	if InStr(1,locBracket,"_") > 0 then
		for i = 1 to Len(locBracket)
			if Mid(locBracket,i,1) = "_" then dY = dY + 1
		next
		locBracket = Replace(locBracket,"_",vbCrLf)
		flMultiLine = true
	end if	
	ArrOfBracket = Split(locBracket,".")

	flCurrWord = false
	if (Doc.SelStartCol = Doc.SelEndCol) and (Doc.SelStartLine = Doc.SelEndLine) then flCurrWord = true
	
	If (Doc.SelStartLine <> Doc.SelEndLine) Or InStr(1,glBracket,"_") Then '���� ��� ��������� ������������� �����������.....
		if doc.SelEndCol <> 0 Then
			Doc.MoveCaret doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine)
		Else
			Doc.MoveCaret doc.SelStartLine, 0, doc.SelEndLine-1, doc.LineLen(doc.SelEndLine-1)
		End If
	End If

	Komment = ""
	if ((tTextOfKomment = glBracket) Or (tTextOfKommentWithDade = glBracket))And Not flCurrWord Then
		Komment = InputBox("�����������","������� �����������","")
		if Len(Komment)>0 And glBracket = tTextOfKommentWithDade Then 
			ArrOfBracket(0) = Replace(ArrOfBracket(0),"date", Trim(CStr(Now())))
			ArrOfBracket(1) = Replace(ArrOfBracket(1),"date", Trim(CStr(Now()))+ " ")
		elseIf Len(Komment) = 0 Then '���������� �� �����.....
			' ������ ������� "date" �� ������....
			ArrOfBracket(0) = Replace(ArrOfBracket(0),"date", "")
			ArrOfBracket(1) = Replace(ArrOfBracket(1),"date", " ")
		End If
		ArrOfBracketOne = Split(ArrOfBracket(0),vbcrlf) 
		ArrOfBracketOne(0) = ArrOfBracketOne(0) & Komment
		ArrOfBracket(0) = Join(ArrOfBracketOne,vbcrlf)
	End If
	
	if flCurrWord = true then	
		Borders = GetCurrentWordBorders (Doc.Range(Doc.SelStartLine), Doc.SelStartCol, NameDelimiters)	
		tText = doc.Range(Doc.SelStartLine, Borders(0)-1, Doc.SelEndLine,  Borders(0) + Borders(1) - 1)
		tText = ArrOfBracket(0) & tText & ArrOfBracket(1)
		Doc.Range(Doc.SelStartLine , Borders(0)-1 , Doc.SelEndLine , Borders(0) + Borders(1) - 1) = tText
	else
		tText = Doc.Range(Doc.SelStartLine , Doc.SelStartCol , Doc.SelEndLine , Doc.SelEndCol)
		tText = ArrOfBracket(0) & tText & ArrOfBracket(1) & Komment 
		Doc.Range(Doc.SelStartLine , Doc.SelStartCol , Doc.SelEndLine , Doc.SelEndCol) = tText
	end if
  
	if flCurrWord then	
		Doc.MoveCaret Doc.SelStartLine, Borders(0)-1 , Doc.SelEndLine + dY ,Borders(0) + Borders(1) + Len(ArrOfBracket(0)) + Len(ArrOfBracket(1)) - 1
	else
		Doc.MoveCaret Doc.SelStartLine , Doc.SelStartCol , Doc.SelEndLine + dY, Doc.SelEndCol+2+Len(ArrOfBracket(0))-1
	end if
	
	if flMultiLine then Doc.FormatSel
	
	for i = Doc.SelStartLine to Doc.SelEndLine
		PositionMoveKaret = InStr(1,Doc.Range(i),"�")
		if PositionMoveKaret > 1 then
			Doc.Range(i,  PositionMoveKaret - 1, i, PositionMoveKaret) = ""
			Doc.MoveCaret i , PositionMoveKaret - 1
			exit for
		end if	
	next
	
End Sub 'AddBracket  

'========================================================================================
Sub ChoiseTypeBracket()	
	If CommonScripts.GetTextDocIfOpened(0) Is Nothing Then Exit Sub
	tText = CommonScripts.SelectValue(BracketsDict,"")  
	If Len(tText) > 0 Then
		glBracket = tText 
		AddBracket()
	End If  
End Sub  'ChoiseTypeBracket

'========================================================================================
Private Sub InitDict()
	Set BracketsDict = CreateObject("Scripting.Dictionary")
	
	BracketsDict.Add """ """, """."""
	BracketsDict.Add "( )","(.)"  
	BracketsDict.Add "[ ]","[.]"  
	BracketsDict.Add "������()", "������(.)"
	BracketsDict.Add "����()", "����(.)"
	BracketsDict.Add "�����()", "�����(.)"
	BracketsDict.Add "��������()", "��������(.)"
	BracketsDict.Add "��������()", "��������(.)" 
	BracketsDict.Add "������������()","������������(.)" 
	BracketsDict.Add "��������������() = 1", "��������������(.) = 1"
	BracketsDict.Add "?(,,)", "?(.,�,)"
	BracketsDict.Add "�����()", "�����(.)"
	BracketsDict.Add "�����()","�����(.)" 
	BracketsDict.Add "������()","������(.)" 
	BracketsDict.Add "���()", "���(.,�)"
	BracketsDict.Add "����()", "����(.,�)"
	BracketsDict.Add "����()", "����(,�)"
	BracketsDict.Add "�����()", "�����(.,�)"
	BracketsDict.Add "�����������()","�����������(.,�,)" 
	BracketsDict.Add "�����������������()","�����������������(.,�)" 
	BracketsDict.Add "������������������()", "������������������(.)"
	BracketsDict.Add "����()", "����(.)"
	BracketsDict.Add "����()", "����(.)"
	BracketsDict.Add "OemToAnsi()", "OemToAnsi(.)"
	BracketsDict.Add "AnsiToOem()", "AnsiToOem(.)"
	BracketsDict.Add "����()", "����(.)"
	BracketsDict.Add "�������()", "�������(.)"
	BracketsDict.Add "��������������()", "��������������(.)"
	BracketsDict.Add "���()", "���(.)"
	BracketsDict.Add "���()","���(.)" 
	BracketsDict.Add "���()", "���(.,�)"
	BracketsDict.Add "����()", "����(.,�)"
	BracketsDict.Add "������()", "������(.)"
	BracketsDict.Add "��� - ���� - ����������", "��� �= ��  ����_._����������;"
	BracketsDict.Add "���� - ���� - ����������", "���� � ����_._����������;"
	BracketsDict.Add "������� - ���������� - ������������", "�������_._�_����������_������������;"
	BracketsDict.Add "��������� - ��������������", "//" & String(70,"=") & "_��������� �()_._��������������; //"
	BracketsDict.Add "������� - ������������","������� �()_._������������;"
	BracketsDict.Add "�����������","//�_._//"
	BracketsDict.Add "����������� + ����","// date �_._//"
	BracketsDict.Add "����","//{�_._//}" 
	BracketsDict.Add "����������","����������������();_._�����������������������();�"
	BracketsDict.Add "�������������","�������������()�;_���� ��������������() = 1 ����_._����������;"
	BracketsDict.Add "���� - ����� - ����� - ���������", "���� � �����_._�����__���������;"
	BracketsDict.Add "���� * ����� - ���������", "���� . � �����__���������;"
	BracketsDict.Add "���� - ����� * ���������", "���� � �����_._���������;"
	
	tTextOfKomment = "//�_._//"
	tTextOfKommentWithDade = "// date �_._//"
End Sub 'InitDict

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
	InitDict()
End Sub

'========================================================================================
Init ' ��� �������� ������� ��������� �������������
