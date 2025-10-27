' -----------------------------------------------------------------------------
'       ������� �������� aka slavka
'       e-mail: webmastert@mail.ru
'       ICQ:286-688-594
'	������: $Revision: 1.1 $
' -----------------------------------------------------------------------------
'����� �������� ��� ������ � �������: ��������� ����������� �������������� �����
'������ (�-�� OfficeXP/2003), ������� ������� ������ � ������� �� ��������������.
'��� ������ ������� ����������:
'	1- Svcsvc.Service (svcsvc.dll)
'	2- OpenConf.CommonServices (commonservices.wsc)
'	3- OpenConf.Registry (registry.wsc)
'
'	������� CopyText, PasteText, EmptyBufer, EmptyAll � Settings ����������
'������ � �������.
'	������ Settings ��������� ���������� ���������: ����� ������� � ������������ �����.
'���� ����� �������� �� ����� �������������� ����������� �����. ���� ��������� ������ 
'������������ ����� �� ��� ����������� ������ ����� ���������� ������ �� ������� �����
'� ����������� ������. ����� ����� ����� �������������� ��� ������ ����������� ������
'� �������� PasteText � EmptyBufer. ���� ����� �� ������������ �� � �������� PasteText
'� EmptyBufer ����� ������� ��� �����.
'	������ CopyText (Ctrl-C)- �������� ����� � �����. ���� ���� �� ���������, �� �����������
'�������������� ����������� �������.
'	������ PasteText (Ctrl-V)- ������������ ������� ������. ���� ���� �� ���������, �� �������
'���������� �� ������������ ������.
'   ������ EmptyBufer (Ctrl-Shift-D) - ������� �� ������ ��������� ��������. 
'   ������ EmptyAll (Ctrl-D) - ������� �� ������ ��� ����������.
'
'   ������ Translit - ����� ������� ��������������.(����� trdm, ��������� MetaEditor)
'   ������ ToggleCase - ����� �������� ��� � ����� �� Shift-F3. (����� MetaEditor)
'
'������������� �� ������ ��� ���������� ������� MetaEditor'�, a13x, artbear'� � trdm'�.

Dim svc, CommonScripts, Buf, ff, curdoc, CurrentText, IsEnable, UseMet, Reg

Sub CopyText() ' ����������� �����
If IsEnable = "0" Then '�� ������������ ������ 
	CommonScripts.SendCommand(57634)
	Exit Sub	
End If	
If CommonScripts.IsTextWindow() = true   Then
	Set doc = CommonScripts.GetTextDoc(0)
	If doc Is Nothing Then Exit Sub
	Line1 = doc.SelStartLine
	Line2 = doc.SelEndLine
	Col1 = doc.SelStartCol
	Col2 = doc.SelEndCol
	
	If Col1 <> Col2 OR Line1<>Line2 Then 
		CurrentText = doc.Range(Line1, Col1, Line2, Col2)
	Else 
		'CurrentText = ""
		CurrentText = doc.CurrentWord
	End If
	
	If CurrentText = "" Then Exit Sub
	  If Buf.Exists(CurrentText) = False Then ' [+]MetaEditor (������ ������ �������� �� ����� �.� Exists ���������� ���� � �� ��������, ���� ���� � ���������)
		If UseMet = "1" Then
			akey = InputBox("������� �����", "����� ������")
		If akey = "" Then Exit Sub
		Else
			akey = CurrentText
	    End If
		
		If Buf.Exists(akey) then msgbox "����� �� ���������.",1,"Clipboard" : Exit Sub '[+]MetaEditor
		
	    Buf.Add replace(akey,vbCrlF,vbverticaltab), replace(CurrentText,vbCrlF,vbverticaltab)
	  Else
		MsgBox "���������� ����� ��� ���� � ������!!!",0,"Clipboard"
	  End If
	CurrentText = ""
Else 	
	CommonScripts.SendCommand(57634)		
End If	
End Sub

Sub PasteText() '�������� �����
If IsEnable = "0" Then'�� ������������ ������
        CommonScripts.SendCommand(57637)
Exit Sub
End If
If Buf.Count = 0 Then
 MsgBox "����� ������."
 Exit Sub
End If
	If CommonScripts.IsTextWindow() = true   Then
		Set doc = CommonScripts.GetTextDoc(0)
		If doc Is Nothing Then Exit Sub
		LineStart = doc.SelStartLine
		ColStrart = doc.SelStartCol
		LineEnd = doc.SelEndLine
		ColEnd = doc.SelEndCol
		If Buf.Count = 1 then
			keys = Buf.Keys
			pasttext = Buf.Item(Keys(0))
		Else	
			pasttext = CommonScripts.SelectValue(Buf,"��� ��������?",True)
		End If
		pasttext = replace(pasttext,vbverticaltab,vbCrlF)
		if pasttext = "" then Exit Sub 
		doc.Range(LineStart, ColStrart, LineEnd, ColEnd) = pasttext
		doc.movecaret LineStart, ColStrart + Len(pasttext)
	Else
		CommonScripts.SendCommand(57637)		
	End If	
End Sub

Sub EmptyBufer() '������� �� ������ ��������� ��������
If IsEnable = 0 Then Exit Sub
If Buf.Count = 0 Then 
	MsgBox "����� ������."
	Exit Sub
End If
keys = Buf.Keys
List=""
for i = 0 to Buf.Count-1 
	List = List & Keys(i) & vbCrLf
next
del = CommonScripts.SelectValue(List,"��� ������� �� ������?",,True)
If del = "" Then Exit Sub 
Buf.Remove(del)
End Sub

Sub EmptyAll() ' ��������� �������� ����� 
If IsEnable = 0 Then Exit Sub
	Buf.RemoveAll
End Sub

'���-�� � ������� ToggleCase �� MetaEditor'a
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

Sub ToggleCase()
	set doc = Commonscripts.GetTextDoc(0)
	if doc is nothing then exit sub
	SelText=Trim(doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol))
	
	if Doc.SelStartCol = Doc.SelEndCol then 
		Borders = GetCurrentWordBorders (Doc.Range(Doc.SelStartLine), Doc.SelStartCol, " .,;:|#=+-*/%?<>\()[]{}!~@$^&'""" & vbTab)
		SelText = doc.Range(Doc.SelStartLine, Borders(0)-1, Doc.SelEndLine,  Borders(0) + Borders(1) - 1)
		if SelText = "" then exit sub
	end if	            
	
	LastMode = not LastMode

	If SelText=LCase(SelText) Then
		If LastMode then Mode = 3 Else Mode = 2
		LastMode = not LastMode				
	ElseIf SelText=UCase(SelText) Then
		If LastMode then Mode = 1 Else Mode = 3
		LastMode = not LastMode		
	Else 
		If LastMode then Mode = 2 Else Mode = 1
		LastMode = not LastMode
	End If

	If Mode=1 Then
		if Doc.SelStartCol = Doc.SelEndCol then
  		    Doc.Range(Doc.SelStartLine , Borders(0)-1 , Doc.SelEndLine , Borders(0) + Borders(1) - 1) = LCase(SelText) 
		else
			Doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol)=LCase(SelText)
		end if
	ElseIf Mode=2 Then
		if Doc.SelStartCol = Doc.SelEndCol then
  		    Doc.Range(Doc.SelStartLine , Borders(0)-1 , Doc.SelEndLine , Borders(0) + Borders(1) - 1) = UCase(SelText) 
		else
			doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol)=UCase(SelText)
		end if	
	ElseIf Mode=3 Then
		NewText = split(SelText," ") 
	    newstring_ = ""
		If Ubound(NewText)>0 Then
			For Each wd in NewText
				wd1 = UCase(Left(wd,1)) & LCase(Right(wd,Len(wd)-1))
				newstring_ = newstring_  & wd1 & Chr(32)
			Next
			Doc.Range(Doc.SelStartLine ,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = newstring_
		Else
			ProperCase = UCase(Left(SelText,1)) & LCase(Right(SelText,Len(SelText)-1))
			if Doc.SelStartCol = Doc.SelEndCol then
				Doc.Range(Doc.SelStartLine , Borders(0)-1 , Doc.SelEndLine , Borders(0) + Borders(1) - 1) = ProperCase
			else
				doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = ProperCase
			end If	
		End If	
	End If           
End Sub

Sub Translit()
	' (c) ���� 2004 trdm@fromru.com
	' (MetaEditor - ����������� ��� ��������� � ����� ����������� � ����� ��������� ���������)
	' ��������� ������ "Gecnm ,eltn RHENJ" � "����� ����� �����"
	Set doc = CommonScripts.GetTextDoc(0)
	if doc is Nothing then Exit Sub
    
	SelText = doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol)
  
	Lat = "qwertyuiop[]asdfghjkl;'zxcvbnm,./QWERTYUIOP{}ASDFGHJKL:" & Chr(34) &"ZXCVBNM<>?"
	Rus = "��������������������������������.��������������������������������,"
	
	'Lat = "qwertyuiop[]asdfghjkl;'zxcvbnm,..QWERTYUIOP{}ASDFGHJKL;'ZXCVBNM,./<>!@#$%^&*()_+|�;:?1234567890 """ & vbTab & vbCrLf
	'Rus = "��������������������������������.��������������������������������.<>!@#$%^&*()_+|�;:?1234567890 """ & vbTab & vbCrLf
	
	ResulTtext = "" 
	for i = 1 To Len(SelText)
		Char = Mid(SelText, i,1)
		If Instr(Lat,Char) > 0 then 
			Direction = "LatToRus"
		ElseIf Instr(Rus,Char) > 0 then 
			Direction = "RusToLat"                
		Else Direction="unknown"
		End If	
		
		If Direction = "unknown" then 
			ResulTtext = ResulTtext + Char	
		Else	
			If Direction = "RusToLat" then 
				pos = Instr(Rus, Char)
				ResulTtext = ResulTtext + Mid(Lat,pos,1)
			end if
			If Direction = "LatToRus" then 
				pos = Instr(Lat, Char)
				ResulTtext = ResulTtext + Mid(Rus,pos,1)
			end if
		End if
	next
  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = ResulTtext
End Sub

Sub Settings()
 Sett(2)
End Sub

Sub Sett(p) '���������� �����������
rk = Reg.ScriptRootKey(SelfScript.Name)
  If Reg.KeyExists(rk&"IsEnable") = False Then
	Reg.Param(rk,"IsEnable") = "0"
	IsEnable="0" 
  Else IsEnable = Reg.Param(rk,"IsEnable")
  End If
  If Reg.KeyExists(rk&"UseMet") = False Then
	Reg.Param(rk,"UseMet") = "0"
	UseMet="0" 
  Else UseMet = Reg.Param(rk,"UseMet")
  End If
  If p=1 Then '�������� �� Init(param)
	Exit Sub
  End If
  If IsEnable = "1" Then
	pometka = "|c"
  Else
	pometka = ""
  End If
  vv = "����� �������" & pometka & vbCrLf
  If UseMet = "1" Then
	pometka = "|c"
  Else
	pometka = ""
  End If	
  vv = vv & "������������ �����" & pometka
  MsgBox "��������!!! ���� ����� ������ ������ ������, �� ��� ������ ����� �����!!!",0,"!!! Clipboard !!!"
  asd = svc.SelectValue(vv,"���������",True)
  If instr(asd,"����� �������") = 0 Then
	IsEnable="0"
  Else IsEnable="1"
  End If
  If instr(asd,"������������ �����") = 0 Then
	UseMet="0"
  Else UseMet="1"
  End If
  Reg.Param(rk,"IsEnable") = IsEnable
  Reg.Param(rk,"UseMet") = UseMet
End Sub             

Sub Init(param)
  ff = 0
  Set svc = CreateObject("Svcsvc.Service")
  Set CommonScripts = CreateObject("OpenConf.CommonServices")
  CommonScripts.SetConfig(Configurator)
  Set Reg=CreateObject("OpenConf.Registry")
  Reg.SetConfig(Configurator)
  Set Buf = CreateObject("Scripting.Dictionary")
  Sett(1)
End Sub

Init 0