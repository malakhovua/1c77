'dots.vbs
'
'������: $Revision: 1.13 $
'
'Intellisence ��� OpenConf
'	(� trdm)	������ �. �.
'��������:
'	����� ������� aka artbear
'
'���������� ������, ��� ������� �� ����� ��������� �����������
Const NotIntellisenceExtensions = ".vbs.js.wsc.prm"

const cnstNameTypeSZ = "��������������"
const cnstNameTypeTZ = "���������������"
const cnstNameTypeTabl = "�������"
const BagKatch		= False '����� �������?
const cnstRExWORD = "a-zA-Z�-��-�0-9_"
const KakVibiraem = 2
FindFirstInFindInStrEx = False
'1 - �������� � ��� svcsvc.dll SelectValue
'2 - �������� � ��� svcsvc.dll FilterValue
'3 - �������� � ��� SelectValue.dll FilterValue

'������� � ������� ���������� ������� ������� ��������� �
'����� �� ��������, ��� ��������� ���������
Set vk_dict_CreateObject = CreateObject("Scripting.Dictionary")

'������ �� ��� ������� ����� ���������� ����������  0 = ��� �� ������� 1 = ������� � �� �������� � ���  (��� ���������� ����������)
GlobalModuleParse = 0
' ��� �� ������ ������ ���������� ���������� � �� ����� � �������  ���������� = ���.���# ���������� = ���.���#
Set GlobalVariableType = CreateObject("Scripting.Dictionary")
Set GlobalVariableNoCase = CreateObject("Scripting.Dictionary")
Set FSO = CreateObject("Scripting.FileSystemObject")

Set UniMetodsAndRekv = CreateObject("Scripting.Dictionary")

Dim GlobalModule '������ TheModule ��� ����������� ������
Dim LocalModule	 '������ TheModule ��� ������ �������� �������
Dim glDoc		 '������ ������� ����� ������������� (������ � ���������� ������������� ����� ������� �� ������ ������ � ������)
glDoc = ""
Dim glStartLileEnd '�������� �� ������� ��������� ����� ������

ttextModule = ""
ttextModuleGM = ""
ttextModuleGM_PriNachRabSys = ""
Dim GRegExp, regEx, Match, Matches, glRWord

' �������� ���������, � �� �� ��� ����� ����������, ��� ������ ������...
Class AWord					'������������� �����
	Public RW				'��������� �� ���� ������ �������� �����

	Public RWAll			'�������������� ����� ��� ����������� ���� �� ������� ����� ����� �������

	Public RWOld			'�������������� ������, ������� ���� ����������, �������� � ������...
	Public RWFullStr		'������ ������ ������, � ������������� ������, ������ �������� ��� ����,
							'���-�� ������� ����� � �������...
	Public TypeVid			'�������������� ���/��� �����
	Public TypeVid2			'�������������� ���/��� ����� ��������: ��������+���������������
	Public AddWord			'������ ����, �� ������� ��������������� �� ��������� � �������� ��� �����������...

	Public AddTypeVid		'������ ����� ��� ���� ����.

	Public BTObj			'������ ������, ��� ����� �������������
	Public BTMeth			'����� ��� ������ �������, ��� ������ �������������..
	Public BTNumberParams	'����� ���������, � ������� ������������� �����...
	Public IsBetweenToken	'������������� ������� ��������� � �������
	Public IsIcvalVid		'���� ��������� ����� � ������ ������
	Public IsNeedBrasket	'����� ������ (��� ����� ����������.���() = "!!")

	Public AsCreateObj		'����� ������ (��� ����� ����������.���() = "!!")


	Public RecvOfForms		'���������� - �������� �����

	Public LineText			'������ ������ � ������� �������
	Public LineStartToCaret	'������ ������ � ������ � �� �������
	Public LineCaretToEnd	'������ ������ � ������� � �� �����

	Public FindAtribute		'(������ ��������) ���� ������ ������, ���� �� ����������, ������ ��� ����, ��� ��� ����...

	Public TempArray		'��������� ������� ��� ������ �����������.....

	Public doc_SelStartLine		'������ ���������, ��� �������� ��������
	Public doc_CurParseLine		'������� �������������� ������


	Public doc_IsObject			'������ ��������� ��������/����������
	Public doc_ObjectType		'��� ������� ���������
	Public doc_ObjectVid		'��� ������� ���������
	Public doc_ObjectHaveAtrib	'������ ��������� ����� ��������

	'�������...
	Public ArrProcFuncFromTheWord	'��������� � ������� � ������� �������� ���� ������������� �������...
								'������ ����������������������������������
	Public ArrProcFuncNumbParam	'����� ��������� ��� ��������� ��� �������.
								'���� �������� ���������� �������� - ����� ����� "-1"


	Private Sub Class_Initialize
		doc_IsObject = false
		doc_ObjectHaveAtrib = false

		RW  = ""
		RWAll	= ""
		RWOld = ""
		RWFullStr = ""
		AddWord = "" :	AddTypeVid = "" : BTObj = "" : BTMeth = "" :
		TypeVid = ""
		TypeVid2=""
		BTNumberParams = ""
		IsBetweenToken = false
		IsIcvalVid = false
		IsNeedBrasket = false
		FindAtribute = true
		AsCreateObj = false

		RecvOfForms	= 0

		LineText			= ""
		LineStartToCaret	= ""
		LineCaretToEnd		= ""
		doc_IsObject		= false
		doc_ObjectType		= ""
		doc_ObjectVid		= ""
		doc_ObjectHaveAtrib	= ""
		doc_SelStartLine	= 0
		doc_CurParseLine	= 0

		ProcFuncFromTheWord = ""
		ArrProcFuncFromTheWord	= ""
		ArrProcFuncNumbParam	= ""

	End Sub

	Sub GetDocInfo(doc)
		DocName = doc.Name
		Arr = Split(DocName,".")
		doc_SelStartLine	= doc.SelStartLine
		doc_CurParseLine	= 0
		if UBound(Arr)>0 Then
			if instr(1,"/��������/����������/�����/���������/������/",lcase(Arr(0)))>0 Then
				doc_IsObject = true
				doc_ObjectType = Arr(0)
				doc_ObjectVid = Arr(1)
				if instr(1,"/��������/����������/������/",lcase(Arr(0)))>0 Then
					doc_ObjectHaveAtrib = true
				End IF
			End IF
		End IF
	End Sub
End Class

'����� ��� ����������
Class TheVariable
	Public V_Vid	'��� ����������
	Public V_Type	'��� ����������
	Public Words	'������ �������������� ����������
	Public WordsCnt	'������ ������� ����

	Function Verify() ' ��������� �������������..
		Verify = false
		If (Len(V_Vid)>0) And (Len(V_Type)>0) Then
			Verify = true
		End If
	End Function

	Function VerifyAgr() ' ��������� ������������� �� ���������� �������
		VerifyAgr = false
		If (Len(V_Type)>0) Then
			if InStr(1,"/��������/����������/�������/������������/�����/���������/","/" & V_Type & "/") Then
				VerifyAgr = true
			End If
		End If
	End Function

	Sub ExtractDef(ttext)
		ttext = lcase(ttext)
		if Len(ttext)> 0 Then
			if InStr(1, ttext, ".")	> 0 Then
				Words = Split(ttext,".")
			Else
				Words = Array(ttext)
			End If
			WordsCnt = UBound(Words)
			if WordsCnt<>-1 Then
				V_Type = Words(0)
			End If
			if WordsCnt>0 Then
				V_Vid = Words(1)
			End If
		End If

	End Sub
	Private Sub Class_Initialize
		V_Vid = ""
		V_Type = ""
	End Sub

End Class


Dim StartScanLineText	'������ �������� ����������� �����....
Dim GlobalStr			'��� ���������� ����������� �����....

Dim SoobshitType '�������� ���/��� ������ ����, ����� ������ ������ ����������/�������......


'*********************************************************************
' �������� ��� WordOfCaret(), ��������� ����������� ����� �����������
' ������������ ��������� ��� �������� ��� ��������....
Sub TypeWordOfCaret()
	SoobshitType = True :	WordOfCaret()	:	SoobshitType = False
End Sub

'������������� �� ������ ��� ������ ������� � �������...
Dim IndexModule
IndexModule = True
'IndexModule = False ' - ��� �������, ����� �� ����� ������ ��� ����� ������ �����������...

'������������� �� ������ ��� ������ ������� � �������...
Dim NeedCheckModule
NeedCheckModule = True
'NeedCheckModule = False '��������� ��� �����������...

Dim SuccessfulWork
Function IsSuccessfulWork()
	WordOfCaret()
	IsSuccessfulWork = SuccessfulWork
End Function

'*********************************************************************
'�������� ���������
Sub WordOfCaret()
	Set RWord = new AWord
	Set var = new TheVariable
	glDoc = "" : glStartLileEnd = -1
	SuccessfulWork = false
	Doc = ""
	If Not CheckWindow(Doc) Then
		Exit Sub
	End IF

	IF NeedCheckModule Then
		SyntaxCheckModule()
	End IF
'stop
	Rword.GetDocInfo(doc)
	'���������� ��������� RWord, ���� �� ������������ ��� ������ ��� ����������������...
	Set glRWord = RWord

	'��� ��� ��������, �� ��������� �� �� � ����������� � ��������� ���� ������������
	'���������� ������ ���� �������� � ����������� ....
	Configurator_OnActivateWindow Windows.ActiveWnd,true

	StartScanLineText	= doc.SelStartLine
	IF (GlobalModuleParse = 0) And IndexModule Then
		ParseGlobalModule(doc)
	End If
	'�������� ���� ����� ��� ����� � �������
	RWord.RW = GetWordFromCaret(RWord.RWAll,RWord.RWFullStr,RWord.LineCaretToEnd,  RWord)
	If Not (RWord.IsIcvalVid) Then
		RWord.IsBetweenToken = BetweenToken(doc.SelStartCol,RWord.BTObj, RWord.BTMeth, RWord.BTNumberParams, doc,RWord)
		if RWord.IsBetweenToken Then
			If Len(RWord.BTMeth)>0 Then
				RWord.IsBetweenToken = CheckIsBetweenToken(RWord.BTMeth, RWord.BTNumberParams, RWord, doc)
				If RWord.IsBetweenToken Then
					If Len(RWord.BTObj)<>0 Then
						RWord.RW		= RWord.BTObj
						RWord.RWAll	= ""
					End If
				End If
			ElseIf (Len(RWord.BTMeth)=0) And (Len(RWord.RW)=0) Then
				Exit Sub
			End If
		End If
	End If
	if (RWord.RW = "?") Then
		Exit Sub
	End If
	LocalModule	= ""
	If Doc.Name <> "���������� ������" Then
		Set LocalModule = New TheModule
		LocalModule.SetDoc(Doc)
		LocalModule.InitializeModule(1)
	Else
		If Not IsEmpty(GlobalModule) Then
			Set LocalModule = GlobalModule
		End If
	End If

	GlobalStr = Doc.SelStartLine - 1
	RWord.RW = Replace(RWord.RW,vbCr,".")
	LastTipeVid = ""
	aroff = 0

	'������� ������������� (����� ��������� ������� �������� ����������� ������...)
	TypeVid1 = GetTypeVid(RWord.RW,LastTipeVid, Doc, RWord.RecvOfForms,RWord.RW)
	TypeVid  = TypeVid1
	IF Len(TypeVid1)>0 And Len(RWord.AddWord)>0 Then
		TypeVid1 = GetTypeVid(RWord.AddWord,TypeVid1, Doc, RWord.RecvOfForms,RWord.RW)
	End If
	If SimpleType(TypeVid1) Then
		Status RWord.RWAll & "->" & TypeVid1
		Exit Sub
	End If
	TypeVid  = TypeVid1
	NameOfTableFromTable = ""
	'��� ��� - OR (lCase(TypeVid) = "��������") OR (lCase(TypeVid) = "����������")
	'��� ����� ������ ������� �� ����� � ����� ��� ���������� ��� ��������
	If (len(TypeVid) = 0) OR (lCase(TypeVid) = "��������") OR (lCase(TypeVid) = "����������")  Then	'������ �� ���������� :(
		' ��������� ��������� �����
		if RWord.RW <> "!" Then
			TypeVid = GetTypeFromTextRegExps(RWord.RW,RWord.AddWord)
			If Len(TypeVid) = 0 And Len(glRWord.TypeVid)<>0 Then
				TypeVid = glRWord.TypeVid
			End If
			If ((lCase(TypeVid1) = "��������") OR (lCase(TypeVid1) = "����������")) And (Len(TypeVid) = 0) Then 'And InStr(1,TypeVid,"#")>0 Then
				TypeVid = TypeVid1
			End If
		End If
	End If

	IF InStr(1,UC("/���������/��������/����������/������/������������/�����/���������/����������/�������/"),UC(RWord.RW))>0 And (Len(TypeVid)=0) Then
		TypeVid = RWord.RW
		RWord.FindAtribute = False
		if Len(RWord.AddWord)> 0 Then
			TypeVid = GetTypeVid(RWord.AddWord,RWord.RW, Doc, RWord.RecvOfForms,RWord.RW)
		End If
	End If
	IF InStr(1,TypeVid,"+")>0 Then
		aroff = Split(TypeVid,"+")
		RWOrd.TypeVid  = aroff(0)
		RWOrd.TypeVid2  = aroff(1)
		TypeVid = aroff(0)
		TypeVid1 = aroff(1)
	End If

	ttext = ""
	strRekv =""
	if (Len(ttext) = 0) And (Not RWord.IsBetweenToken) then
		IF Len(TypeVid) = 0 Then
			If ((INStr(1,UCase(Doc.Name),UCase("CWBModuleDoc::"))>0) And (INStr(1,UCase(Doc.Name),UCase("ert"))>0)) _
			OR (INStr(1,UCase(Doc.Name),UCase("�����."))>0) OR (INStr(1,UCase(Doc.Name),UCase("���������."))>0) Then '������� �����
				if RWord.RW = "!" Then
					'TypeVid = "�����"
				End IF
			End IF
		Else
			ttt = Split(TypeVid, ".")
			if UBound(ttt) = 2 Then
				TypeVid = GetTypeVid(ttt(2),ttt(0)&"."&ttt(1), Doc, RWord.RecvOfForms,ttt(2))
			End IF
		End IF
	End IF
	IF SoobshitType Then
		IF Len(TypeVid)>0 Then
			MsgBox TypeVid
		Else
			MsgBox "�� ��������� :("
		End If
		Exit Sub
	End If
	if ((instr(1,"/������/�����/����/", lcase(TypeVid))>0) Or (Len(TypeVid) = 0)) And (Not RWord.IsBetweenToken) And (Len(RWord.RW) = 0) Then
		If Len(TypeVid) = 0 Then exit sub
		MsgBox TypeVid
		exit sub
	End If

	If (Len(TypeVid)<>0) And (RWord.FindAtribute) And Not (Ucase(RWord.BTMeth) = UCase("����")) And Not RWord.IsIcvalVid Then
		IF Instr(1,TypeVid,".")>0 Then
			Arr = Split(TypeVid,".")
			if UBound(Arr) = 1 Then
				SMetodami = 1
				if (RWord.IsBetweenToken) OR (rword.IsIcvalVid) Then
					SMetodami = 0
				End IF
				strRekv = GetStringRekvizitovFromObj(Arr(0), Arr(1),SMetodami,RWord.RecvOfForms, RWord.BTMeth,	RWord.BTNumberParams, RWord)
				if Len(strRekv)<>0 Then
					ttext = SelectFrom(strRekv, Caption)
					IF Len(ttext) = 0 Then
						SuccessfulWork = true
						Exit Sub
					End If
				End If
			End If
		Else
			IF RWord.IsBetweenToken Then
				SMethodami = 0
			else
				SMethodami = 1
			End If
			TypeVid1 = TypeVid
			IF InStr(1,TypeVid, "#����������������������#")>0 Then
				SMethodami = 0
				ArrTtext2 = Split(TypeVid,"#")
				RWord.RW = ArrTtext2(0)
				TypeVid = "���������������"
			End If

			strRekv = GetMethodAndRekvEx(TypeVid, RWord.RW, SMethodami,RWord.RecvOfForms, RWord.IsBetweenToken, NameOfTableFromTable)
			If Lcase(TypeVid) = "��������" Then
				strRekv0 = GetORD()
				if lcase(RWord.RW) = "��������" Then
					Set MetaDoc  = MetaData.TaskDef.Childs(CStr(TypeVid))
					For tt = 0 To MetaDoc.Count-1
						strRekv0 = strRekv0 &  vbCrLf & MetaDoc(tt).Name
					Next
				End If
				strRekv = strRekv0 & vbCrLf & strRekv
			End If
			if len(strRekv) <> 0 Then
				IF InStr(1,TypeVid1, "#����������������������#")>0 Then
					Arr = Split(strRekv,vbCrLf)
					If UBound(Arr)<>-1 Then
						strRekv = ""
						For cntWords = 0 To UBound(Arr)
							if cntWords = 0 Then
								strRekv = " = """ & Arr(cntWords) & """"
							Else
								strRekv = strRekv & vbCrLf & " = """ & Arr(cntWords) & """"
							End If
						Next
					End If
				End If
				ttext = SelectFrom(strRekv, Caption)
				ttext = PodrabotkaVibora(strRekv, ttext)
				if len(ttext) = 0 Then
					SuccessfulWork = true
					Exit Sub
				End If
			End If
		End If
	End If
	IF RWOrd.FindAtribute And Not SoobshitType And Not RWord.IsIcvalVid Then
		IF (UCase(RWord.RW) = UCase("���������")) Then
			ttext = GetConstantExA()
		End If

		IF RWord.IsBetweenToken And Len(ttext) = 0 Then
			IF (Ucase(RWord.BTMeth) = UCase("����")) Or (Ucase(RWord.BTMeth) = UCase("Color")) Then
				ttext = CStr(VibColor)
			End If
		End If


		if (Len(ttext) = 0) And (Right(RWord.RWAll,1)<>".") And (Not RWord.IsBetweenToken) And (UCASE(TypeVid) = "�����") then
			ArrDocPresent = Split(Doc.Name,".")
			IF UBound(ArrDocPresent)>=2 Then
				IF (UCase(ArrDocPresent(0)) = "��������") OR (UCase(ArrDocPresent(0)) = "����������") Then
					strRekv = GetStringRekvizitovFromObj(ArrDocPresent(0), ArrDocPresent(1),1,1, RWord.BTMeth,	RWord.BTNumberParams, RWord)
					if Len(strRekv)<>0 Then
						tType = ""
						If EtoFormaDokumenta(tType) Then
							if tType = "��������" Then
								strRekv = strRekv & vbCrLf &  "* ������ ����������"
								strRekv = strRekv & vbCrLf &  "* ������ ���������� �����"
								strRekv = strRekv & vbCrLf &  "* ������ ���������� �����"
								strRekv = strRekv & vbCrLf &  "* ������ ���������� �������"
								strRekv = strRekv & vbCrLf &  "* ������ �� ����������"
								strRekv = strRekv & vbCrLf &  "* �������"
							Else
								strRekv = strRekv & vbCrLf &  "* ������ ����������"
								strRekv = strRekv & vbCrLf &  "* ������ �� ����������"
								strRekv = strRekv & vbCrLf &  "* �������"
							End If
						End If
						ttext = SelectFrom(strRekv, Caption)
						ttext = PodrabotkaVibora(strRekv, ttext)
						if Len(ttext) = 0 Then
							SuccessfulWork = true
							Exit Sub
						End If
					End If
				End IF
			End IF
		End IF

		If RWord.IsBetweenToken And (Len(ttext) = 0) Then '2004.12.03 ������� ������� �� �������������
			RWord.RWAll  = ""
			IF UC(RWord.BTMeth) = UC("�������������") Then
				If InStr(1, UC(RWord.BTObj),UC("����������."))>0 Then
					ArrBTObj = Split(RWord.BTObj,".")
					If UBound(ArrBTObj) = 1 Then
						Set MetaRef  = MetaData.TaskDef.Childs(CStr(ArrBTObj(0)))(CStr(ArrBTObj(1)))
						Set CHI = MetaRef.Childs
						IF CHI.Item(1).Count<> 0 Then
							strRekv = ""
							For iii = 0 To CHI.Item(1).Count - 1
								iF Len(strRekv) = 0 Then
									strRekv = CHI.Item(1).Item(iii).Name
								Else
									strRekv = strRekv & vbCrLf & CHI.Item(1).Item(iii).Name
								End IF
							Next
							if Len(strRekv)<>0 Then
								ttext = SelectFrom(strRekv, Caption)
								if Len(ttext) = 0 Then
									SuccessfulWork = true
									Exit Sub
								End If
							End If
						End IF
					End IF
				End IF
			ElseIF (UC(RWord.BTMeth) = UC("������������")) Or (UC(RWord.BTMeth) = UC("AccessRight")) Then
				If UCase(RWord.TypeVid) = UCase("#�������������#") Then
					'������� ��������� �����
				End IF
			ElseIF (UC(RWord.BTMeth) = UC("������������")) Or (UC(RWord.BTMeth) = UC("AccessRight")) Then
				IF RWord.BTNumberParams = 1 Then
					strRekv = GetStrAccessRight()
					ttext = SelectFrom(strRekv, Caption)
					if Len(ttext) = 0 Then
						SuccessfulWork = true
						Exit Sub
					End If
				End IF
			ElseIF UC(RWord.BTMeth) = UC("����") Or UC(RWord.BTMeth) = UC("Total") Then
				if (RWord.RW = "!") And (RWord.doc_ObjectType = "��������") Then
'Stop
					TekObj = "" : strRekv = ""
					If Not GetTekObj(TekObj) Then
						Exit Sub
					End If
					Set ch = TekObj.Childs("����������������������")
					For cnmr = 0 To ch.Count-1
						Set MetaRekv = ch(cnmr)
						if MetaRekv.Props(9) = "1" Then	 'MetaRekv.Props.Name(9)	"�������������"	String
							AddToString strRekv, MetaRekv.Name,vbCrLf
						End IF
					Next
					ttext = SelectFrom(strRekv, Caption)
					if Len(ttext) = 0 Then
						SuccessfulWork = true
						Exit Sub
					End If
				End IF
			End IF
		End IF
	End If
	If Len(ttext) = 0 Then '{ trdm 2005 01 22
		If InStr(1,UC("/���������/��������/����������/������/������������/�����/���������/����������/�������/"),UC("/"&RWord.RW&"/")) Or (RWord.IsIcvalVid)   Then
		'If (UC(RWord.RW) =  UC("�����")) OR (UC(RWord.RW) =  UC("���������")) Or (RWord.IsIcvalVid)   Then
			ttext = RWord.RW
			if RWord.IsIcvalVid Then
				if Len(TypeVid)>0 Then
					var.ExtractDef(TypeVid)
					if var.VerifyAgr() Then
						ttext = var.V_Type
					End If
				End If
			End If
			if Len(ttext)>0 Then
				var.ExtractDef(ttext)
				if var.VerifyAgr() Then
					strRekv = ""
					if var.VerifyAgr() And (Lcase(var.V_Type) = "������������") Then
						'Set MetaDoc  = MetaData.TaskDef.Childs(CStr(var.V_Type))(CStr(var.V_Vid))
						strRekv = GetStringRekvizitovFromObj(var.V_Type, var.V_Vid,0,0, "",	"", RWord)
					Else
						Set MetaDoc  = MetaData.TaskDef.Childs(CStr(ttext))
						For tt = 0 To MetaDoc.Count-1
							IF Len(strRekv) = 0 Then
								strRekv = MetaDoc(tt).Name
							Else
								strRekv = strRekv &  vbCrLf & MetaDoc(tt).Name
							End If
						Next
					End If
					if Len(strRekv)>0 Then
						ttext = SelectFrom(strRekv, Caption)
						if Len(ttext) = 0 Then
							SuccessfulWork = true
							Exit Sub
						End If
						' ���� � ��� ttext - ��������� �������� ������������, ���� ��������� ����� ���:
						if var.VerifyAgr() And (Lcase(var.V_Type) = "������������") Then
							ttext = "������������." & var.V_Vid &"."& ttext
						End If
					End If
				Else
					ttext = ""
				End If
			End If
		End If
	End If ' }trdm 2005 01 22

	if Len(ttext) = 0  then
		StrObjekta = ""
		If RWord.AsCreateObj Then
			StrObjekta = Get_file_vk_dict(TypeVid)
		End IF
		for fff = 0 To 4
			if fff = 0 Then strRekv = LoadMethodFromFile(StrObjekta)
			if fff = 1 Then strRekv = LoadMethodFromFile(RWord.RWFullStr)
			if fff = 2 Then strRekv = LoadMethodFromFile(RWord.RW)
			if fff = 3 Then strRekv = LoadMethodFromFile(RWord.BTMeth)
			if fff = 4 Then strRekv = LoadMethodFromFile(TypeVid)

			If Not IsEmpty(strRekv) Then
				strRekv = Trim(strRekv)
				strRekv = Replace(strRekv,"(f)","(<?>)")
				ttext = SelectFrom(strRekv, Caption)
				if len(ttext) = 0 Then
					SuccessfulWork = true
					Exit Sub
				Else
					TypeVid = ""
				End If
				Exit for
			End If
		Next
	End If

	if Len(ttext)<>0 then
		DotIsEpsent = False
		IcvalPrezent = False
		IF InStr(1, ttext, "=") > 0 Then '���� � ��� ��������� ���� "= ������������.�����������.������" ����� ����� �� �����...
			IcvalPrezent = True
		End IF
		RWord.LineCaretToEnd = DelKomment(RWord.LineCaretToEnd)
		RWord.LineCaretToEnd = Trim(RWord.LineCaretToEnd)
		if Len(RWord.LineCaretToEnd)>0 Then
			If InStr(1,ttext,";")>0 Then
				ttext = Replace(ttext,";","")
			End If
		End If
		IF (Not RWord.IsBetweenToken) And (RWord.RW <> "!") And Not IcvalPrezent And not RWord.IsIcvalVid Then
			if Len(RWord.RWAll)>0 Then '�������� �����
				If Not (Mid(RWord.RWOld,Len(RWord.RWOld))=".") Then
					If (RWord.RWAll<>"!") And (Not RWord.IsBetweenToken) Then
						DotIsEpsent = True
						ttext = "." & ttext
					end if
				end if
			end if
			if Not DotIsEpsent Then
				LinefromText = doc.range(doc.SelStartLine)
				RWord.RWFullStr = mid(LinefromText,1,doc.SelStartCol)
				RWord.RWFullStr = Trim(RWord.RWFullStr)
				ls = Len(RWord.RWOld)
				if Mid(RWord.RWOld,ls,1) <> "." Then
					ttext = "." & ttext
				end if
			end if
		end if
		if RWord.IsIcvalVid Then
			'���� � ��� ������������ ������������ � �� ���������� �������� ����� ���� "������������.���" ��
			' �������� ������������ ���������� ���������� ����� (������������.���) ����
			if (StrCountOccur(ttext, ".") = 2) And (StrCountOccur(lcase(ttext),"������������.") = 1) Then
				ExtraktGlobalVariableName ttext
			end if
			if RWord.IsNeedBrasket Then
				If Not ((Mid(trim(RWord.LineStartToCaret),Len(trim(RWord.LineStartToCaret)),1) = """") OR (Mid(trim(RWord.LineCaretToEnd),1,1) = """")) Then
					ttext = """" & ttext & """"
				end if
			end if
		end if

		Pos = InStr(ttext,"<?>")
		if InStr(ttext,"<?>") = 0 Then
			Pos = len(ttext)
		else
			Pos = Pos - 1
		end if
		ttext = Replace(ttext, "<?>", "")
		doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine,doc.SelEndCol ) = ttext
		'doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelStartLine ,doc.SelStartCol) = ttext
		doc.MoveCaret doc.SelStartLine, doc.SelStartCol+Pos
		Status "dots"
		SuccessfulWork = true
	End If
	Status RWord.RWAll & "->" & TypeVid

end sub

Function GetNextMetaDataChild(TheWord, ReturnWord)
end Function
Function GetMetaObjectChild(TheWord, ReturnWord)
end Function


Private Function PodrabotkaVibora(strRekvIn, ttext)
	PodrabotkaVibora = ttext
	strRekvIn = "#" & Replace(strRekvIn, vbCrLf,"#") & "#"
	strRekv = ""
	SformirovaliZdes = False
	If InStr(1, ttext, "*")>0 Then
		PodrabotkaVibora = ""
		If InStr(1, ttext, "������")>0 Then
			TekObj = ""
			If Not GetTekObj(TekObj) Then
				Exit function
			End If
			If Instr(1,TekObj.FullName,"��������.") Then
				For tt = 0 To 2
					if tt=0 Then Set ch = MetaData.TaskDef.Childs("����������������������")
					if tt=1 Then Set ch = TekObj.Childs("�������������")
					if tt=2 Then Set ch = TekObj.Childs("����������������������")
					For cnmr = 0 To ch.Count-1
						Set MetaRekv = ch(cnmr)
						If ttext = "* ������ �� ����������" Then
							strRekvIn = ReplaceEx(strRekvIn, Array("#" & MetaRekv.Name & "#" ,"#", "#�������#" ,"#", "#��������#" ,"#", "#��������#" ,"#"))
						else
							AddedText = ""
							If ttext = "* ������ ����������" Then
								AddedText = MetaRekv.Name
							ElseIf (ttext = "* ������ ���������� �����") And (tt = 0) Then
								AddedText = MetaRekv.Name
							ElseIf (ttext = "* ������ ���������� �����") And (tt = 1) Then
								AddedText = MetaRekv.Name
							ElseIf (ttext = "* ������ ���������� �������") And (tt = 2) Then
								AddedText = MetaRekv.Name
							End IF
							If Len(AddedText)>0 Then
								If InStr(1,strRekvIn,"#" & AddedText & "#") > 0 Then
									AddToString strRekv, MetaRekv.Name,vbCrLf
								End IF
							End IF
							SformirovaliZdes = True
						End IF
					Next
				Next
			ElseIf Instr(1,TekObj.FullName,"����������.") Then
				Set ch = TekObj.Childs("��������")
				For cnmr = 0 To ch.Count-1
					Set MetaRekv = ch(cnmr)
					If ttext = "* ������ �� ����������" Then
						strRekvIn = Replace(strRekvIn, "#" & MetaRekv.Name & "#" ,"#")
					else
						AddedText = ""
						If ttext = "* ������ ����������" Then
							AddedText = MetaRekv.Name
						End IF
						If Len(AddedText)>0 Then
							If InStr(1,strRekvIn,"#" & AddedText & "#") > 0 Then
								AddToString strRekv, MetaRekv.Name,vbCrLf
							End IF
						End IF
						SformirovaliZdes = True
					End IF
				Next
			End If
		ElseIf ttext = "* �������" Then
			SformirovaliZdes = True
			strRekv = LoadSlovar()
			strRekv = Trim(strRekv)
		End If
	Else
		Exit Function
	End If
	if Not SformirovaliZdes Then
		strRekvIn = ReplaceEx(strRekvIn, Array("#* ������ ����������#" ,"#", "#* ������ ���������� �����#" ,"#", "#* ������ ���������� �����#" ,"#", "#* ������ ���������� �������#" ,"#", "#* ������ �� ����������#" ,"#"))
		Do While Instr(1,strRekvIn,"##")>0
			strRekvIn = Replace(strRekvIn, "##" ,"#")
		Loop
		strRekvIn = Mid(strRekvIn,2,Len(strRekvIn)-2)
		ArrRekv = Split(strRekvIn,"#")
		strRekvIn = ""
		If UBound(ArrRekv)<> -1 Then
			For tt = 0 To UBound(ArrRekv)
				If Not (InStr(1,ArrRekv(tt),"(")>0) Then
					AddToString strRekvIn, ArrRekv(tt),vbCrLf
				End If
			Next
		End If
	Else
		strRekvIn = strRekv
	End If
	If Len(strRekvIn)>0 Then
		ttext = SelectFrom(strRekvIn, Caption)
		if len(ttext) > 0 Then
			PodrabotkaVibora = ttext
		End If
	End If

end Function



Private Function AddToString(Str, AddStr,Spl)
	If Len(Str)>0 Then
		Str = Str & Spl & AddStr
	Else
		Str = AddStr
	End IF
	AddToString = Str
end Function

Private Function AddToStringUni(Str, AddStr,Spl)
	AddToStringUni = False
	If Len(Str)>0 Then
		ArrUniStr = Split(Str,Spl)
		For qq = 0 To UBound(ArrUniStr)
			if Trim(Lcase(ArrUniStr(qq))) = Trim(Lcase(AddStr)) Then
				Exit Function
			End IF
		Next
		Str = Str & Spl & AddStr
	Else
		Str = AddStr
	End IF
	AddToStringUni = True
end Function



Private Function GetTekObj(TekObj)
	GetTekObj = False
	Doc = ""
	If CheckWindow(Doc) Then
		NameDoc = Doc.Name
		ArrName = Split(NameDoc,".")
		If UBound(ArrName) > 1 Then
			if (ArrName(0) = "��������") OR (ArrName(0) = "����������") Then
				Set TekObj = MetaData.TaskDef.Childs(CStr(ArrName(0)))(CStr(ArrName(1)))
				GetTekObj = True
			End If
		End If
	End If
End Function



Private Sub ParseGlobalModule(doc)
	GlobalModuleParse = 1 ' ���� ��������� � ��� ��������� ������� �� �����..

	GlobalVariableType.RemoveAll()
	GlobalVariableNoCase.RemoveAll()
	Set DocGM = Documents("���������� ������")
	ttextGM = DocGM.text
	if Doc.Name <> "���������� ������" Then
		Set GlobalModule = New TheModule
		GlobalModule.SetDoc(DocGM)
		GlobalModule.InitializeModule(1)
	End IF

	IF Len(ttextGM) = 0 Then
		Exit Sub
	End IF
	ttextObjiavlVARIABLE = ""
	For iii = 0 To DocGM.LineCount 	'��������� ����������
		ttext = DocGM.Range(iii)
		ttext = Trim(ttext)
		IF (InStr(1,lcase(ttext),"���������") = 1) Or (InStr(1,lcase(ttext),"�������") = 1) or (InStr(1,lcase(ttext),"function") = 1 )or (InStr(1,lcase(ttext),"procedure") = 1 )  Then
			Exit For
		End IF
		ttext = DelKomment(ttext)
		If Len(ttext)>0 Then
			ttextObjiavlVARIABLE = ttextObjiavlVARIABLE &" "& ttext
		End IF
	Next
	ttextObjiavlVARIABLE = lcase(ttextObjiavlVARIABLE)
	ttextObjiavlVARIABLE = ReplaceEx(ttextObjiavlVARIABLE, Array(" �������","", vbTab&"�������","", "����� ","", "�����"&vbTab,"", ";",",", " ","", vbCrLf,"", vbCr,"", vbTab,"" ) )
	ttextObjiavlVARIABLE = trim(ttextObjiavlVARIABLE)
	ttextObjiavlVARIABLE = mid(ttextObjiavlVARIABLE,1,len(ttextObjiavlVARIABLE)-1)
	ArrGlobVARIABLE = split(ttextObjiavlVARIABLE,",")
	For iii = 0 To UBound(ArrGlobVARIABLE)
		IF Not GlobalVariableType.Exists(ArrGlobVARIABLE(iii)) Then
			GlobalVariableType.Add ArrGlobVARIABLE(iii),""
			GlobalVariableNoCase.Add ArrGlobVARIABLE(iii),""
		End IF
	Next
	' ������ ���������� ����������, ������� ��� � ��� ���
	' �������� ����������� �� ����� ������ ������ ������������� ����������

	ttextInicializeGM = ""
	For iii = DocGM.LineCount  To 0 Step -1
		ttext = DocGM.Range(iii)
		ttext = Trim(ttext)
		IF (InStr(1,lcase(ttext),"��������������") = 1) Or (InStr(1,lcase(ttext),"������������") = 1)  Then
			Exit For
		End IF
		ttext = DelKomment(ttext)
		if Len(ttext)>0 then
			ttextInicializeGM = ttext & vbCrLf & ttextInicializeGM
		End IF
	Next
	ttextModuleGM = ttextInicializeGM
	ttextInicializeGM = ReplaceEx(ttextInicializeGM, Array(vbTab,"",  " ",""))
	ParseTextFromGlobalVariable ttextInicializeGM, GlobalVariableType
	'�������� ����� � ��������� ���������

	'������ ������ ��������� ����������������������
	Patern = "(���������)+[\s]+(����������������������)+[\s]*[\(]+[^~]+(��������������)+"
	ttext = FindInStrEx(Patern,ttextGM)
	ttextInicializeGM = ""
	ArrGlobVARIABLE = split(ttext,vbCrLf)
	For iii = 0 To UBound(ArrGlobVARIABLE)
		tt = ArrGlobVARIABLE(iii)
		tt = DelKomment(tt)
		if Len(ttext)>0 then
			ttextInicializeGM = ttextInicializeGM & vbCrLf & tt
		End IF
		IF (InStr(1,lcase(tt),"��������������") = 1) Then
			Exit For
		End IF
	Next
	ttextModuleGM_PriNachRabSys = ttextInicializeGM
	ParseTextFromGlobalVariable ttextInicializeGM, GlobalVariableType
end sub

Private Sub ParseTextFromGlobalVariable(textForParse, Dict)
	ArrStrforParse = split(textForParse,vbCrLf)
	' ���� �����: ��� ������� ����� ������ (���������, ������� ��� ������ ������������� ������),
	' �� ��� ������������� �� ������� ������������� ����������.
	' ���� ��������� ������ ����������, ��������� � ���� ������, ���� ������ ����������
	' ���������� �������� � ��� ��� ��������� ����������, ���� ���� ���������� ����������,
	' ��������� ��� ��� � ���� ���� �������.
	Set PNRSDict = CreateObject("Scripting.Dictionary") '��������� �������
	for iii = 0 To ubound(ArrStrforParse)
		VarName = ""
		ttext = ArrStrforParse(iii)
		ttext = Replace(ttext,vbTab,"")

		if Len(ttext) > 1 Then
			posIcval = InStr(1,ttext,"=")
			if posIcval<>0 Then
				VarName = Mid(ttext,1,posIcval)
				VarName = Replace(VarName,"=","")
				VarNameNoCase = VarName
				VarName = lcase(VarName)
				Expr = mid(ttext,posIcval+1)
				Expr = lcase(Expr)
				' �������������� Expr, ���� � ��� ������ ���������� ������, ����� ��� ��� ���� :(
				'� ��� - ������ ��������� ��� �������������, ���� ������, ���� �����
				If Not Mid(Expr,1,1) = """" Then
					IF InStr(1,VarName,".") = 0 Then
						If (InStr(1,Expr,"���������.") = 1) Then
							Expr = Replace(Expr,"���������.","")
							Expr = Replace(Expr,";","")
							On Error Resume Next
							set Obj = MetaData.TaskDef.Childs("���������")(CStr(Expr))
							If len(Obj.Props(3))>0 Then '���, ���� ��� ��������� ��������������� ����, ����� ��� �����
								If InStr(1,"����������.������������.��������",lcase(Obj.Props(3)))>0 Then
									OstavimVariable = True
									If Obj.Props(4)> 0 Then
										If Dict.Exists(VarName) Then
											Dict.Remove(VarName)
											Dict.Add VarName, Obj.Props(3) &"."&Obj.Props(4)
											if GlobalVariableNoCase.Exists(VarName) Then
												GlobalVariableNoCase.Remove(VarName)
											End IF
											GlobalVariableNoCase.Add VarName, VarNameNoCase
											'Message "VarName6=" & VarName & " Expr = " & Obj.Props(3) &"."&Obj.Props(4)
										End IF
									End IF
								End IF
							End IF
						ElseIf  (InStr(1,Expr,"������������.") = 1) Then
							Expr = Replace(Expr,"������������.","")
							Expr = Replace(Expr,";","")
							If InStr(1,Expr,".") = 0 Then '������� ������ ����� ������������.�������
								'On Error Resume Next
								IF Dict.Exists(VarName) Then
									Dict.Remove(VarName)
									Dict.Add VarName, "������������." & Expr
									'Message "VarName5=" & VarName & " Expr = " & "������������." & Expr
									if GlobalVariableNoCase.Exists(VarName) Then
										GlobalVariableNoCase.Remove(VarName)
									End IF
									GlobalVariableNoCase.Add VarName, VarNameNoCase
								Else
									If PNRSDict.Exists(VarName) Then
										PNRSDict.Remove(VarName)
									End IF
									PNRSDict.Add VarName, Expr
								End IF
							End IF
						ElseIf (InStr(1,Expr,"�������������(") = 1) Or (InStr(1,Expr,"createobject(.") = 1) Then
							Expr = ReplaceEx(Expr, Array("�������������(","",  "createobject(","", ";","", ")","", """","" ))
							ArrExprWord = Split(Expr,".")
							IF UBound(ArrExprWord) <> -1 Then
								If InStr(1,"���������������.��������.��������������.������������������.����������.��������.�������.addin.",lcase(ArrExprWord(0)))>0 Then
									if Dict.Exists(VarName) Then
										Dict.Remove(VarName)
										Dict.Add VarName, Expr
										if GlobalVariableNoCase.Exists(VarName) Then
											GlobalVariableNoCase.Remove(VarName)
										End IF
										GlobalVariableNoCase.Add VarName, VarNameNoCase
									Else
										If PNRSDict.Exists(VarName) Then
											PNRSDict.Remove(VarName)
										End IF
										PNRSDict.Add VarName, Expr
									End IF
								End IF
							End IF
						Else
							If GlobalVariableType.Exists(VarName) Then
								TypeVarName = GlobalVariableType.Item(VarName)
								TypeVarName = trim(TypeVarName)
								Expr = Replace(Expr," ","")
								ArrWordExpr = Split(Expr,".")
								IF UBound(ArrWordExpr)=1 Then '��������� ���� "���.��������������();"
									If (InStr(1,ArrWordExpr(1),"��������������();")=1) Then
										If Dict.Exists(ArrWordExpr(0)) Then
											ittext = Dict.Item(ArrWordExpr(0))
											if (InStr(1,ittext,"����������")=1) And (InStr(1,ittext,".")>0) And (InStr(1,ittext,"+")=0) Then
												If GlobalVariableType.Exists(VarName) Then
													GlobalVariableType.Remove(VarName)
													GlobalVariableType.Add VarName, ittext
													If GlobalVariableNoCase.Exists(VarName) Then
														GlobalVariableNoCase.Remove(VarName)
													End IF
													GlobalVariableNoCase.Add VarName, VarNameNoCase
													'Message "VarName4=" & VarName & " ittext = " & ittext
												End IF
											End IF
										ElseIf PNRSDict.Exists(ArrWordExpr(0)) Then
											ittext = PNRSDict.Item(ArrWordExpr(0))
											if (InStr(1,ittext,"����������")=1) And (InStr(1,ittext,".")>0) And (InStr(1,ittext,"+")=0) Then
												If GlobalVariableType.Exists(VarName) Then
													GlobalVariableType.Remove(VarName)
													GlobalVariableType.Add VarName, ittext
													If GlobalVariableNoCase.Exists(VarName) Then
														GlobalVariableNoCase.Remove(VarName)
													End IF
													GlobalVariableNoCase.Add VarName, VarNameNoCase
													'Message "VarName4=" & VarName & " ittext = " & ittext
												End IF
											End IF

										End IF
									End IF
								End IF
							End IF
						End IF
					End IF
				End IF
			End IF
		End IF
	 next
end sub

Private Function DelKomment(ttext)
	PosKomment = InStr(1,ttext, "//")
	IF PosKomment>0 Then
		If PosKomment <> 1 Then
			ttext = Mid(ttext,1,PosKomment-1)
			ttext = Trim(ttext)
		Else
			ttext = ""
		End IF
	End IF
	DelKomment = ttext
end Function

Private Function CheckIsBetweenToken(BTMeth, BTNumberParams, RWord, doc)
	IsBetweenToken = False
	ArrMethods = Array(UCase("��������"), UCase("GroupBy"),_
		UCase("�����������"), UCase("Group"), _
		UCase("����"), UCase("Color"), _
		UCase("���������"), UCase("IsItAGroup"),_
		UCase("�������������"), UCase("CreateObject"),_
		UCase("����������������"), UCase("UseLayer"),_
		UCase("�������������"), UCase("PutSection"),_
		UCase("����������������"), UCase("ColumnVisibility"),_
		UCase("������������������"), UCase("AttachSection"), _
		UCase("������������"), UCase("AccessRight")) '2004.12.30
	MethodsForBetweenToken = "#"& Join(ArrMethods,"#") & "#"
	Patern = "[#]+(" & UC(BTMeth) & ")+[#]+"
	FindStr = FindInStrEx(Patern,MethodsForBetweenToken)
	IF Len(FindStr)>0 Then
		IsBetweenToken = True
	End If
	IF (BTNumberParams = 1) Then
		IF	UC(BTMeth) = UC("������������") OR _
			UC(BTMeth) = UC("NewColumn") OR _
			UC(BTMeth) = UC("����") OR _
			UC(BTMeth) = UC("Total") OR _
			UC(BTMeth) = UC("��������������") OR _
			UC(BTMeth) = UC("Activate") OR _
			UC(BTMeth) = UC("��������������������������") OR _
			UC(BTMeth) = UC("SetColumnParameters") OR _
			UC(BTMeth) = UC("�������������������������") OR _
			UC(BTMeth) = UC("SetFilterValue") OR _
			UC(BTMeth) = UC("�����������") OR _
			UC(BTMeth) = UC("Sort") OR _
			UC(BTMeth) = UC("���������������") OR _
			UC(BTMeth) = UC("GetAttrib") OR _
			UC(BTMeth) = UC("�����������������") OR _
			UC(BTMeth) = UC("SetAttrib") OR _
			UC(BTMeth) = UC("��������������������������") OR _
			UC(BTMeth) = UC("SelectItemsByAttribute") OR _
			UC(BTMeth) = UC("�������������������") OR _
			UC(BTMeth) = UC("ShowImages") OR _
			UC(BTMeth) = UC("����������������") OR _
			UC(BTMeth) = UC("ColumnVisibility") Then
			IsBetweenToken = True
			IF UC(BTMeth) = UC("��������������") OR UC(BTMeth) = UC("Activate")  Then
				glRWord.TypeVid = "#�������������#"
			End If
		End If
	ElseIF (BTNumberParams = 2) Then
		IF UC(BTMeth) = UC("�������������") OR UC(BTMeth) = UC("OpenPermanentChoice") OR _
			UC(BTMeth) = UC("�������") OR UC(BTMeth) = UC("Choose") OR _
			UC(BTMeth) = UC("�����������������������") OR UC(BTMeth) = UC("UnloadTable")  Then
			IsBetweenToken = True
			Set tvar = new TheVariable
			tvar.ExtractDef(Doc.Name)
			if tvar.VerifyAgr() And (UC(BTMeth) = UC("�����������������������") OR UC(BTMeth) = UC("UnloadTable")) Then
				if lcase(tvar.V_Type) = lcase("��������") Then
					'RWord.BTObj = tvar.V_Type & "." & tvar.V_Vid
					RWord.RW = "���������������();"
				else
					IsBetweenToken = false
				End If
			else
				IsBetweenToken = True
			End If
		End If
	ElseIF (BTNumberParams = 3) Then
		IF UC(BTMeth) = UC("�������������") OR UC(BTMeth) = UC("FindValue") Then
			IsBetweenToken = True
		End If
	ElseIF (BTNumberParams = 4) Then
		IF UC(BTMeth) = UC("���������") OR UC(BTMeth) = UC("Fill") OR _
			UC(BTMeth) = UC("���������") OR UC(BTMeth) = UC("Unload")  Then
			IsBetweenToken = True
		End If
	End If
	if (Ucase(RWord.BTMeth) = UCase("�������������")) OR (Ucase(RWord.BTMeth) = UCase("CreateObject")) Then
		If Len(RWord.RW)>0 Then
			IsBetweenToken = true
			RWord.IsIcvalVid = FALSE
			RWord.FindAtribute = FALSE
		End If
	End If
	CheckIsBetweenToken = IsBetweenToken
End Function


Private Function GetMethodAndRekvEx(TypeVid, ResultWord, SMethodami,AsRecvOfForms, IsBetweenToken, NameOfTableFromTable)
	ttext = ""
	'���� � ��� ��� ��� ����� ������, �������, ���������������, ��������������
	If UCase(TypeVid) = "������" Then
		strRekv = GetVariableAndFunctionZapros(ResultWord,0)
		if SMethodami = 1 Then
			strRekv = strRekv & vbCrLf & GetZaprosMet()
		End If
	ElseIf UCase(TypeVid) = "���������������" Then
		strRekv = GetColumnsFromTZ(ResultWord,0)
		if SMethodami = 1 Then
			strRekv	= strRekv & vbCrLf & GetMethodsOfTablicaZnacheniy(0)
		End If
	ElseIf UCase(TypeVid) = UCase("����������������������") Then
		strRekv = strRekv & vbCrLf & GetMethodsOfFormPodbor()
	ElseIf UCase(TypeVid) = "�����" Then
		strRekv = GetTableRecvFromForms(0,IsBetweenToken)
		if SMethodami = 1 Then
			strRekv = strRekv & vbCrLf & GetMethodsOfForm()
		End If
		tType = ""
		If EtoFormaDokumenta(tType) Then
			if tType = "��������" Then
				strRekv = strRekv & vbCrLf &  "* ������ ����������"
				strRekv = strRekv & vbCrLf &  "* ������ ���������� �����"
				strRekv = strRekv & vbCrLf &  "* ������ ���������� �����"
				strRekv = strRekv & vbCrLf &  "* ������ ���������� �������"
				strRekv = strRekv & vbCrLf &  "* ������ �� ����������"
			Else
				strRekv = strRekv & vbCrLf &  "* ������ ����������"
				strRekv = strRekv & vbCrLf &  "* ������ �� ����������"
			End If
		End If
	ElseIf UCase(TypeVid) = UCase("#�������������#") Then
		strRekv = GetMethodsOfFormRecv()
	ElseIf UCase(TypeVid) = UCase("��������������") Then
		strRekv = GetMethodsOfListBox(AsRecvOfForms)
	ElseIf UCase(TypeVid) = UCase("��������") Then
		strRekv = GetDocMet(0)
		strRekv = SortStringForList(strRekv, vbCrLf)
	ElseIf UCase(TypeVid) = UCase("��������") Then
		strRekv = GetListKindVariable("�������")
		if SMethodami = 1 Then
			strRekv = strRekv & vbCrLf & GetMethodsRegisters()
		End If
	ElseIf UCase(TypeVid) = UCase("�������") Then
		strRekv = GetListKindVariable("�������")
	ElseIf UCase(TypeVid) = UCase("�����") Then
		strRekv = GetMethodsText("")
	ElseIf UCase(TypeVid) = UCase("�������") Then
		if IsBetweenToken Then
			strRekv = GetListMethodAndRekvtablica(ResultWord,0,NameOfTableFromTable)
		else
			strRekv = GetListMethodAndRekvtablica(ResultWord,SMethodami,NameOfTableFromTable)
		End If
	ElseIf UCase(TypeVid) = UCase("������������") Then
		strRekv = GetEnum("")
	'Else
	'	strRekv = LoadMethodFromFile(TypeVid)
	End If
	GetMethodAndRekvEx = strRekv
End Function



Private Function GetEnum(par)
	GetEnum = "" : tGetEnum = ""
	Set Enums = MetaData.TaskDef.Childs(CStr("������������"))
	For tt = 0 To Enums.Count - 1
		Set mdo = Enums(tt)
		AddToString tGetEnum, mdo.Name,vbCrLf
	Next
	GetEnum = tGetEnum
End Function


Private Function EtoFormaDokumenta(tType)
	EtoFormaDokumenta = False
	tdoc = ""
	if CheckWindow(tdoc) Then
		ArrN = Split(tdoc.Name,".")
		If UBound(ArrN)<> -1 Then
			If (ArrN(0) = "��������") Or (ArrN(0) = "����������") Then
				EtoFormaDokumenta = True
				tType = ArrN(0)
			End IF
		End IF
	End IF
End Function

Private Function LoadMethodFromFile(FileNm)
	LoadMethodFromFile = ""
	FileName = BinDir + "\Config\Intell\" + FileNm + ".ints"
	FileExists = FSO.FileExists(FileName)
	if FileExists = False then
		FileName = BinDir + "\Config\Intell\1�++\" + FileNm + ".ints"
		FileExists = FSO.FileExists(FileName)
	end if

	if FileExists = True then
		Set Fl = FSO.GetFile(FileName)
		Set FileStream = Fl.OpenAsTextStream()
		if FileStream.AtEndOfStream = true then  Exit Function

		AllMeth = FileStream.ReadAll()
		AllMethods = Split(AllMeth, vbCrLf)
		for i = 0 to UBound(AllMethods)
			Methods = Methods + vbCrLf + picMeth + Mid(AllMethods(i), 6)
		next
	end if
	LoadMethodFromFile = Methods
End Function


'���������� �� ��������� ���� "���������.��������("�����, ���������","");"
'������, � �.�.: "���������"
'�����, � �.�.: "��������"
'� ����� ������ � ������� ����� ������ � (BTNumberParams)
'������ ����� ������������ ��� � ���� ������ ������ :)
Private Function BetweenToken(SelStartCol,BTObj, BTMeth, BTNumberParams, doc, RWord)
	LinefromText = doc.range(doc.SelStartLine)
	BetweenToken = False
	RWord.BTNumberParams = 0
	'Patern = "[" & cnstRExWORD & "]+[\s]*[.]*[\s]*[" & cnstRExWORD & "]+[\s]*[\(]+[\s]*[" & cnstRExWORD & ",.\s\(\)""\+]*[\s]*[\)]+[\s]*[;]*"
	Patern = "[" & cnstRExWORD & "]+[\s]*[.]*[\s]*[" & cnstRExWORD & "]+[\s]*[\(]+.*[\)]+"
	ttext = FindInStrEx(Patern,LinefromText)
	ttextAll = ttext
	'����� ��������� �� ��������
	if Len(ttext)>0 Then
		'����������� � ������ ���������� � �������
		Patern = "[" & cnstRExWORD & "]+[\s]*[.]*[\s]*[" & cnstRExWORD & "]+[\s]*"
		FindFirstInFindInStrEx = True
		VariableAndMethod = FindInStrEx(Patern,ttext)
		VariableAndMethod = Replace(VariableAndMethod, vbTab, " ")
		Do While InStr(1,VariableAndMethod,"  ")>0
			VariableAndMethod = Replace(VariableAndMethod,"  "," ")
		Loop
		VariableAndMethod = Trim(VariableAndMethod)
		'������� ������ �� ��� � ������
		Patern = "[" & cnstRExWORD & "]{1,1}[\s]*[.]+[\s]*[" & cnstRExWORD & "]{1,1}[\s]*"
		pttext = FindInStrEx(Patern,VariableAndMethod)
		tArr = split(pttext,vbCrLf)
		If UBound(tArr) <> -1 Then
			pttext = ""
			For ee = 0 To UBound(tArr)
				VariableAndMethod = Replace(VariableAndMethod,tArr(ee), Left(tArr(ee),1) & "."& Right(tArr(ee),1))
			Next
		End IF
		tArr = split(VariableAndMethod," ")
		If UBound(tArr)>0 Then
			VariableAndMethod = tArr(UBound(tArr))
		End IF


		'Patern2 - ��� ���������� ������������ ������� � ������ � ������
		Patern2 = "[" & cnstRExWORD & "]+[\s]*[.]*[\s]*[" & cnstRExWORD & "]+[\s]*[\(]+"
		VariableAndMethodSoSkobkoy = FindInStrEx(Patern2,ttext)
		FindFirstInFindInStrEx = False
		if Len(VariableAndMethodSoSkobkoy) = 0 Then
			Exit Function
		End IF
		If Len(VariableAndMethod)>0 Then
			VariableAndMethod = Replace(VariableAndMethod," ","")
			IF InStr(1,VariableAndMethod,".")>0 Then '������ ���� � ��� � ���������� � �����
				tArr = Split(VariableAndMethod,".")
				RWord.BTObj = Trim(tArr(0))
				RWord.BTMeth = Trim(tArr(1))
			else
				RWord.BTMeth = Trim(VariableAndMethod)
			End If
		Else
			Exit Function
		End If
		'�������� ��������� �� ����� ������������ ������ � ������� (VariableAndMethod),
		'����� �� ������ �� ��� � �� �� �������
		KursorNaMeste = True
		StrokaGdeIchem = Left(LinefromText,doc.SelStartCol)
		If InStr(1,StrokaGdeIchem,VariableAndMethodSoSkobkoy) = 0 Then
			'������ ������ ����� ����� �� ������ ��� �������
			KursorNaMeste = False
		End IF
		Patern = "[\s]*[\)]+[\s]*[;]*"
		ZaSkobkoy = FindInStrEx(Patern,LinefromText)

		'�������� �� ��������� �� ������ �� �������
		StrokaGdeIchem = Mid(LinefromText,doc.SelStartCol)
		If InStr(1,StrokaGdeIchem,ZaSkobkoy) = 0 Then
			'������ ������ ����� ����� �� ������ ��� �������
			KursorNaMeste = False
		End IF
		if Not KursorNaMeste Then
			BetweenToken	= False
			RWord.BTObj			= ""
			RWord.BTMeth			= ""
			Exit Function
		End IF
		'���� ��� � ��� � �������
		Patern = "[\(]+.*[\)]+"

		InBrasket = FindInStrEx(Patern,LinefromText)

		If Len(InBrasket)>0 Then
			'���� �����, ������� ������ � ��������
			Patern = "[\(]+.*[\)]+"
			InBrasket = FindInStrEx(Patern,InBrasket)
			If Len(InBrasket)>0 Then
				Params = GetParams(InBrasket)
				'MsgBox Params
				ArrParams = Split(Params, vbCrLf)
				NumberMeth = 1
				if UBound(ArrParams)>0 Then
					'������ ����� �������� ��� ��������� (����� ������� � ������)
					'��������� � ����� ������� ����� � ������ � ���� ������ �� �������,
					'���� �� ������� ���������, ����� ������ �� ��������� � ��� :)
					StrokaGdeIShem = Mid(LinefromText,doc.SelStartCol+1)
					StrokaKotoruIShem = ""
					For tt = UBound(ArrParams) To 0 Step -1
						StrokaKotoruIShem = ArrParams(tt) & StrokaKotoruIShem
						If InStr(1,StrokaGdeIShem,StrokaKotoruIShem) = 0 Then
							NumberMeth = tt + 1
							Exit For
						End If
					Next
					'MsgBox BTNumberParams
				End If
				RWord.BTNumberParams = NumberMeth
			End If
		End If
	End If
	If (RWord.BTNumberParams<>0) And (RWord.BTMeth<>"") Then
		BetweenToken = True
		'2004.12.03 - ��������� ������� �� �������������
		If UBound(ArrParams)>0 Then
			IF (RWord.BTNumberParams = 2) And (UC(RWord.BTMeth) = UC("�������������")) Then
				If (InStr(1,uC(ArrParams(0)),"����������.")>0) And (InStr(1,uC(ArrParams(0)),"+")=0) Then
					RWord.BTObj = Replace(ArrParams(0),"""","")
				End If
			End If
		End If
	Else
		BetweenToken = False
	End If
	RWord.IsBetweenToken = BetweenToken

End Function

'� ��� ��� ��������� ��������� ��� ������� ���� ��������� � �������� �\� ������� �� ��� �� �������.
'��������� �� �������� � ��������
'�����: ���� ����������� � ����������� �����������, � ����� ���� �������,
'���� ����������� ������� ������� �� �������� ������������ �������, � ��������� ���-��
'������ ������� ��������. � ���� ����� ������ ���������� ��������� ���������
'OpenBraskets = "" -  � ��� �������� ������ ���� ""; ()
Private Function GetParams(InBrasket)
	Params = ""
	InBrasket2 = ""
	For tt = 2 To Len(InBrasket)-1 '����� ��� ����������� ������
		Char = Mid(InBrasket,tt,1)
		IF (Char = """") Or (Char = "(") Or (Char = ")") Then
			IF Len(Braskets)>0 Then '���� �������� �����������
				IF (Char = """") Then
					IF Right(Braskets,1) = Char Then '����������� ������� �����
						Braskets = Left(Braskets,Len(Braskets)-1)
					End If
				ElseIF (Char = "(") Or (Char = ")") Then
					IF (Char = "(") Then
						Braskets = Braskets & Char
					ElseIF Char = ")" Then
						IF Right(Braskets,1) =  "(" Then '����������� ������� �����
							Braskets = Left(Braskets,Len(Braskets)-1)
						End If
					End If
				End If
			Else
				Braskets = Braskets & Char
			End If
		End If
		'InBrasket2 = InBrasket2 & Mid(InBrasket,tt,1)
		if (Len(Braskets) = 0) And (Char = ",") Then
			IF Len(Params)>0 Then
				Params = Params & vbCrLf & InBrasket2
			Else
				Params = InBrasket2
			End IF
			InBrasket2 = ""
			InBrasket2 = InBrasket2 & Mid(InBrasket,tt,1)
		else
			InBrasket2 = InBrasket2 & Mid(InBrasket,tt,1)
		End IF
	Next
	if Len(InBrasket2)>0 Then
		IF Len(Params)>0 Then
			Params = Params & vbCrLf & InBrasket2
		Else
			Params = InBrasket2
		End IF
		InBrasket2 = ""
	End IF
	if Char = "," Then
		Params = Params & vbCrLf
	End IF
	GetParams = Params
End Function



'Private Function GetWordFromCaret()
Function GetWordFromCaret(OldResultWord,ResultWordFull,ResultWordEnd,  RWord)
	'FindInStrEx(Patern,ResultWordFull)
	ResultWordFull = ""
	doc = ""
	if Not CheckWindow(doc) then Exit Function

	BadWord = ""
	ResultWord = ""
	LinefromText = doc.range(doc.SelStartLine)
	RWord.LineText = LinefromText
	SelStartCol = doc.SelStartCol
	If Len(LinefromText)<>0 Then
		If Len(LinefromText)>(SelStartCol+1) then
			RWord.LineCaretToEnd = Mid(LinefromText,SelStartCol+1)
		End IF
		RWord.LineStartToCaret = Mid(LinefromText,1,SelStartCol)
	End IF
	tmpStrOfBadSimbols = " " & vbTab & vbCr & vbCrLf &"""()-=?,;+<>|/*"
	if (SelStartCol = 0) Then
		For i = (doc.SelStartCol+1) To Len(LinefromText)
			Char = Mid(LinefromText, i,1)
			if InStr(1,tmpStrOfBadSimbols,Char)>0  Then
				Exit For
			End If
			ResultWord = ResultWord + Char
		next
	else
		'��������� ������ ����� ��� ���������� �� ������ ints
		ResultWordFull = mid(LinefromText,1,doc.SelStartCol)
		ResultWordFull = Trim(ResultWordFull)
		Do While InStr(1,ResultWordFull,"  ")>0
			ResultWordFull = Replace(ResultWordFull,"  "," ")
		Loop
		ResultWordEnd = Replace(LinefromText, ResultWordFull, "")

		for ttt = 0 to 3
			If ttt = 0 Then
				Patern = "[\s]*[\.]+[\s]*"
			ElseIf ttt = 1 Then	Patern = "[\s]*[\(]+[\s]*[""]+"
			ElseIf ttt = 2 Then	Patern = "[\s]*[""]+[\s]*[\)]+"
			ElseIf ttt = 3 Then	Patern = "[\(]+[\s]*[""]+[" & cnstRExWORD & "]+[""]+[\s]*[\)]+"
			End If
			ttextZameni = FindInStrEx(Patern,ResultWordFull)
			If ttextZameni<>"" Then
				ArrWordZameni = Split(ttextZameni,vbCrLf)
				If UBound(ArrWordZameni)<> -1 Then
					For iii = 0 to UBound(ArrWordZameni)
						itZameni = ArrWordZameni(iii)
						If Len(itZameni)<>0 Then
							If ttt = 0 Then
								ResultWordFull = Replace(ResultWordFull,itZameni,".")
							ElseIf ttt = 1 Then	ResultWordFull = Replace(ResultWordFull,itZameni,"(""")
							ElseIf ttt = 2 Then	ResultWordFull = Replace(ResultWordFull,itZameni,""")")
							ElseIf ttt = 3 Then	ResultWordFull = Replace(ResultWordFull,itZameni,"")
							End If
						End If
					Next
				End If
			End If
		Next
		LinefromText = ResultWordFull
		LenLinefromText = Len(LinefromText)+1
		If Len(ResultWordFull)<>0 Then
			ArrWord = Split(ResultWordFull)
			ResultWordFull = ""
			If Ubound(ArrWord)<> -1 Then
				ResultWordFull = ArrWord(Ubound(ArrWord))
			End If
		End If
		ResultWordFull = Replace(ResultWordFull,".","")
		'LenLinefromText
		'For i = doc.SelStartCol+1 To Len(LinefromText)
		For i = LenLinefromText To Len(LinefromText)
			Char = Mid(LinefromText, i,1)
			if InStr(1,tmpStrOfBadSimbols,Char)>0  Then
				Exit For
			End If
			BadWord = BadWord + Char
		next
		if (LenLinefromText <> 0) Then
			'for i = (doc.SelStartCol-1) To 1 Step -1
			for i = (LenLinefromText-1) To 1 Step -1
				Char = Mid(LinefromText,i,1)
				if  InStr(1,tmpStrOfBadSimbols,Char)>0  Then 'tmpStrOfBadSimbols
					Exit For
				End If
				ResultWord = Char + ResultWord
			next
		End If
	end if
	if (len(ResultWord) = 0 ) Then
		ResultWord = "!"
	ElseIF WordInArray(GetListOFStatments(), ResultWord) Then
		ResultWord = "!"
	end if
	OldResultWord = ResultWord
	If Len(ResultWord) > 0 Then	'����������� ����� ���� ��� ���� � ����� �����
		If Mid(ResultWord,Len(ResultWord),1) = "." Then
			ResultWord = Mid(ResultWord,1,Len(ResultWord)-1)
		End IF
	End IF
	If InStr(1,ResultWord,".")>0 Then
		DimResultWordAdd = split(ResultWord,".")
	Else
		DimResultWordAdd = Array(ResultWord)
	End IF
	if InStr(ResultWord,".")>0 Then
	RWord.RWAll = ResultWord
	' ���� � ��� ���� ����� � �����, ����� ���� ������ ����� �����, � ���������� ���������� � ������
		tempWords = ""
		Dim tempDimResultWord
		tempDimResultWord = split(ResultWord, ".")
		for i = 0 to UBound(tempDimResultWord)
			if Len(tempDimResultWord(i))>0 Then
				if i = 0 Then
					ResultWord = tempDimResultWord(i)
				else
					if (Len(RWord.AddWord)>0) then
						RWord.AddWord = RWord.AddWord & "." & tempDimResultWord(i)
					else
						RWord.AddWord = tempDimResultWord(i)
					end if
				end if
			end if
		Next
		' �=������ � ������� DimResultWordAdd ����� ��� ������ �� "���������������()" ��� ��������� ���������.
	end if
	GetWordFromCaret = ResultWord
	if Len(BadWord)>0 Then
		GetWordFromCaret = "?"
	end if
	GetWordFromCaret = Trim(GetWordFromCaret)
	RWord.RWFullStr = ResultWordFull
	RWOrd.RWOld = OldResultWord

	'��������� ���������� �� �� ���-�� � ���-��
	'��������� �� �������� � ����.����.���() = [<!>|"<!>"]
	ttext = FindInStrEx("[" & cnstRExWORD &"\.]+\s*\.+\s*(���|Kind)+\(+\s*\)+\s*\=+\s*[""""]*",RWord.LineStartToCaret)
	if Len(ttext) = 0 Then
		RWord.IsNeedBrasket = false
		ttext2 = FindInStrEx("[" & cnstRExWORD &"\.]+\s*=+\s*[""""]*",RWord.LineStartToCaret)
		ttext = ttext2
	else
		RWord.IsNeedBrasket = true
	End IF
	' FindInStrEx - ����� ������� ��������� ��������, ����������� vbCrLf, ��� ���� ����� ���������
	if InStr(1,ttext, vbCrLf)>0 Then
		tempDimResultWord = Split(ttext,vbCrLf)
		ttext = tempDimResultWord(UBound(tempDimResultWord))
	End IF
	'������ ������ ������ � ������ ��������� "= [<!>|"<!>"]"
	If InStr(1,StrReverse(RWord.LineStartToCaret),StrReverse(ttext)) = 1 Then
		' ������� "���()"
		ttext2 = FindInStrEx("\s*\.+\s*(���|Kind)+\(+\s*\)+\s*",ttext)
		if Len(ttext2)>0 Then
			ttext = Replace(ttext,ttext2,"")
		End IF

		if Len(ttext)>0 Then
			ttext = FindInStrEx("[" & cnstRExWORD &"\.]+\s*",ttext)
			if Len(ttext) = 0 Then
				RWord.IsNeedBrasket = false
				ttext = FindInStrEx("[" & cnstRExWORD &"\.]+",ttext2)
			End IF
			ttext = Trim(ttext)
			if Len(ttext)>0 Then
				ttext = ReplaceEx(ttext, Array(" ","", vbTab, "", "=",""))
				if (Right(ttext,1) = ".") Then
					ttext = Left(ttext,len(ttext)-1)
				End IF
				if (InStr(1,ttext,".") > 0) Then
					tempDimResultWord = Split(ttext,".")
					if UBound(tempDimResultWord)<> -1 Then
						ttext = tempDimResultWord(0)
					End IF
					if UBound(tempDimResultWord)>0 Then
						RWord.AddWord = ""
						For e=1 To UBound(tempDimResultWord)
							if Len(RWord.AddWord)>0 Then
								RWord.AddWord = RWord.AddWord & "." & tempDimResultWord(e)
							Else
								RWord.AddWord = tempDimResultWord(e)
							End IF
						Next
					End IF
				End IF
				RWord.IsIcvalVid = true
				RWOrd.BTObj = ttext
				GetWordFromCaret = RWOrd.BTObj
			End IF
		End IF
	End IF
	'��������� �� �������� � ��������������(��������) = [<!>|"<!>"]
	'ttext = FindInStrEx("[" & cnstRExWORD &"\.]+\s*\.+\s*(���|Kind)+\(+\s*\)+\s*\=+\s*[""""]*",RWord.LineStartToCaret)
	ttext = FindInStrEx("\s+(��������������|ValueTypeStr)+\s*\(+[" & cnstRExWORD &"\.]+\s*\)+\s*\=+\s*[""""]*",RWord.LineStartToCaret)
	If Len(ttext)>0 Then
		RWord.IsIcvalVid = true
		RWOrd.BTObj = ttext
		RWOrd.BTMeth = "��������������"
		'	���������, ���� �� � ��� ������� � ������ ���� �� ���������� ����������
		'	��� � ���� �� ���, ������ ��������� ��� ��� ��� �����...
		IF InStr(1,ttext,"""") = 0 Then
			IF InStr(1,RWord.LineCaretToEnd,"""") = 0 Then
				RWord.IsNeedBrasket = True
			End IF
		End IF
	End IF
End Function

'������ ��������� � �������� ������ �������
'NameOfTableFromTable - ������������ �������� �������
Private Function GetListMethodAndRekvTablica(ResultWord,SMethodami,NameOfTableFromTable)
	GetListMethodAndRekvTablica = ""
	If SMethodami Then
		GetListMethodAndRekvTablica = GetMethodsOfTablica("")
	End If

	doc = ""
	if Not CheckWindowOnWorkbook(doc) then Exit Function

	'���� ������� ��������
	' �������� ����� � ������� �����������:
	' ��������������������.�� = ���������1, ���������2 [, ���������N]
	' ��������������������.�� = ���������1, ���������2 [, ���������N]
	If SMethodami = 0 Then
		TextMod = GetLocationText()
		Patern = "[\s|^]*(//)+[\s]*(" & ResultWord & ")+[\s]*[.]+[\s]*(��|��)+[\s]*[=]+[\s]*[" & cnstRExWORD & "\s,]+[\s]*"
		ttext = FindInStrEx(Patern,TextMod)
		ArrFindStrok = split(ttext,vbCrLf)
		If UBound(ArrFindStrok)<> -1 Then
			For ee = 0 To UBound(ArrFindStrok)
				ttext = ArrFindStrok(ee)
				if Len(ttext)>0 Then
					Patern = "(��|��)+[\s]*[=]+"
					ttextTypeSekc = FindInStrEx(Patern,ttext)
					if Len(ttext)>0 Then
						Patern = "(��|��)+[\s]*"
						ttextTypeSekc = FindInStrEx(Patern,ttextTypeSekc)
						if Len(ttext) <> 0 Then
							ttextTypeSekc = Trim(ttextTypeSekc)
							ttTypeSekc = UCase(ttextTypeSekc)
							Patern = "[=]+[\s]*[" & cnstRExWORD & "\s,]+[\s]*"
							ttextSekc = FindInStrEx(Patern,ttext)
							if Len(ttextSekc) > 0 Then
								ttextSekc = ReplaceEx(ttextSekc, Array("=","", " ",""))
								ArrSech = split(ttextSekc,",")
								Rezult = join(ArrSech, vbCrLf)
								if ttTypeSekc = "��" Then
									Rezult = "|"& Replace(Rezult, vbCrLf,vbCrLf&"|")
								End If
								If Len(GetListMethodAndRekvTablica)>0 Then
									GetListMethodAndRekvTablica = GetListMethodAndRekvTablica & vbCrLf & Rezult
								Else
									GetListMethodAndRekvTablica = Rezult
								End If
							End If
						End If
					End If
				End If
			Next
		End If
	End If
End Function

'���������� ����������, ���������� ����� "���.���"
Private Function GetTypeVidA(Variable)
	Set Var = New TheVariable
	Var.ExtractDef(Variable)

End Function

'������� ���������� ���������� "ResultWord" � �����������
'�� ����������� LastTipeVid, ������� ����� ���� ������, ����� ���� ������
'�� ������ ��� � �����
'Doc - �������� �� �������� ���������� �����
' ���������� "" ���� �� ������ ���� �������� � ���.��� ��� ���.���.�� (�� - ������ ��������� ���������������)
Private Function GetTypeVid(Variable,LastTipeVid, Doc, AsRecvOfForms, FirstWord)
	GetTypeVid	= ""
	If (Len(LastTipeVid) = 0) And (LCase(Variable) = "�����") Then
		If Len(glRWord.AddWord)>0 Then
			LastTipeVid	= "�����"
			Variable = glRWord.AddWord
		End If
	End If
	'���� � ��� Variable � �������, ����� ���������� ������� ���� ��� ��������� ���� �������...
	If InStr(1,Variable,".")>0 Then
		ttt = Split(Variable,".")
		tRezult = LastTipeVid
		For qqq = 0 To UBound(ttt)
			tRezult = GetTypeVid(ttt(qqq),tRezult, Doc, AsRecvOfForms, FirstWord)
		Next
		if Len(tRezult)>0 Then
			GetTypeVid	= tRezult
		End IF
		Exit Function
	End IF
	if (LCase(glRWord.AddWord) = LCAse("��������")) And (CompareNoCase(Variable,"�����",1)) Then
		GetTypeVid = cnstNameTypeSZ
		Exit Function
	End IF
	AsRecvOfForms = 0
	LastTipe	= ""
	LastVid		= ""
	IF InStr(1,LastTipeVid,".")>0 Then
		ttt = Split(LastTipeVid,".")
		if UBound(ttt)=1 Then
			LastTipe	= ttt(0)
			LastVid		= ttt(1)
			IF false And (UCase(LastTipe) = UCase("�������")) Then
				GetTypeVid = LastTipeVid
				Exit Function
			End IF
		End IF
	End IF
	On Error Resume Next
	If (Len(glRWord.AddWord)>0) And (lcase(Variable) = "���������") Then
		IF StrCountOccur(glRWord.AddWord, ".") = 0 Then
			Set kon  = MetaData.TaskDef.Childs(CStr("���������"))
			For tt = 0 To kon.Count-1
				if (Lcase(kon(tt).Name) = trim(lcase(glRWord.AddWord))) Then
					GetTypeVid = kon(tt).Type.FullName
					Exit Function
				End IF
			Next
		End IF
	End IF

	If LastTipeVid = "" Then
		If Is1CObject(Variable) Then
			GetTypeVid = Variable
			if Len(FirstWord)>0 Then
				if FirstWord = Variable Then
					If Len(glRWord.AddWord)<>0 And 	CompareNoCase(Variable,"������������",1) Then
						IF InStr(1,glRWord.AddWord,".") = 0 Then
							GetTypeVid = "������������." & glRWord.AddWord
							Exit Function
						End IF
					End IF
				Else
					ttt = Split(FirstWord,".")
					if UBound(ttt)=1 Then
						LastTipe	= ttt(0)
						LastVid		= ttt(1)
						IF (UCase(LastTipe) = UCase("�������")) Then
							GetTypeVid = LastTipeVid
							Exit Function
						End IF
					End IF
				End IF
			End IF
		Else
			'������� ����������� � ����� �������� ��������
			Patern = "(^)+(��������.|������.|����������.|�����.|���������.|CWBModuleDoc::)+"
			If Len(FindInStrEx (Patern, Doc.Name))>0 Then '�� � ������ �������
				if InStr(1,Doc.Name,"CWBModuleDoc::")>0 Then
					LastTipe = "CWBModuleDoc::"
				Else
					StrMetaObj = Doc.Name
					ttt = Split(StrMetaObj,".")
					LastTipe	= ttt(0)
					AsRecvOfForms = 1
					if Ubound(ttt)>0 Then
						LastVid = ttt(1)
					End IF
				End IF
			End IF
		End IF
	Else 'IF InStr(1,LastTipeVid,".") > 0 Then '������ ����������
		If UCase(LastTipeVid) = "�����" Then
			IF UCase(Variable) = UC("��������") Then
				GetTypeVid = cnstNameTypeSZ
			Else
				GetTypeVid = "#�������������#"
			End IF
		ElseIf (UCase(LastTipeVid) = UCase("������������")) And (Len(Variable)<>0) Then
			GetTypeVid = LastTipeVid & "." & Variable '���������....
		End IF
	End IF
	If UC(LastTipe) = UC("��������") Then
		if Len(LastVid)>0 Then
			IF UCase(Variable) = UCase("�������") Then
				GetTypeVid = "�������"
			ElseIF UCase(Variable) = UCase("���������������();") Then
				GetTypeVid = LastTipe & "." & LastVid
			Else
				Set MetaDoc  = MetaData.TaskDef.Childs(CStr(LastTipe))(CStr(LastVid))
				'��� ���� ����� �������������� ���� ��������� ������� � �������� � �����
				For tt = 0 To 2
					if tt=0 Then Set ch = MetaData.TaskDef.Childs("����������������������")
					if tt=1 Then Set ch = MetaDoc.Childs("�������������")
					if tt=2 Then Set ch = MetaDoc.Childs("����������������������")
					For cnmr = 0 To ch.Count-1
						Set MetaRekv = ch(cnmr)
						If UCase(MetaRekv.Name) = UCase(Variable) Then
							GetTypeVid = MetaRekv.type.FullName
							if AsRecvOfForms = 1 Then
								AsRecvOfForms = 2
							End IF
							glRWord.RecvOfForms = AsRecvOfForms
							Exit Function
						End IF
					Next
				Next
			End IF
		End IF
	ElseIF UC(LastTipe) = UC("�������") Then
		if Len(LastVid)>0 Then
			If UCase("���������������") = UCase(Variable) Then
				GetTypeVid = "��������"
				Exit Function
			Else
				Set MetaReg  = MetaData.TaskDef.Childs(CStr(LastTipe))(CStr(LastVid))
				'��� ���� ����� �������������� ���� ��������� ������� � �������� � �����
				For tt = 0 To 2
					if tt=0 Then Set ch = MetaReg.Childs("���������")
					if tt=1 Then Set ch = MetaReg.Childs("������")
					if tt=2 Then Set ch = MetaReg.Childs("��������")
					For cnmr = 0 To ch.Count-1
						Set MetaRekv = ch(cnmr)
						If UCase(MetaRekv.Name) = UCase(Variable) Then
							GetTypeVid = MetaRekv.type.FullName
							if AsRecvOfForms = 1 Then
								AsRecvOfForms = 2
							End IF
							Exit Function
						End IF
					Next
				Next
			End IF
		End IF

	ElseIF UC(LastTipe) = UC("����������") Then
		if Len(LastVid)>0 Then
			IF UCase(Variable) = "��������" Then
				GetTypeVid = LastTipe &"."& LastVid
				Exit Function
			End If
			Set MetaDoc  = MetaData.TaskDef.Childs(CStr(LastTipe))(CStr(LastVid))
			IF UCase(Variable) = "��������" Then
				'MetaDoc.Props(3)	" [����������.������������]"	String
				GetTypeVid = ReplaceEx(MetaDoc.Props(3), Array("[","", "]","", " ",""))
				Exit Function
			End If
			if Instr(1,Variable,".")>0 Then
				ArrWOfVariable = Split(Variable,".")
				for q = 0 To UBound(ArrWOfVariable)
					LastTipeVid = GetTypeVid(ArrWOfVariable(q),LastTipeVid, Doc, AsRecvOfForms, FirstWord)
				Next
				GetTypeVid = LastTipeVid
			Else
				'��� ���� ����� �������������� ���� ��������� ������� � �������� � �����
				Set ch = MetaDoc.Childs("��������")
				For cnmr = 0 To ch.Count-1
					Set MetaRekv = ch(cnmr)
					If UCase(MetaRekv.Name) = UCase(Variable) Then
						GetTypeVid = MetaRekv.type.FullName
						if AsRecvOfForms = 1 Then
							AsRecvOfForms = 2
						End IF
						Exit Function
					End IF
				Next
			End IF
		End IF
	ElseIF UC(LastTipe) = UC("������") Then
		Set MetaDoc  = MetaData.TaskDef.Childs(CStr(LastTipe))(CStr(LastVid))
		'��� ���� ����� �������������� ���� ��������� ������� � �������� � �����
		Set ch = MetaDoc.Childs("�����")
		For cnmr = 0 To ch.Count-1
			Set MetaRekv = ch(cnmr)
			If UCase(MetaRekv.Name) = UCase(Variable) Then
				Set Prop = MetaRekv.Props
				' ��� � Prop � 3 ����� ����, ������� ��� ����������,
				' ������� ������
				StrTypes = Prop(3)
				GrafaItems = Split(StrTypes,",")
				if UBound(GrafaItems)>0 Then
					GrafaItem = GrafaItems(0)
					ttt = Split(GrafaItem,".")
					If Ubound(ttt) = 2 Then
						GetTypeVid = GetTypeVid(ttt(2),ttt(0)&"."&ttt(1), Doc, AsRecvOfForms)
					End IF
				End IF
				for iii = 0 To Prop.count - 1
					'Message Prop(iii)
				Next
				'GetTypeVid = MetaRekv.type.FullName
				if AsRecvOfForms = 1 Then
					AsRecvOfForms = 2
				End IF
				Exit Function
			End IF
		Next
	End IF
	if (UCase(LastTipeVid) = "������") OR (UCase(LastTipeVid) = "���������������") Then
		IF UCase(LastTipeVid) = "������" Then
			ttextVar = GetVariableAndFunctionZapros(FirstWord,1) 	'FirstWord
		ElseIf UCase(LastTipeVid) = "���������������" Then
			ttextVar = GetColumnsFromTZ(FirstWord,1)
		End IF
		ttextVar = ReplaceEx(ttextVar, Array(",","", ";","", " ",""))
		ArrVar = Split(ttextVar,vbCrLf)
		if UBound(ArrVar)<> -1 Then
			For iii=0 To UBound(ArrVar)
				ttt = Split(ArrVar(iii),"=")
				if UBound(ttt) > 0 Then
					IF (UCase(ttt(0))=UCase(Variable)) Then
						ttt2 = Split(ttt(1),".")
						if UBound(ttt2) = 2 Then
							GetTypeVid = GetTypeVid(ttt2(2),ttt2(0)&"."&ttt2(1), Doc, AsRecvOfForms, FirstWord)
							Exit Function
						ElseIf UBound(ttt2) = 1 Then
							GetTypeVid = ttt(1)
						End IF
					End IF
				End IF
			Next
		End IF
	End IF
	if (UCase(LastTipeVid) = UCase("�������")) Or (UCase(LastTipeVid) = UCase("��������")) Then
		If UCase(LastTipeVid) = UCase("��������") Then
			LastTipeVid = "�������"
		End IF

		Set MetaReg  = MetaData.TaskDef.Childs(CStr(LastTipeVid))
		'��� ���� ����� �������������� ���� ��������� ������� � �������� � �����
		For tt = 0 To MetaReg.Count-1
			Set Register = MetaReg(tt)
			IF UCase(Register.Name) = UCase(Variable ) Then
				GetTypeVid = "�������." & Register.Name
				Exit Function
			End IF
		Next
	End IF
	'IF (UC(LastTipe) = UC("�����")) Or (UC(LastTipe) = UC("���������")) Or (UC(LastTipe) = UC("CWBModuleDoc::")) Then
	IF Len(GetTypeVid)=0 Then
		'������� �� � �����
		TableRekv = GetTableRecvFromForms(2,False)
		ListRekv = Split(TableRekv,vbCrLf)
		if UBound(ListRekv)<> -1 Then
			For eee = 0 To UBound(ListRekv)
				ITRekv = ListRekv(eee)
				if InStr(1,ITRekv," ")>0 Then
					ttt = Split(ITRekv," ")
					if UBound(ttt) = 1 Then
						If UCase(Variable) = UCase(ttt(0)) Then
							GetTypeVid = ttt(1)
							Exit For
						End IF
					End IF
				ElseIf Len(UCase(Trim(ITRekv))) > 0 Then
					If UCase(Trim(ITRekv)) = UCase(Variable) Then
						GetTypeVid = "#�������������#"
						Exit For
					End IF
				End IF
			Next
		End IF
	End IF
End Function


Private Function GetORD()
	ReturnString = ""
	Set CommonRekv = MetaData.TaskDef.Childs("����������������������")
	CommonRekvCount = CommonRekv.Count - 1
	For CRC = 0 To (CommonRekvCount)
	  	If Len(ReturnString)>0 Then
	  		ReturnString = ReturnString  &  vbCrLf & CommonRekv(CRC).Name
	  	Else
	  		ReturnString = CommonRekv(CRC).Name
	  	End IF
	Next
	GetORD = ReturnString
End Function


'SMetodami = 1 ������ ��� � ������ ������� 0 - ������ ���������
'AsRecvOfForms = 1 ���������� ��������� �����, ����� ����� ���� ������� "���������������"
Private Function GetStringRekvizitovFromObj(TypeObj, NameObj,SMetodami,AsRecvOfForms, BTMeth,	BTNumberParams, RWord)
	ReturnString = ""
	MethodsSring = ""
	Find = False
	if (UC(TypeObj) = UC("����������")) Or(UC(TypeObj) = UC("��������")) Or(UC(TypeObj) = UC("�������")) Or (UC(TypeObj) = UC("������������")) Then
	    Set Childs = MetaData.TaskDef.Childs(CStr(TypeObj))
	    For i = 0 To Childs.Count - 1
	        Set mdo = Childs(i)
	        if UC(TypeObj) = UC("�������") Then '��������
				'MetaData.TaskDef.Childs(CStr(TypeObj))(CStr(NameObj))
				if UCase(NameObj) = UCase(mdo.Name) Then
					Find = True
					Set MetaReg = MetaData.TaskDef.Childs(CStr("�������"))(CStr(NameObj))
					'������� ��������� �������� ����� �� ����� ����������� �� �������� �������....
					'��������: RWord.TypeVid2	= "���������������" � RWord.BTMeth	= "��������������"
					DasIsTZ = false

					if (CompareNoCase(RWord.TypeVid2, "���������������",1))_
						 And ((CompareNoCase(RWord.BTMeth, "��������������",1))_
						 or (CompareNoCase(RWord.BTMeth, "RetrieveTotals",1))) Then
						DasIsTZ = true
					End IF
					for regBench = 0 To 2
						NameBench = ""
						TypeBenchStr = ""
						if regBench = 0 Then
						    Set HeadBench = MetaReg.Childs("���������")
						Elseif regBench = 1 Then
						    Set HeadBench = MetaReg.Childs("������")
						Elseif regBench = 2 Then
							if DasIsTZ Then Exit For
						    Set HeadBench = MetaReg.Childs("��������")
						end if
	  					for hd_cnt = 0 To HeadBench.Count - 1
	  						set hd_chld = HeadBench(hd_cnt)
	  						If Len(ReturnString)>0 Then
	  							ReturnString = ReturnString  &  vbCrLf & hd_chld.Name
	  						Else
	  							ReturnString = hd_chld.Name
	  						End IF
	  						If SMetodami = 2 Then
	  							ReturnString = ReturnString & "=" & hd_chld.Type.FullName
	  						End IF
	  					Next
	  				Next
	  				if DasIsTZ Then
	  					ReturnString = ReturnString & vbCrLf & GetMethodsOfTablicaZnacheniy(0)
	  				else
	  					IF (SMetodami = 1) And (Len(MethodsSring) = 0) Then
		  					MethodsSring = vbCrLf & GetMethodOfRegister(1, mdo)
		  				end if
					end if
	  			end if
				If 	Find Then
	  				Exit For
	  			end if
	        end if
	        if UC(TypeObj) = UC("��������") Then '��������
				if UC(NameObj) = UC(mdo.Name) Then
					Find = True
					Set CommonRekv = MetaData.TaskDef.Childs("����������������������")
					CommonRekvCount = CommonRekv.Count - 1
					For CRC = 0 To (CommonRekvCount)
	  					If Len(ReturnString)>0 Then
	  						ReturnString = ReturnString  &  vbCrLf & CommonRekv(CRC).Name
	  					Else
	  						ReturnString = CommonRekv(CRC).Name
	  					End IF
					Next

				  	Set Head = mdo.Childs("�������������")
				  	Set Table = mdo.Childs("����������������������")
	  				for hd_cnt = 0 To Head.Count - 1
	  					set hd_chld = Head(hd_cnt)
	  					If Len(ReturnString)>0 Then
	  						ReturnString = ReturnString  &  vbCrLf & hd_chld.Name
	  					Else
	  						ReturnString = hd_chld.Name
	  					End IF
	 				Next

	 				if ((RWord.IsBetweenToken) And (CompareNoCase(RWord.BTMeth, "�����������������������", 1))) OR (trim(lcase(RWord.TypeVid2)) = "���������������") Then
	 					ReturnString = ""
	 				End IF
	  				for tbl_cnt = 0 To Table.Count - 1
	  					set tbl_chld = Table(tbl_cnt)
	  					If Len(ReturnString)>0 Then
	  						ReturnString = ReturnString  &  vbCrLf & tbl_chld.Name
	  					Else
	  						ReturnString = tbl_chld.Name
	  					End IF
	  				Next
	  			end if
	  			IF (SMetodami = 1) And (Len(MethodsSring) = 0) And Not (trim(lcase(RWord.TypeVid2)) = "���������������") Then
	  				MethodsSring = vbCrLf & GetDocMet(1)
				end if
				IF (trim(lcase(RWord.TypeVid2)) = "���������������") Then
					ReturnString = ReturnString & GetMethodsOfTablicaZnacheniy(0)
				end if
				If 	Find Then
	  				Exit For
	  			end if
	        end if
	        if UC(TypeObj) = UC("����������") Then
	          	if UC(NameObj) = UC(mdo.Name) Then
	          		Find = True
	          		if (Clng(mdo.Props("���������"))>0) Then
	 					ReturnString = "���"
	          		End if
	          		if (Clng(mdo.Props("�����������������"))>0) Then
	  					If Len(ReturnString)>0 Then
	  						ReturnString = ReturnString  & vbCrLf & "������������"
	  					End IF
	          		End if
	          		Set Ref = mdo.Childs("��������")
	          		For r = 0 To Ref.Count -1
	          			Set RefChild = Ref(r)
	  					If Len(ReturnString)>0 Then
	  						ReturnString = ReturnString  &  vbCrLf & RefChild.Name
	  					Else
	  						ReturnString = RefChild.Name
	  					End IF
	          		next
	          		ReturnString = SortStringForList(ReturnString, vbCrLf)
	          	End if
				If Find Then
		  			IF (SMetodami = 1) And (Len(MethodsSring) = 0) Then
	  					MethodsSring = vbCrLf & GetSprMet(1)
					end if
					If (Lcase(BTMeth) = Lcase("�������")) OR (Lcase(BTMeth) = Lcase("�������")) Then
						If (BTNumberParams	= 2) Then
							ReturnString = GetFormListsOfSprav(TypeObj,NameObj)
							MethodsSring = ""
						end if
					end if
	  				Exit For
	  			end if
	        End if
       		if (UC(TypeObj) = UC("������������")) And (NameObj <> "`") then
	          	if UC(NameObj) = UC(mdo.Name) Then
	          		Find = True
	          		Set Ref = mdo.Childs("��������")
	          		For r = 0 To Ref.Count -1
	          			Set RefChild = Ref(r)
	          			if Len(ReturnString)>0 Then
	          				ReturnString = ReturnString & vbCrLf & RefChild.Name
	          			Else
	          				ReturnString =  RefChild.Name
	          			End if
	          			'if glRWord.IsIcvalVid
						'ReturnString = " = " & TypeObj & "." & NameObj & "." & RefChild.Name & vbCrLf & ReturnString

	 					' ���������� ������ ���������������������,��,������������
	          		next
	          		ReturnString = SortStringForList(ReturnString, vbCrLf)
    			    ReturnString = ReturnString & vbCrLf &  LoadMethodFromFile(TypeObj)
	          	End if
				If 	Find Then
	  				Exit For
	  			end if

			end if
	    Next
	Elseif (UC(TypeObj) = UC("�����")) Or (UC(TypeObj) = UC("���������")) Or (UC(TypeObj) = UC("�����")) Or (UC(TypeObj) = UC("CWBModuleDoc::")) Then
		ReturnString = GetTableRecvFromForms(0,False)
	Elseif (TypeObj = cnstNameTypeSZ) Then
		ReturnString = GetMethodsOfListBox("")
	Elseif (TypeObj = cnstNameTypeTZ) Then
		ReturnString = GetMethodsOfTablicaZnacheniy("")
	Elseif (TypeObj = cnstNameTypeTabl) Then
		ReturnString = GetMethodsOfTablica("")
	Elseif (TypeObj = UC("�����")) Then
		ReturnString = GetMethodsText("")
	End if
	'ReturnString = SortStringForList(ReturnString, vbCrLf)
	GetStringRekvizitovFromObj = ReturnString & MethodsSring
End Function


Private Function GetFormListsOfSprav(TypeObj,NameObj)
	GetFormListsOfSprav = ""
	Set MetaRef  = MetaData.TaskDef.Childs(CStr(TypeObj))(CStr(NameObj))
	Set CHI = MetaRef.Childs
	IF CHI.Item(1).Count<> 0 Then
		For iii = 0 To CHI.Item(1).Count - 1
			If Len(GetFormListsOfSprav)>0 Then
				GetFormListsOfSprav = GetFormListsOfSprav & vbCrLf & CHI.Item(1).Item(iii).Name
			Else
				GetFormListsOfSprav = CHI.Item(1).Item(iii).Name
			End IF
		Next
	End IF

End Function



Private Function GetListKindVariable(tmpType)
	GetListKindVariable = ""
	If Len(tmpType)>0 Then
		strListKind = ""
	    Set Childs = MetaData.TaskDef.Childs(CStr(tmpType))
	    For i = 0 To Childs.Count - 1
	        Set mdo = Childs(i)
	        if Len(strListKind)>0 Then
				strListKind = strListKind & vbCrLf & mdo.Name
			else
				strListKind =  mdo.Name
			End IF
		next
		GetListKindVariable = SortStringForList(strListKind, vbCrLf)
	End IF
End Function


Private Function GetTableRecvFromForms(STipom,IsBetweenToken) ' ��������� ��������� �� �����
	IDDimRekvOfForms = ""
	IDDimRekvOfFormsTCH = ""
	FormMarkers = Array("{""Fixed"",", "{""Controls"",", "{""Cnt_Ver"",""10001""}}") '������ �������� � ������� ���������� ����� ���������, � ������ ���������������.
	ShablonEndControl = Array(""",""{"""""  ,  """"","""""  ,  """""}""},") '������ ��������� �������� �� ������ <{"Controls",>
	GetTableRecvFromForms = ""
 	If Windows.ActiveWnd Is Nothing Then
 		Exit Function
 	End If
	Set doc = Windows.ActiveWnd.Document
	If doc<>docWorkBook Then
		Exit Function
	End If
	Set docForm = Windows.ActiveWnd.Document.Page(0)
	DlgText = docForm.Stream
	Dim DimDSFormStrims
	if IsBetweenToken Then
		Patern = "[^]+[{]+[" & cnstRExWORD & """]+[,]+[0-9""]+[}]{1,2}[,]+"
		ttext = FindInStrEx(Patern,DlgText)'IsBetweenToken
		If Len(ttext)>0 Then
			IDDimRekv = ""
			Sloi = split(ttext,vbCrLf)
			If UBound(Sloi)<> -1 Then
				For ee = 0 To UBound(Sloi)
					OneSloy	= Sloi(ee)
					OneSloy = Replace(OneSloy,"""","")
					OneSloyArr = Split(OneSloy,",")
					if Len(IDDimRekv)>0 Then
						IDDimRekv = IDDimRekv & vbCrLf & OneSloyArr(0)
					else
						IDDimRekv = OneSloyArr(0)
					End if
				Next
			End if
			GetTableRecvFromForms  = IDDimRekv
			Exit Function
		End if
	End if
	DimDSFormStrims = split(DlgText,vbCrLf)
	'�������� ��������� � 4-� ������.
	LastStatus = 0 '��������� ������ ���������� ���� � �����.
	for i=4 To UBound(DimDSFormStrims)
		StringDialogStrmia = DimDSFormStrims(i)
		For FM = 0 To Ubound(FormMarkers)
			ItamFM = FormMarkers(FM)
			if InStr(1, StringDialogStrmia, ItamFM) = 1 Then
				'����� ���� "������" ���� "��������" 'Fixed	Controls
				if InStr(1,ItamFM,"Fixed")>0 Then
					LastStatus = 1
				Elseif InStr(1,ItamFM,"Controls")>0 Then
					LastStatus = 2
				Elseif InStr(1,ItamFM,"Cnt_Ver")>0 Then
					LastStatus = 0
				End If
			End If
		Next
		DSElementsStrin = split(DimDSFormStrims(i), """,""")
		cntItem = Ubound(DSElementsStrin)
		if ((cntItem > 14 ) And (Mid(DimDSFormStrims(i),1,2) = "{""")) Then
			If LastStatus = 1 Then '�������������
				IF Len(DSElementsStrin(7)) <> 0 Then
					IndexID = 7
					IDRekv = DSElementsStrin(7)
					Do While InStr(1,IDRekv,"""")>0
						IndexID = IndexID +1
					    Exit Do
					Loop
					IDRekv = DSElementsStrin(IndexID)
					if Len(IDRekv)>0 Then
						If Len(IDDimRekvOfFormsTCH) = 0 Then
							IDDimRekvOfFormsTCH = Trim(DSElementsStrin(7)) '& GetRusTypeAtribOfForms(DSElementsStrin(3))
						Else
							IDDimRekvOfFormsTCH = IDDimRekvOfFormsTCH & vbCrLf & Trim(DSElementsStrin(7)) '& GetRusTypeAtribOfForms(DSElementsStrin(3))
						End if
						if STipom = 1 Then
							IDDimRekvOfFormsTCH = IDDimRekvOfFormsTCH & " " & DSElementsStrin(3)
						elseIf STipom = 2 Then
						End if
					End if
					if UBound(DSElementsStrin)>= (13+IndexID-7) Then
						IDMetaObj = DSElementsStrin(9+IndexID-7)
						IDMetaObj2 = DSElementsStrin(13+IndexID-7) 'ID ���� �������� (��� �����������, ���������, ������������ ��� �����)
					End if

				End if
			ElseIf LastStatus = 2 Then '��� ���� ������������, ���������� �������� �� �����
				If Len(DSElementsStrin(12))>0 Then
					IndexID = 12
					prirashenie = 0
					IDRekv = DSElementsStrin(12)
					Do While InStr(1,IDRekv,"""")>0
						prirashenie = prirashenie + 1
					    Exit Do
					Loop
					IndexID = IndexID + prirashenie

					IDRekv = DSElementsStrin(IndexID)
					if Len(IDRekv)>0 Then
						If Len(IDDimRekvOfForms) = 0 Then
							IDDimRekvOfForms = Trim(IDRekv) '& GetRusTypeAtribOfForms(DSElementsStrin(1))
						Else
							IDDimRekvOfForms = IDDimRekvOfForms & vbCrLf & Trim(IDRekv) '& GetRusTypeAtribOfForms(DSElementsStrin(1))
						End if
					End if
					if STipom = 1 Then
						IDDimRekvOfForms = IDDimRekvOfForms & " " & DSElementsStrin(1)
					elseIf STipom = 2 Then
						aaa = CLng(DSElementsStrin(17))
						if (aaa<>0) Then
							Set obj  = MetaData.FindObject(aaa)
							IDDimRekvOfForms = IDDimRekvOfForms & " " & obj.FullName
						ElseIF (DSElementsStrin(1) = "LISTBOX") OR (DSElementsStrin(1) = "COMBOBOX") Then
							IDDimRekvOfForms = IDDimRekvOfForms & " " & "��������������"
						ElseIF DSElementsStrin(1) = "TABLE" Then
							IDDimRekvOfForms = IDDimRekvOfForms & " " & "���������������"
						ElseIF DSElementsStrin(14) = "B" Then
							IDDimRekvOfForms = IDDimRekvOfForms & " " & "����������" '& "����������."
						ElseIF DSElementsStrin(14) = "O" Then
							IDDimRekvOfForms = IDDimRekvOfForms & " " & "��������"	'& "��������."
						End if
					End if

					if UBound(DSElementsStrin)>= (13+IndexID-7) Then
						IDMetaObj2 = DSElementsStrin(13+IndexID-7) 'ID ���� �������� (��� �����������, ���������, ������������ ��� �����)
					End if
				End if
			end if
		end if
	Next
	iii = 12
	IDDimRekv = IDDimRekvOfForms & vbCrLf &  IDDimRekvOfFormsTCH
	GetTableRecvFromForms  = IDDimRekv
End Function

Private Function GetMethodsText(GGG)
	MethodsTextArr = Array("���������������()", _
	"��������������(<?>)", _
	"�������(<?>);", _
	"������(<?>)", _
	"����������(<?>)", _
	"��������������(<?>,);", _
	"��������������(<?>);", _
	"��������������(<?>,);", _
	"�������������(<?>);", _
	"��������������(<?>)", _
	"��������(<?>,);", _
	"��������();", _
	"���������������(<?>)", _
	"��������(<?>);")
	MethodsText = Join(MethodsTextArr,vbCrLf)
	GetMethodsText = SortStringForList(MethodsText, Spliter)

End Function
private Function GetMethodsOfFormRecv()
	MethodsOfFormRecv = Array("���������(<?>);", _
	"�����������(<?>);", _
	"��������������(<?>);", _
	"����(<?>);", _
	"���������(<?>);", _
	"�����������(<?>);", _
	"����������������������������������(<?>);", _
	"�������������(<?>);", _
	"������������(<?>);", _
	"�������������(<?>);", _
	"�����������������(<?>,);")
	GetMethodsOfFormRecv = Join(MethodsOfFormRecv,vbCrLf)
	MethodsOfFormRecv = SortStringForList(GetMethodsOfFormRecv, Spliter)
End Function

private Function GetMethodsOfForm()
	MethodsOfFormRecv = Array("��������������(<?>)", _
	"��������", _
	"�������������������", _
	"��������(<?>);", _
	"��������������������(<?>)", _
	"����������������(""<?>"",2)", _
	"��������", _
	"��������.����������������(<?>,);", _
	"���������(<?>,)", _
	"������������������(<?>);", _
	"�����������������(<?>);", _
	"���������������������(<?>);", _
	"��������������������������(<?>,);", _
	"�������������������������(<?>)", _
	"��������������(<?>,);", _
	"��������������������(<?>);", _
	"��������������(<?>);", _
	"�����������();", _
	"��������������()", _
	"���������������(<?>);", _
	"���������������()", _
	"��������������();", _
	"�������();")
	GetMethodsOfForm = Join(MethodsOfFormRecv,vbCrLf)
	GetMethodsOfForm = SortStringForList(GetMethodsOfForm, Spliter)
End Function

private Function GetMethodsOfFormPodbor()
	GetMethodsOfFormPodborRecv = Array("��������������(<?>)", _
	"��������", _
	"�������������������", _
	"��������(<?>);", _
	"����������������(<?>);", _
	"��������������������(<?>,);", _
	"���������������������(<?>,);", _
	"�������������������(<?>,)", _
	"���������������������(<?>,);", _
	"���������������������������������(<?>,);", _
	"����������(<?>,);", _
	"���������������(<?>,);", _
	"�������������(<?>,);", _
	"����������(<?>)", _
	"��������������(<?>,);", _
	"���������������������������(<?>);", _
	"�����������(<?>);", _
	"���������������(<?>);")
	GetMethodsOfFormPodbor = Join(GetMethodsOfFormPodborRecv,vbCrLf)
	GetMethodsOfFormPodbor = SortStringForList(GetMethodsOfFormPodbor, Spliter)
End Function



private Function GetMethodsOfTablica(InForms)
	MethodsOfTablica = Array("�������������",_
	"���������������(<?>);",_
	"�������(<?>);",_
	"�������();",_
	"��������������(<?>);",_
	"�������������(<?>);",_
	"������������������(<?>);",_
	"�������������(<?>);",_
	"������������(<?>);",_
	"�������������()",_
	"�������������()",_
	"������������(<?>)",_
	"������������(<?>)",_
	"��������������(<?>)",_
	"��������(<?>);",_
	"��������();",_
	"������(<?>)",_
	"��������(<?>,);",_
	"�������(<?>)",_
	"������������������������(<?>,);",_
	"�������������������������(<?>,);",_
	"�����(<?>);",_
	"�����������������(<?>,);",_
	"���������������������(<?>)",_
	"����������(<?>);")
	GetMethodsOfTablica = Join(MethodsOfTablica,vbCrLf)
	GetMethodsOfTablica = SortStringForList(GetMethodsOfTablica, Spliter)
End Function




private Function GetMethodsOfTablicaZnacheniy(InForms)
	if Len(InForms) = 0 Then
		MethOfListBox = Array("�����������", _
		"�����������������(<?>)", _
		"������������(<?>);", _
		"���������������(<?>);", _
		"��������������(<?>);", _
		"��������������������������(<?>);", _
		"������������������������(<?>);", _
		"���������������()", _
		"�����������();", _
		"�������������(<?>);", _
		"�������������();", _
		"�������������();", _
		"��������������() = 1", _
		"�������������(<?>);", _
		"����������������������(<?>);", _
		"��������������(<?>,);", _
		"������������������(<?>);", _
		"����������������(<?>,);", _
		"�������������(<?>);", _
		"�����������(<?>,);", _
		"��������();", _
		"����(<?>);", _
		"���������(<?>,);", _
		"��������(<?>,);", _
		"���������(<?>,);", _
		"���������(<?>);")
	Else
		MethOfListBox = Array("�����������", _
		"�����������������(<?>)", _
		"������������(<?>);", _
		"���������������(<?>);", _
		"��������������(<?>);", _
		"��������������������������(<?>,);", _
		"������������������������(<?>,);", _
		"���������������()", _
		"�����������(<?>);", _
		"�������������(<?>);", _
		"�������������();", _
		"�������������();", _
		"��������������()", _
		"�������������(<?>);", _
		"����������������������(<?>);", _
		"��������������(<?>,);", _
		"������������������(<?>);", _
		"����������������(<?>,);", _
		"�������������(<?>);", _
		"�����������(<?>,);", _
		"��������();", _
		"����(<?>);", _
		"���������(<?>,);", _
		"��������(<?>,);", _
		"���������(<?>,);", _
		"���������(<?>);", _
		"�������������();", _
		"��������������(<?>,);", _
		"����������������(<?>);", _
		"�����������(<?>,);", _
		"�������������������(<?>,);")
	End if
	GetMethodsOfTablicaZnacheniy = Join(MethOfListBox,vbCrLf)
	GetMethodsOfTablicaZnacheniy = SortStringForList(GetMethodsOfTablicaZnacheniy, Spliter)
End Function


private Function GetMethodsOfListBox(Par)
	if UCase(Par) = "��������" Then
		MethOfListBox = Array("����������������(<?>,);", _
		"����������������(<?>,);", _
		"������������()", _
		"������������������(<?>,);", _
		"����������(<?>,);", _
		"�����������(<?>)", _
		"�������������(<?>)", _
		"�������������();", _
		"���������������(<?>);", _
		"����������();", _
		"����������������(<?>,);", _
		"���������(<?>);")
	Else
		MethOfListBox = Array("����������������(<?>);", _
		"����������������(<?>,);", _
		"������������()", _
		"������������������(<?>,);", _
		"����������(<?>,);", _
		"�����������(<?>,);", _
		"��������������������������(<?>);", _
		"�����������(<?>)", _
		"�������������(<?>)", _
		"����������������(<?>);", _
		"��������(<?>)", _
		"����������������������(<?>)", _
		"���������������������()", _
		"���������������(<?>)", _
		"����������������(<?>,)", _
		"�������(<?>,)", _
		"�������������(<?>);", _
		"���������������(<?>);", _
		"����������();", _
		"����������������(<?>,);", _
		"���������(<?>);")
	End if
	GetMethodsOfListBox = Join(MethOfListBox,vbCrLf)
	GetMethodsOfListBox = SortStringForList(GetMethodsOfListBox, Spliter)
End Function

private Function GetRusTypeAtribOfForms(Param)
	RetValue = "<�����>"
	Param = Trim(Param)
	if Len(Param) <> 0 Then
		If Param = "1CEDIT" Then
			RetValue = "  <���� ����� (������/�����/����)>"
		ElseIf Param = "STATIC" Then 			RetValue = "  <�����>"
		ElseIf Param = "BMASKED" Then
			RetValue = "  <����� ��������>"
		ElseIf Param = "BUTTON" Then
			RetValue = "  <������>"
		ElseIf Param = "PICTURE" Then
			RetValue = "  <��������>"
		ElseIf Param = "COMBOBOX" Then
			RetValue = "  <���������� ������>"
		ElseIf Param = "TABLE" Then
			RetValue = "  <������� ��������>"
		ElseIf Param = "TABLE" Then
			RetValue = "  <������� ��������>"
		ElseIf Param = "1CGROUPBOX" Then
			RetValue = "  <����� ������>"
		ElseIf Param = "CHECKBOX" Then
			RetValue = "  <������>"
		ElseIf Param = "RADIO" Then
			RetValue = "  <�������������>"
		ElseIf Param = "LISTBOX" Then
			RetValue = "  <�������������>"
		End If
	End If
	GetRusTypeAtribOfForms = RetValue
End Function

private Function WordIsAStatment(ResultWord)
	RetValue = False
	ArrStatments = GetListOFStatments()
	Find = 0
	For iii = 0 To UBound(ArrStatments)
		If UCase(Trim(ResultWord)) = UCase(Trim(ArrStatments(iii))) Then
			RetValue = True
			Exit For
		End IF
	Next
	WordIsAStatment = RetValue
End Function

Private Function WordInArray(Arr, Word) GetListOFStatments()
	WordInArray = False
	For qq = 0 To UBound(Arr)
		IF UCase(Arr(qq)) = UCase(Word) Then
			WordInArray = True
			Exit For
		End IF
	Next
End Function

private Function GetListOFStatments()
	MethOfListBox = Array("���������", "Procedure", "��������������", "EndProcedure", _
	"�������", 	"Function", "������������", "EndFunction", _
	"����", "If", "�����", "Then", "���������", "ElsIf", "�����", "Else", "���������", "EndIf", _
	"����", "While", "����", "Do", "����������", "EndDo", "���", "For", "��", "To", _
	"�������", "Try", "����������", "Except", "������������", "EndTry", "�����������������", "Raise", _
	"�������", 	"Goto", "����������", "Continue", "��������", "Break", "�������", 	"Return", _
	"����������������", "LoadFromFile", _
	"�����", "Forward", "�������", 	"Export", _
	"������������������", "PageBreak", "����������������", "LineBreak", "���������������", 	"TabSymbol")
	GetListOFStatments = MethOfListBox
End Function

Private Function GetMethodOfRegister(ASCreateObj, mdo)
	'"-<-----�����----->-", _
	MethodOfRegister = Array("������", _
	"������", _
	"���()", _
	"�����������������();", _
	"������������(<?>,)", _
	"�����������������(<?>,);", _
	"���������������(<?>);", _
	"���������������(<?>);", _
	"������������������������(<?>);", _
	"���������������(<?>);", _
	"����������������();", _
	"���������������();", _
	"�����������();", _
	"������������();", _
	"������������();", _
	"���������������(<?>);", _
	"��������������(<?>);", _
	"����������������(<?>,);", _
	"�������������������������(<?>);")
	GetMethodOfRegister = Join(MethodOfRegister,vbCrLf)

	if mdo.Props(3) = "�������" Then
		'mdo.Props(3)	"�������"	String
		'"-<-----������� ��������----->-", _
		MethodOfRegister = Array("�������(<?>,)", _
		"��������������(<?>,)", _
		"�������(<?>,);", _
		"��������������(<?>,);", _
		"���������������();", _
		"�������������������������(<?>,);", _
		"��������������(<?>,);", _
		"��������������(<?>,);", _
		"�����������������������();", _
		"�����������������������();")
	else
	'"-<-----���������----->-", _
		MethodOfRegister = Array("������������������(<?>);", _
		"����(<?>)", _
		"�����(<?>,);", _
		"�������������();",_
		"��������(<?>,);", _
		"�����������������();")
	End IF
	GetMethodOfRegister = GetMethodOfRegister & vbCrLf & Join(MethodOfRegister,vbCrLf)

	'"-<-----���������----->-", _
	MethodOfRegister = Array("��������������������(<?>,);", _
	"��������������������(<?>,);", _
	"�����������������(<?>);")
	GetMethodOfRegister = GetMethodOfRegister & vbCrLf & Join(MethodOfRegister,vbCrLf)
End Function


private Function SortStringForList(SortStr, Delimiter)
	SortStringForList = CommonScripts.SortString(SortStr, Delimiter)
End Function


private Function GetSprMet(AsCreateObj)
	'IF AsCreateObj = 1 Then
	Method_Spr = Array("��������",_
	"���", _
	"������������", _
	"��������", _
	"��������(<?>)", _
	"����������(<?>,)", _
	"���()", _
	"�����������������();", _
	"�������()", _
	"�����������������(<?>,);", _
	"���������������(<?>);", _
	"���������()", _
	"�����������������(<?>)", _
	"������()", _
	"�������(<?>,)", _
	"�������������(<?>)", _
	"�����������(<?>)", _
	"��������������()", _
	"���������()", _
	"������������������()", _
	"������������(<?>)", _
	"�����������(<?>,);", _
	"�������������������(<?>);", _
	"����������������(<?>);", _
	"���������������(<?>)", _
	"��������������������������(<?>,)", _
	"���������������(<?>)", _
	"���������������() = 1", _
	"����������������(<?>);", _
	"���������������������(<?>,);", _
	"��������������������(<?>,);", _
	"�������������������(<?>);", _
	"������������();", _
	"�������������������();", _
	"����������������(<?>);", _
	"�����();", _
	"�����������();", _
	"�����������(<?>)", _
	"������������������(<?>);", _
	"������������(<?>,);", _
	"��������();", _
	"�������(<?>);", _
	"����������(<?>);", _
	"��������������������();", _
	"���������������()", _
	"����������������(<?>);", _
	"��������������������(<?>,);", _
	"���������������������(<?>,);", _
	"�������������������(<?>,)", _
	"���������������������(<?>,);", _
	"���������������������������������(<?>,);", _
	"����������(<?>,);", _
	"���������������(<?>,);", _
	"�������������(<?>,);", _
	"����������(<?>)", _
	"��������������(<?>,);", _
	"���������������������������(<?>);", _
	"�����������(<?>)", _
	"���������������(<?>);")
	GetSprMet=Join(Method_Spr,vbCrLf)
End Function

private Function GetDocMet(ASRekvOfForm)
	Method_Doc = Array("�����������������(); ", _
	"������������������(); ", _
	"������������������(); ", _
	"����������������(); ", _
	"����������������(); ", _
	"��������������������(<?>); ", _
	"����������(<?>); ", _
	"���() ", _
	"�������������(<?>) ", _
	"������() ", _
	"�������(<?>) ", _
	"����������������(<?>,) ", _
	"���������������������������(<?>) ", _
	"�����������������(<?>,) ", _
	"���������������(<?>) ", _
	"���������������������������(<?>) ", _
	"�������������() ", _
	"�����������������������(<?>,); ", _
	"�����(<?>); ", _
	"������������������() ", _
	"������� ", _
	"�����������������������(<?>); ", _
	"��������(); ", _
	"��������(); ", _
	"���������������������(<?>); ", _
	"������������������(<?>,); ", _
	"����(<?>) ", _
	"��������������() ", _
	"���������������() ", _
	"������������������() ", _
	"������������(<?>,) ", _
	"�������������(<?>) ", _
	"�������������(<?>) ", _
	"�������������������(); ", _
	"�����������(); ", _
	"�����(); ", _
	"�������� ", _
	"���������������(<?>); ", _
	"�������� ", _
	"����������������(<?>); ", _
	"���������������(<?>); ", _
	"�������������(<?>) ", _
	"����������������() ", _
	"���������������(); ", _
	"��������������() ", _
	"����������������������(<?>) ", _
	"���������������() ", _
	"�����������������(); ", _
	"�������������(<?>) ", _
	"�����������������(<?>); ", _
	"����������������������(<?>); ", _
	"�����������������������������(<?>,); ", _
	"��������() ", _
	"��������(<?>,); ", _
	"����������������(<?>,); ", _
	"��������������������(); ", _
	"��������������������(); ", _
	"�����������������(<?>); ", _
	"����������() ", _
	"���������������() ", _
	"�������(<?>); ", _
	"�������������(); ", _
	"�������������(); ", _
	"�����������������(<?>,); ", _
	"���������������(<?>); ", _
	"��������������������(<?>); ", _
	"�����������������������������(<?>); ", _
	"����������������(<?>);")
	GetDocMet=Join(Method_Doc,vbCrLf)
End Function

private Function GetZaprosMet()
	Method_Zapros = Array("��������������()", _
		"���������(<?>)", _
		"��������������()", _
		"�����������()", _
		"�������������()", _
		"������������()", _
		"���������(<?>)", _
		"��������(<?>)", _
		"���������(<?>)", _
		"����������������������(<?>,)", _
		"���������������(<?>)")
	GetZaprosMet=Join(Method_Zapros,vbCrLf)
End Function

Private Function GetMethodsRegisters()
	GetMethodsRegisters = ""
	ArrMethodsRegisters = Array("���������������(<?>);", _
	"��������������������(<?>);", _
	"��������������������(<?>);", _
	"������������(<?>);")
	GetMethodsRegisters = Join(ArrMethodsRegisters,vbCrLf)
End Function

Private Function FindReturnValueFunction(tAllText)
	FindReturnValueFunction = ""
	tempType = FindInStrEx("(;|^|\s)+(Return|�������)+\s+[" & cnstRExWORD & "]+\s*[;]+", tAllText)
	If Len(tempType)>0 Then
		ArrFind = Split(tempType,vbCrLf)
		If UBound(ArrFind)<>-1 Then
			For qq = 0 To UBound(ArrFind)
				tempType = FindInStrEx("(;|^|\s)+(Return|�������)+\s+", ArrFind(qq))
				ArrFind(qq) = Replace(ArrFind(qq),tempType,"")
				ArrFind(qq) = ReplaceEx(ArrFind(qq),Array(";","", vbTab," "))
				ArrFind(qq) = Trim(ArrFind(qq))
				if Len(ArrFind(qq))>0 Then
					FindReturnValueFunction = ArrFind(qq)
					Exit For
				End IF
			Next
		End IF
	End IF
End Function



'� AddWords ����� ���������� �������� ��������� ��
'"���.���������������("������������");" AddWords = ������������
Private Function GetTypeFromTextRegExps(ResultWord, AddWords)
	'message ResultWord & " + " & AddWords
	GetTypeFromTextRegExps = ""
	if ResultWord = "?" Then
		Exit Function
	End If
	If GlobalVariableType.Exists(lcase(ResultWord)) Then
		GetTypeFromTextRegExps = GlobalVariableType.Item(lcase(ResultWord))
		IF Len(GetTypeFromTextRegExps)>0 Then
			Exit Function
		End If
	End If
	if (InStr(1, ResultWord,".") > 0) And False Then
		Set nVar = New TheVariable
		nVar.ExtractDef(ResultWord)
		ResultWord = nVar.V_Type
		'if (nVar.WordsCnt>1)
	End If

	doc = ""
	If glDoc <> "" Then
		Set doc = glDoc
	Else
		if Not CheckWindow(doc) then Exit Function
	End If

	AddText = ""
	RezTempTypeAdd = "" '��������� ��������� ���������..... ������������ ���� ������ ������� � ��������� �� �����....
	'virajenieCO = "[\s|^|;]+(" & ResultWord & ")[\s]*["

	'���� �����, ���� �� � ���������� ������, ����� �������� ������ ����� ������ ���������
	'�����, ������� �������� ������ ����� ���������, �����, ���� �� ��������, �����
	'���� ����� � ������ � �� �����
	'����� �������� ���������� � ��������� ������� � ���� �������

	StartLileEnd = 0
	nctStep	= 1
	if doc.Name = "���������� ������" Then
		nctStep = 0 '� ����������� �� ����� ���� �����
		if (glDoc<>"") And (glStartLileEnd<>-1) Then
			StartLileEnd = glStartLileEnd
		End IF
	End IF
	Rezult = ""
	For cStep = 0 To nctStep
		If cStep = 0 Then
			StartLile = StartScanLineText
		Else
			StartLile = Doc.LineCount
		End If
		For i = StartLile To StartLileEnd Step -1
			StartScanLineText = i
			ttext = " " & Trim(Doc.range(i))
			If InStr(1,UC(ttext),UC(ResultWord))>0 Then
				If inStr(1,ttext,"//")>0 Then
					Patern = "(\s|^)*(//)+(\s)*(" & ResultWord& ")(\s)*[=]+(\s)*[""]*(\s)*[a-zA-z�-��-�0-9_.]+(\s)*[""]*(\s)*" '(\+)*(\s)*(���������������)*"
					tempType = FindInStrEx(Patern, ttext)
					If Len(tempType) = 0 Then
						Patern = "(\s|^)*(//)+(\s)*(" & ResultWord & "." & AddWords	 & ")(\s)*[=]+(\s)*[""]*(\s)*[a-zA-z�-��-�0-9_.]+(\s)*[""]*(\s)*"' (\+)*(\s)*(���������������)*"
						tempType = FindInStrEx(Patern, ttext)
					End If
					If Len(tempType)>0 Then
						RezultStr = tempType
						tempType = Replace(tempType,"""", "")
						tempType = FindInStrEx("[=]+(\s)*[a-zA-z�-��-�0-9_.]+", tempType)
						RezultStr = Replace(RezultStr,tempType,"")
						RezultStr = ReplaceEx(RezultStr,Array("//","", vbTab,""," ",""))
						If Len(tempType)>0 Then
							tempType = FindInStrEx("[a-zA-z�-��-�0-9_.]+", tempType)
							If Len(tempType)>0 Then
								tempType = Trim(tempType)
								tempTypeArr = Split(tempType,".")
								if UBound(tempTypeArr) = 1 Then
									IF (UCase(tempTypeArr(0)) = UCase("��������")) OR (UCase(tempTypeArr(0)) = UCase("����������")) OR (UCase(tempTypeArr(0)) = UCase("�������")) Then
										if CompareNoCase(ResultWord,RezultStr,1) Then
											If Len(AddWords)<>0 Then
												tempType = GetTypeVid(AddWords,tempType, Doc,0,"")
											End If
										End If
										if Len(tempType)<>0 Then
											AddWords = ""
											Rezult = tempType
											Exit For
										End If
									End If
								Else
									AddWords = ""
									Rezult = tempType
									Exit For
								End If
							End If
						End If
					End If
					ttext = Mid(ttext,1,inStr(1,ttext,"//"))
				End If
				' ��������� ������ �� ������� ������ ��������/������� �������� ��� "�����������������" ��
				'����������������������������������("��������.������������������", �����������������);
				RezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "]+\s*\(+.*("&ResultWord&")+.*\)+",ttext) '���������� ��������� � ���������
				tRezultStr = FindInStrEx("(\s|^|;)+("&ResultWord&")+\s*\=+\s*[" & cnstRExWORD & "]+\(+.*\)+",ttext)	' ��� "�������������" �� -> "������������� = �����������������������������(��������);"

				if Len(RezultStr)>0 Or Len(tRezultStr)>0 Then
					Parameters = -1 '��� ����� �����, ������� ���� ����������... -1 �������� ������������ ��������..
					FindPF = false : FindPFInGlobalModule = False
					' � RezultStr �������� ���
					' � tRezultStr �������� ���������
					Shema = 1
					if Len(RezultStr)>0 Then '���� ��������� ������ �����
						RezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "]+\s*\(+.*("&ResultWord&")+.*\)+",RezultStr)
						tRezultStr = FindInStrEx("\(+.*("&ResultWord&")+.*\)+",RezultStr)
						'��������� ���..
						RezultStr = Replace(RezultStr,tRezultStr,"")
						RezultStr = ReplaceEx(RezultStr, Array(vbTab,"", vbCrLf,"", vbCr,""))
						tempTypeArr = Split(GetParams(tRezultStr),vbCrLf)
						If UBound(tempTypeArr)<>-1 Then
							For qq = 0 To UBound(tempTypeArr)
								'Stop
								If InStr(1,LCase(tempTypeArr(qq)),LCase(ResultWord))>0 Then
									Parameters = qq + 1
									Exit For
								End If
							Next
						End If
					Elseif Len(tRezultStr)>0 Then '���� ��������� ������ �����
						Shema = 2
						RezultStr = FindInStrEx("\=+\s*[" & cnstRExWORD & "]+\(+",tRezultStr)
						RezultStr = ReplaceEx(RezultStr, Array(vbTab,"", vbCrLf,"", vbCr,"", "=","", "(",""))
						tRezultStr = FindInStrEx("\(+.*\)+",tRezultStr)
					End If
					RezultStr = Trim(RezultStr)
					Procedure = ""
					If Not IsEmpty(LocalModule) Then
						On Error Resume Next
						if LocalModule.IsProcedure(RezultStr, "", "", "","",Procedure) Then
							FindPF = True
						End If
					End If
					IF Not IsEmpty(GlobalModule) And GlobalModule<>"" Then
						if LocalModule.ModuleName<>GlobalModule.ModuleName Then
							if GlobalModule.IsProcedure(RezultStr, "", "", "","",Procedure) Then
								FindPF = True : FindPFInGlobalModule = True
							End If
						End If
					End If
					if FindPF Then
						tRezultStr = GetParams(tRezultStr)
						tempTypeArr = Split(tRezultStr,vbCrLf)
						If UBound(tempTypeArr)<>-1 Then
							For qq = 0 To UBound(tempTypeArr)
								'Stop
								If InStr(1,LCase(tempTypeArr(qq)),LCase(ResultWord))>0 Or Shema = 2 Then
									tRezultStr = glRWord.ArrProcFuncFromTheWord
									glRWord.ArrProcFuncFromTheWord = tRezultStr
									If AddToStringUni(tRezultStr, RezultStr,",") Then
										tRezultStr = glRWord.ArrProcFuncNumbParam
										if Shema = 1 Then
											AddToStringUni tRezultStr, ""&(qq+1),","
										Else
											AddToStringUni tRezultStr, ""&(-1),","
										End IF
										glRWord.ArrProcFuncNumbParam = tRezultStr
										End If
									Exit For
								End If
							Next
						End If
					End If
					'���� ����� � �����������, ����� ���-�� ������, ���� ���������� ���...
					'������������� �� ���������� � ������� ���...
					if FindPFInGlobalModule Then
						'Stop
							tRezultStr = FindReturnValueFunction(Procedure.text)
							Procedure.RetValueStr = tRezultStr
						if Parameters <> -1 Then '��� ���� �������
							tRezultStr = Procedure.GetParamNumber(Parameters)
						End If
						If Len(tRezultStr)>0 Then
							'���������� ������� ������������
							LastglStartLileEnd = glStartLileEnd
							LastStartScanLineText = StartScanLineText
							glStartLileEnd = Procedure.LineStart
							StartScanLineText = Procedure.LineEnd

							Set glDoc = Documents("���������� ������")
							Set tLastDoc = Doc
							tRezultStr = GetTypeFromTextRegExps(tRezultStr, AddWords)
							'��������������� ������� ������������
							glDoc = ""
							Set Doc = tLastDoc
							LastglStartLileEnd = glStartLileEnd
							LastStartScanLineText = StartScanLineText
							If Len(tRezultStr)>0 Then
								Rezult = tRezultStr
								AddWords = ""
								Exit For
							End If

						End If
					End If
				End If

				'GetColumnsFromTZ(ResultWord, STipom)
				RezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "]+(\s)*[.]+(\s)*(�����������������������|UnloadTable)+(\s)*[\(]+(\s)*(" & ResultWord &")+.*[\)]+(\s|^|;)", ttext)
				tRezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "\s\.]+(\s)*[.]+(\s)*(�����������������������|UnloadTable)+(\s)*[\(]+(\s)*(" & ResultWord &")+.*[\)]+(\s|^|;)", ttext)
				if Len(RezultStr)> 0 Then
					RezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "]+(\s)*[.]+", RezultStr)
					if Len(RezultStr)> 0 Then
						RezultStr = FindInStrEx("[" & cnstRExWORD & "]+", RezultStr)
						if Len(RezultStr)> 0 Then
							RezultStr = GetTypeFromTextRegExps(RezultStr,"")
							if Len(RezultStr)> 0 Then
								IF Len(AddWords)>0 Then
									RezultStr = GetTypeVid(AddWords,RezultStr, Doc,0,"")
									if Len(RezultStr)>0 Then
										Rezult = RezultStr
									End If
								Else
								'Stop
									glRWord.BTMeth = "�����������������������"
									Rezult = RezultStr & "+���������������"
								End If
								exit For
							End If
						End If
					End If
				Elseif Len(tRezultStr)> 0 Then
					'Stop
					tRezultStr = FindInStrEx("(\s|^|;)+[" & cnstRExWORD & "\s\.]+(\s)*[.]+", tRezultStr)
					if Len(tRezultStr)> 0 Then
						tRezultStr = FindInStrEx("[" & cnstRExWORD & "]+", tRezultStr)
						RezultStr = ""
						if InStr(1,tRezultStr,vbCrLf)>0 Then
							tempTypeArr	= Split(tRezultStr,vbCrLf)
							tRezultStr = ""
							for qqq = 0 To UBound(tempTypeArr)
								tempTypeArr(qqq) = trim(tempTypeArr(qqq))
								if qqq > 0 Then
									if Len(tRezultStr)> 0 Then
										tRezultStr = tRezultStr &"."&tempTypeArr(qqq)
									else
										tRezultStr = tempTypeArr(qqq)
									End IF
								else
									RezultStr = tempTypeArr(qqq)
								End IF
							Next
						End IF
						if Len(tRezultStr&RezultStr)> 0 Then
							tRezultStr = GetTypeFromTextRegExps(RezultStr,tRezultStr)
							if Len(AddWords)>0 Then
								tRezultStr = GetTypeVid(AddWords,tRezultStr, Doc,0,"")
								if Len(tRezultStr)> 0 Then
									AddWords = ""
									Rezult = tRezultStr
									exit For
								End If
							Else
								if Len(tRezultStr)> 0 Then
								'Stop
									if InStr(1,tRezultStr,"���������������") = 0 Then
										Rezult = tRezultStr & "+���������������"
									Else
										Rezult = tRezultStr
									End If
									exit For
								End If
							End If

						End If
					End If
				End If


				Patern = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*(���������|Unload)+(\s)*[\(]+(\s)*(" & ResultWord &")+.*[\)]+(\s|^|;)"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then					'�������� ���������� �������
					tRezultStr = ""
					textVarZapros = FindInStrEx("(\s)*[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*",RezultStr)
					'����� ����� - ������ ��������� ����� ���� � �� �������� :(
					' ���������............
					If Len(textVarZapros)>0 Then						'��������
						textVarZapros = FindInStrEx("[a-zA-Z�-��-�0-9_]+",textVarZapros)
					End IF
					Patern = "[\(]+(\s)*(" & ResultWord &")+.*[\)]+"
					RezultStr = FindInStrEx(Patern,RezultStr)
					RezultStr = ReplaceEx(RezultStr,Array("(","",")",""," ",""))
					RezultStr = FindInStrExA("[" & cnstRExWORD & "]+",RezultStr)
					If UC(RezultStr) = UC(ResultWord) Then
						tempType = GetTypeFromTextRegExps(textVarZapros,"")
						if InStr(1,tempType,"+���������������")>0 Then
							tempType = Replace(tempType,"+���������������","")
						End if
						if (UC(tempType) = "������") Then
							RezultStr = GetVariableAndFunctionZapros(textVarZapros, 1)
						Elseif (UC(tempType) = UC(cnstNameTypeTZ)) Then
							RezultStr = GetColumnsFromTZ(textVarZapros, 1)
						ElseIf Instr(1,Lcase(tempType),LCase("�������."))>0 Then
							tempTypeArr = Split(tempType,".")
							glRWord.TypeVid2 = "���������������"
							RezultStr = GetStringRekvizitovFromObj(tempTypeArr(0), tempTypeArr(1),2,0, "",	0, glRWord)
						End if
						'AddWords - �� ������
						if Len(RezultStr)>0 Then
							tempTypeArr = Split(RezultStr,vbCrLf)
							tRezultStr = "" : Arr2 = "" : Arr3 = ""
							if InStr(1,AddWords,".")>0 Then
								Arr3 = Split(AddWords,".")
							Else
								Arr3 = Array(AddWords)
							End IF
							if UBound(tempTypeArr)>0 Then
								For qqq = 0 To UBound(tempTypeArr)
									RezultStr = tempTypeArr(qqq)
									if inStr(1,RezultStr,"=")>0 Then
										Arr2 = Split(RezultStr,"=")
										if UBound(Arr2)>0 Then ' ��� ��� ������ ��� � �����..
											IF Lcase(Trim(Arr2(0))) = Trim(LCase(Arr3(0))) Then 'AddWords
												Arr2(1) = Replace(Arr2(1),";",",")
												Arr2(0) = Split(Arr2(1),",")
												if UBound(Arr2(0))<>-1 Then
													tRezultStr = trim(Arr2(0)(0))
													Exit For
												End if
											End if
										End if
									End if
								Next
								if Len(tRezultStr)>0 Then
									Arr2 = Split(tRezultStr,".")
									if UBound(Arr2) = 2 Then
										tRezultStr = GetTypeVid(Arr2(2),Arr2(0)&"."&Arr2(1), Doc,0,"")
									ElseIf UBound(Arr3)>0 Then
										'���������� ��������� ���� � AddWords
										AddWords = ""
										For qqq = 1 To UBound(Arr3)
											tRezultStr = GetTypeVid(Arr3(qqq),tRezultStr, Doc,0,"")
										Next
									End if
									if Len(tRezultStr)>0 Then
										Rezult = tRezultStr
										AddWords = ""
										Exit For
									End if
									'Exit Function
								End if
								'Rezult = Rezult & vbCrLf &  RezultStr
							End if
						End if
						'Exit For
					End if
				End if


				' ��� ���������� ��������� ��� ���������� �������� "������" ����� ��������� ����������� ����
				'" 		��� (������=""����������������"")" ������ �� ��� ������ ���� ���� ��������� � ���������
				' � ����� ����� ����������...
				If glRWord.IsIcvalVid Then
					if InStr(1,ResultWord,".") = 0 Then
						FindFirstInFindInStrEx = True '����� ������ ������ ���������
						tempType = FindInStrEx("[\s|;|\(]+(" &ResultWord&")+\s*\=+\s*[""""]+[" & cnstRExWORD & "]+[""""]", ttext)
						FindFirstInFindInStrEx = false
						if Len(tempType)>0 Then
							tempType = FindInStrEx("\=+\s*[""""]+[" & cnstRExWORD & "]+[""""]", tempType)
							if Len(tempType)>0 Then
								tempType = ReplaceEx(tempType, Array("""","", "=",""))
								if (Len(tempType)>0) And (InStr(1,tempType,".")=0) Then
									tempType = ReplaceEx(tempType, Array("""","", "=",""))
									tempType = GetTypeFromVid(tempType)
									if Len(tempType)>0 Then
										rezult = tempType
										Exit For
									End If
								End If
							End If
						End If
					End If
				End If

				'���� � ��������
				Patern = "(\s|^)*[^.][" & cnstRExWORD & "]+[\(]+[^@]+[\)]+(\s)*(;)+" ' "
				tempType = FindInStrEx(Patern, ttext)
				if Len(tempType)>0 Then
					tempType = AnalizProcedure(tempType,ResultWord)
					if Len(tempType)> 0 Then
						Rezult = tempType
						Exit For
					End If
				End If

				'�������� ��� ��������� ���������� �������� �������������� �� "����������" ����������/�������
				'UniMetodsAndRekv
				Patern = "(\s|;)+(" & ResultWord &  ")+[\s]*[.]+[\s]*[" & cnstRExWORD & "]+[\s]*[\(|;|=|\.|<]+" ' "
				tempTypeAdd = FindInStrEx(Patern, ttext)
				if Len(tempTypeAdd)>0 Then
					Patern = "(\s|;)+(" & ResultWord &  ")+[\s]*[.]+[\s]*[" & cnstRExWORD & "]+" ' "
					tempTypeAdd = FindInStrEx(Patern, tempTypeAdd)
					if Len(tempTypeAdd)> 0 Then
						Patern = "[.]+[\s]*[" & cnstRExWORD & "]+" ' "
						tempTypeAdd = FindInStrEx(Patern, tempTypeAdd)
						if Len(tempTypeAdd)> 0 Then
							Patern = "[" & cnstRExWORD & "]+" ' "
							tempTypeAdd = FindInStrEx(Patern, tempTypeAdd)
							tempTypeAdd = lcase(Trim(tempTypeAdd))
							If UniMetodsAndRekv.Exists(tempTypeAdd) Then
								RezTempTypeAdd = UniMetodsAndRekv.Item(tempTypeAdd)
								glRWord.TypeVid2 = RezTempTypeAdd
							End If
						End If
					End If
				End If


				'���� ��������� ���� "������.���() = "������������"" ����� �������� ����������/��������/������� ������ ����
				'Patern = "(\s|^)*(" & ResultWord & ")+(\s)*[.]+(\s)*(���|Kind)+(\s)*\((\s)*\)(\s)*(=|<>)+(\s)*[""][" & cnstRExWORD & "]+[""]"
				Patern = "(\s|^)*(" & ResultWord & ")+(\s)*[.]+(\s)*(���|Kind)+\s*\(+\s*\)+\s*(=|<>)+\s*[""][" & cnstRExWORD & "]+[""]"
				tempType = FindInStrEx(Patern, ttext)
				if (Len(tempType) = 0) And (Len(AddWords)>0) Then
					Patern = "(\s|^)*(" & ResultWord & "." & AddWords & ")+(\s)*[.]+(\s)*(���|Kind)+\s*\(+\s*\)+\s*(=|<>)+\s*[""][" & cnstRExWORD & "]+[""]"
					tempType = FindInStrEx(Patern, ttext)
				End IF
				if Len(tempType)> 0 Then
					Patern = "(=|<>)+\s*[""][" & cnstRExWORD & "]+[""]"
					tempType = FindInStrEx(Patern, tempType)
					if Len(tempType)> 0 Then
						Patern = "[""][" & cnstRExWORD & "]+[""]"
						tempType = FindInStrEx(Patern, tempType)
						if Len(tempType)> 0 Then
							tempType = Trim(tempType)
							tempType = Replace(tempType,"""","")
							ArrTempType = Split(tempType,vbCrLf)
							tempType = ArrTempType(0)
							ArrTipes = Array("��������","����������","�������")
							For tttt = 0 To 2
								Set MDObjekts = Metadata.TaskDef.Childs(CStr(ArrTipes(tttt)))
								For cntdoc = 0 To MDObjekts.Count - 1
									Set MDObj = MDObjekts(cntdoc)
									if (UCase(MDObj.Name) = UCase(tempType)) Then
										Rezult = ArrTipes(tttt) & "." & tempType
										if Len(AddWords)>0 Then
											tRezultStr = GetTypeVid(AddWords,Rezult, Doc,0,"")
										End If
										if Len(tRezultStr)>0 Then
											GetTypeFromTextRegExps	= tRezultStr
										Else
											GetTypeFromTextRegExps	= Rezult
										End If
										AddWords = ""
										Exit Function
									End If
								Next
							Next
						End If
					End If
				End If

				'���� ��������� ���� "���� ��������������(�������������������) = "��������������" �����
				Patern = "(\s)+(��������������|ValueTypeStr)+[\s]*[\(]+[\s]*(" & ResultWord & ")+[\s]*[\)]+[\s]*[=]+[\s]*[""]+(��������������|���������������|����������|��������|�������)+[""]+[\s]+"
				tempType = FindInStrEx(Patern, ttext)
				if Len(tempType)> 0 Then
					tRezultStr = ""
					If InStr(1,lcase(tempType),lcase("��������������"))>0 Then
						tRezultStr	= cnstNameTypeSZ
					ElseIf InStr(1,lcase(tempType),lcase("���������������"))>0 Then
						tRezultStr	= cnstNameTypeTZ
					ElseIf InStr(1,lcase(tempType),lcase("����������"))>0 Then
						tRezultStr	= "����������"
					ElseIf InStr(1,lcase(tempType),lcase("��������"))>0 Then
						tRezultStr	= "��������"
					ElseIf InStr(1,lcase(tempType),lcase("�������"))>0 Then
						tRezultStr	= "�������"
					End IF
					if Len(tRezultStr)>0 Then
						glRWord.TypeVid = tRezultStr
						'2005 02 07 ���� �����, ���������� ����� �������� ���� ���....
						'Exit Function
					End If
				End If

				'���� ��������� ���� "������ = ����.���(); ����� �������� ����������/��������/������� ������ ����
				' � AddText � ��� ����� ��������� � ������ � ������� "[����������=���������]����������������"
				' ���������� - ���������� �������� �� ����� ���� "������ = ����.���();
				'Patern = "(\s|^)*(" & ResultWord & ")+(\s)*[.]+(\s)*(���|Kind)+(\s)*\((\s)*\)(\s)*[=](\s)*[""][" & cnstRExWORD & "]+[""]"
				If Len(AddText)>0 Then
					Patern = "[\s]+[" & cnstRExWORD & "]+[\s]*[=]+[\s]*(" & ResultWord & ")+[\s]*[.]+[\s]*(���|Kind)+[\s]*[\(]+[\s]*[\)]+"
					tempType = FindInStrEx(Patern, ttext)
					if Len(tempType)> 0 Then
						Patern = "[\s]+[" & cnstRExWORD & "]+[\s]*[=]+"
						tempType = FindInStrEx(Patern, tempType)
						if Len(tempType)> 0 Then
							Patern = "[" & cnstRExWORD & "]+"
							tempType = FindInStrEx(Patern, tempType)
							if Len(tempType)> 0 Then
								' ��� � ��� ���������� ���������� ��� �������
								tempType = Trim(tempType)
								ArrAddText = Split(AddText,vbCrLf)
								If Ubound(ArrAddText)<>-1 Then
									For qqq = 0 To Ubound(ArrAddText)
										Vids = Split(ArrAddText(qqq),"=")
										If UBound(Vids)>0 Then
											if LCase(Vids(0)) = LCase(tempType) Then
												Vid = Vids(1)
												ArrTipes = Array("��������","����������","�������")
												For tttt = 0 To 2
													Set MDObjekts = Metadata.TaskDef.Childs(CStr(ArrTipes(tttt)))
													For cntdoc = 0 To MDObjekts.Count - 1
														Set MDObj = MDObjekts(cntdoc)
														if (UCase(MDObj.Name) = UCase(Vid)) Then
															Rezult = ArrTipes(tttt) & "." & Vid
															If Len(AddWords)>0 Then
																Rezult	= GetTypeVid(AddWords,Rezult, Doc,0,"")
															End If
															If Len(Rezult)>0 Then
																AddWords = ""
																GetTypeFromTextRegExps	= Rezult
																Exit Function
															End If
														End If
													Next
												Next
												exit For
											End If
										End If
									Next
								End If
							End If
						End If
					End If
				End If


				Patern = "(\s|^)(���������|�������|Procedure|Function)(\s)+[a-zA-z�-��-�0-9_]+\([a-zA-z�-��-�0-9_,=""\s]*[\)]*"
				ProcFunc = FindInStrEx(Patern, ttext)
				If Len(ProcFunc)>0 Then
				'���� �� ����� ���� ������, ��������� ���������������� �� ���������
				'�� ���������� � ����� ����������� "��������� ���������������(������������)"
				'���� �����, ����� ������� ������ ��������� ������� ()
					Patern = "(\s|^)(���������|�������|Procedure|Function)(\s)+(���������������)+(\s)*\((\s)*("&ResultWord&")(\s)*[\)]+"
					ProcFunc = FindInStrEx(Patern, ttext)
					If Len(ProcFunc)>0 Then
						ThisDoc = Split(Doc.name,".") 					'����� �������� ������ ��������� (���� �� ����)
						if UBound(ThisDoc)>0 Then
							Rezult = Trim(GetListOsnovaniyDoca(ThisDoc(1)))
						End If
						if InStr(1,Rezult," ") = 0 Then			'������� ���� # ��� ��� �������� ���, ������ �������� ����.
							Rezult = Mid(Rezult,2)
						End If
					End If
					if nctStep = 0 Then
						Exit For
					End If
				End If

				Patern	 = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*(�������������|CreateObject)+[\s]*[\(]+["& cnstRExWORD &"\.""]+[\)]+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					RezultStr = FindInStrEx("\((.*?)\)", RezultStr)
					if Len(RezultStr)>0 Then
						If (InStr(1,RezultStr,"""")>0) And (InStr(1,RezultStr,"+") = 0) Then
							Rezult = FindInStrEx("\""(.*?)\""", RezultStr)
							Rezult = Replace(Rezult,"""","")
							glRWord.AsCreateObj = true
							if Len(AddWords)>0 Then
								RezultStr = GetTypeVid(AddWords,Rezult, Doc,0,"")
								if Len(RezultStr)<>0 Then
									AddWords = ""
									Rezult = RezultStr
								End if
							End if
							'If lcase(Rezult) = "��������" Then :	Rezult = "�������"	: End if

							Exit For
						End if
					End if
				End if

				Patern	 = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*(���������������|CurrentDocument|��������������|CurrentItem)+[\s]*[\(]+[\s]*[\)]+[\s]*[;]+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					ArrRezultStr = split(doc.Name,".")
					if UBound(ArrRezultStr)>0 Then
						Rezult = ArrRezultStr(0) & "." & ArrRezultStr(1)
					End if
				End if

				If (InStr(1, Doc.Name,	"������.") = 1) 	Then
					Patern	 = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*(���������������|CurrentDocument)+[\s]*[\(]*[\s]*[\)]*[\s]*[;]+"
					RezultStr = FindInStrEx(Patern, ttext)
					if (Len(RezultStr)> 0) And (InStr(1,RezultStr,"(") =0 ) Then
						Rezult = "��������"
					End if
				End IF

				RezultStr2 = FindInStrEx(Patern, ttext)

				'�����������.��������������(����������);
				Patern	 =	"[\s|^|;]+[" & cnstRExWORD & "]+[\s]*[\.]+[\s]*(��������������|RetrieveTotals)+[\s]*[\(]+[\s]*(" & ResultWord & ")+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					Patern	 = "[\s]*[\.]+[\s]*(��������������|RetrieveTotals)+[\s]*[\(]+[\s]*(" & ResultWord & ")+"
					AddRezultStr = FindInStrEx(Patern, RezultStr)
					if Len(AddRezultStr)>0 Then
						RezultStr = Replace(RezultStr,AddRezultStr,"")
						Patern = "[" & cnstRExWORD & "]+"
						RezultStr = FindInStrEx(Patern, RezultStr)
						tRezultStr = AddWords
						RezultStr = GetTypeFromTextRegExps(RezultStr, AddWords)
						if Len(RezultStr)>0 Then
						IF Len(tRezultStr) = 0 Then
							Rezult = RezultStr & "+���������������"
						Else
							Rezult = RezultStr
						End if
							glRWord.BTMeth = "��������������"
							glRWord.AddWord = ""
							if Len(AddWords)>0 Then
								AddWords = ""
							End if
							Exit For
						End if

					End if
				End if

				Patern	 = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*[" & cnstRExWORD & "]+[\s]*[;]+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					Patern	 = "[=]+[\s]*[" & cnstRExWORD & "]+[\s]*[;]+"
					RezultStr = FindInStrEx(Patern, ttext)
					if Len(RezultStr)> 0 Then
						RezultStr = ReplaceEx(RezultStr,Array( "=","", ";",""))
						RezultStr = Trim(RezultStr)
						tRezultStr = GetTypeVid(RezultStr,"", Doc, 0,RezultStr)
						'tRezultStr = GetTypeFromTextRegExps(RezultStr, AddWords)
						if Len(tRezultStr)>0 Then
							Rezult = tRezultStr
							Exit For
						End if
					End if
				End if

				Patern	 = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*[" & cnstRExWORD & "\s\.\(\)]+[\s]*[;]+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					Patern	 = "[=]+[\s]*[" & cnstRExWORD & "\s\.\(\)]+[\s]*[;]+"
					RezultStr = FindInStrEx(Patern, ttext)
					if Len(RezultStr)> 0 Then
						RezultStr = ReplaceEx(RezultStr, Array("=","", ";","", vbTab," ", " ",""))
						RezultStr = Trim(RezultStr) '"����������.���������������()"
						' ���������� ���� ���� "��� = ���.��������.��������������();" � ��� = ���.��������;; ��� "���"..
						RezultStr = Replace(RezultStr,FindInStrEx("[\.]+\s*(��������������|���������������)+\s*\(+\s*\)+", ttext),"")
						ArrRezultStr = Split(RezultStr,".")
						IF UBound(ArrRezultStr) = 1 Then
							tRezultStr = GetTypeVid(ArrRezultStr(1),ArrRezultStr(0), Doc, 0,RezultStr)
							if Len(tRezultStr)>0 Then
								Rezult = tRezultStr
								Exit For
							End IF
						ElseIf UBound(ArrRezultStr) = 0 Then
							tRezultStr = GetTypeVid(ArrRezultStr(0),"", Doc, 0,RezultStr)
							if Len(tRezultStr)>0 Then
								Rezult = tRezultStr
								Exit For
							End IF
						End IF
						IF UBound(ArrRezultStr)>0 Then
							tRezultStr = GetTypeVid(ArrRezultStr(0),"", Doc, 0,RezultStr)
							IF Len(tRezultStr) = 0 Then
								Patern = ""
								for qwqe = 1 To UBound(ArrRezultStr)
									IF Patern = "" Then
										Patern = ArrRezultStr(qwqe)
									Else
										Patern = Patern & "." & ArrRezultStr(qwqe)
									End if
								Next
								tRezultStr = GetTypeVid(AddWords,RezultStr, Doc,0,"")
								if Len(tRezultStr) = 0 Then
									tRezultStr = GetTypeFromTextRegExps(ArrRezultStr(0), Patern)
									tRezultStr0 = tRezultStr
									If (lcase(tRezultStr) = "��������") And (Len(Patern)>0) Then
										tRezultStr = "�������" &"."&Patern
									elseif Len(Patern)>0 Then
										if Instr(1,Patern,".")>0 Then
											ArrWOfVariable = Split(Patern,".")
											for q = 0 To UBound(ArrWOfVariable)
												tRezultStr = GetTypeVid(ArrWOfVariable(q),tRezultStr, Doc,0,"")
											Next
										else
											tRezultStr = GetTypeVid(Patern,tRezultStr, Doc,0,"")
										End if
									End if
									Patern = ""
									if Len(AddWords)>0 Then
										if Instr(1,AddWords,".")>0 Then
											ArrWOfVariable = Split(AddWords,".")
											for q = 0 To UBound(ArrWOfVariable)
												tRezultStr = GetTypeVid(ArrWOfVariable(q),tRezultStr, Doc,0,"")
											Next
										else
											tRezultStr = GetTypeVid(AddWords,tRezultStr, Doc,0,"")
										End if
									End if
								End if
								AddWords = ""
								if Len(tRezultStr)>0 Then
									Rezult = tRezultStr
									Exit For
								End if

								Patern = ""
							ElseIf (UC(tRezultStr) = UC("���������������")) And ( UC(ArrRezultStr(1)) = UC("��������������()")) Then
								tRezultStr = ArrRezultStr(0) & "#����������������������#"
							End if

						End if
						if Len(tRezultStr)>0 Then
							if InStr(1,tRezultStr,".")>0 Then
								tRezultStr = GetTypeVid(ArrRezultStr(1),tRezultStr, Doc, 0,"")
								if Len(tRezultStr)<>0 Then
									Rezult = tRezultStr
								End if
							Else
								Rezult = tRezultStr
							End if
							if Len(Rezult)>0 Then
								Exit For
							End if
						End if
					End if
				End if

				Patern = "[\s|^]+(" & ResultWord & ")+[\s]*[.]+[\s]*(���������������|SourceTable)+[\s]*[\(]+[\s]*[" & cnstRExWORD & """]+[\)]+[\s]*[;]+[\s]*"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					Patern = "[\(]+[\s]*[" & cnstRExWORD & """]+[\)]+"
					RezultStr = FindInStrEx(Patern, RezultStr)
					if Len(RezultStr)>0 Then
						Patern = "[" & cnstRExWORD & """]+"
						RezultStr = FindInStrEx(Patern, RezultStr)
						if Len(RezultStr)>0 Then
							RezultStr = Replace(RezultStr,"""","")
							AddWords = RezultStr
						End if

					End if
				End if

				'Patern = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*[" & cnstRExWORD & ".]+[\s$;]+"
				'RezultStr = FindInStrEx(Patern, ttext)


				Patern = "[\s|^|;]+(" & ResultWord & ")+[\s]*[=]+[\s]*[" & cnstRExWORD & ".]+[\s$;]+"
				RezultStr = FindInStrEx(Patern, ttext)
				if Len(RezultStr)> 0 Then
					RezultStr = FindInStrEx("[\s]*[=]+[\s]*[" & cnstRExWORD & ".]+", RezultStr)
					if Len(RezultStr)>0 Then
						RezultStr = FindInStrEx("[" & cnstRExWORD & ".]+", RezultStr)
						if Len(RezultStr)>0 Then
							IF InStr(1,RezultStr,".") > 0 Then
								tmpArrVidType = Split(Trim(RezultStr),".")
								if UBound(tmpArrVidType)>0 Then
									tmpType	 = UCase(Trim(tmpArrVidType(0)))
									If Not Is1CObject(tmpType) Then
										AddWords = ""
										if UC(tmpType) <> UC(RezultStr) Then
											tRezultStr = GetTypeFromTextRegExps(tmpType,AddWords)
											If Is1CObject(tRezultStr) Then
												Rezult = Trim(tRezultStr)
												For uu = 1 To UBound(tmpArrVidType)
													Rezult = Rezult & "." & tmpArrVidType(uu)
												Next
											Else
												Rezult = RezultStr
											End if
										End if
									else
										Rezult = RezultStr
									End if
								End if
							End if
						End if
					End if
				End if
			End if
			FindRegim = 1  ' ������ = "�����������"
			patrn = "(\s|\()+[" & cnstRExWORD & "]+[\s]*[=]+[\s]*[""]+[" & cnstRExWORD & "]+[""]+"
			ttext0 = FindInStrEx (patrn, ttext)
			if Len(ttext0) = 0 Then
				patrn = "(\s|\()+[""]+[" & cnstRExWORD & "]+[""]+[\s]*[=]+[\s]*[" & cnstRExWORD & "]+"
				ttext0 = FindInStrEx (patrn, ttext)
				FindRegim = 2 '"�����������" = ������
			End if
			if Len(ttext0)>0 Then
				ttext0 = ReplaceEx(ttext0,Array(" ","", vbTab,"", "(","",  """",""))
				Arrttext0 = Split(ttext0,"=")
				If UBound(Arrttext0)=1 Then
					IF FindRegim = 2 Then
						ttext0 = Arrttext0(1)&"="&Arrttext0(0)
					End if
					If Len(AddText) = 0 Then
						AddText = ttext0
					Else
						AddText = AddText & vbCrLf & ttext0
					End if
				End if
			End if
		Next
		if Len(Rezult)>0 Then
			Exit For
		End if
	Next

	if Len(Rezult)>0 Then
		if Mid(Rezult,Len(Rezult),1) = "." Then
			Rezult = Mid(Rezult,1,Len(Rezult)-1)
		End If
	Else
		if Len(RezTempTypeAdd)>0 Then
			Rezult = RezTempTypeAdd
		End If
	End If
	Rezult = Replace(Rezult,"""","")
	IF InStr(1,Rezult,".")> 0 Then
	End If
	GetTypeFromTextRegExps	= Rezult
End Function

Private Function Is1CObject(NameObj)
	Is1CObject	 = False
	If UCase(NameObj) = UCase("�������") Then
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("��������") Then
		NameObj =	"�������"
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("����������") Then
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("�����") Then
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("��������") Then
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("������������") Then
		Is1CObject	 = True
	ElseIf UCase(NameObj) = UCase("���������") Then
		Is1CObject	 = True
	End If
End Function


Private Function AnalizProcedure(tempType, RezultWord)
	AnalizProcedure = ""
	Patern = "[^\s][^.][" & cnstRExWORD & "]+[\(]+" ' "
	NamePF = FindInStrEx(Patern, tempType)
	if Len(NamePF)>0 Then
		Patern = "[\(]+[^@]+[\)]+" ' "
		ParametrsPF = FindInStrEx(Patern, tempType)
		ParametrsPF = Replace(ParametrsPF," ","")
		ParametrsPFArr = split(ParametrsPF,",")
		NamePF = Trim(Replace(NamePF,"(",""))
		IF UCase(NamePF) = UCase("�������������") Then
			If UBound(ParametrsPFArr)> 2 Then
				If UCase(ParametrsPFArr(2)) = UCase(RezultWord) Then
					if InStr(1,UCase(ParametrsPFArr(0)),UCase("����������")) >0 Then
						AnalizProcedure	= "����������������������"
						Exit Function
					End If
				End If
			End If
		End If
	End If
End Function

'ttextModuleGM = ""
'ttextModuleGM_PriNachRabSys = ""
'�������� ��������, ��������
Private Function GetColumnsFromTZ(ResultWord, STipom)
	GetColumnsFromTZ = ""
' 	If Windows.ActiveWnd Is Nothing Then
' 		Exit Function
' 	End If
	OldGlobalStr = GlobalStr
	GetColumnsFromTZ = GetColumnsFromTZA(ResultWord, STipom, ttextModuleGM)
	If Len(GetColumnsFromTZ)<>0 Then
		Exit Function
	End If
	GetColumnsFromTZ = GetColumnsFromTZA(ResultWord, STipom, ttextModuleGM_PriNachRabSys)
	If Len(GetColumnsFromTZ)<>0 Then
		Exit Function
	End If
	TextAll = ""
	'ArrProcFuncFromTheWord	= ""
	'ArrProcFuncNumbParam	= ""
	'Dim GlobalModule '������ TheModule ��� ����������� ������
	'Dim LocalModule	 '������ TheModule ��� ������ �������� �������



	'��������������� �� ���������� � ��������� ������� � ������� �������� ������� ������� ��������..
	if Len(glRWord.ArrProcFuncFromTheWord)>0 Then
		ArrProcFuncFromTheWord = Split(glRWord.ArrProcFuncFromTheWord,",")
		ArrProcFuncNumbParam = Split(glRWord.ArrProcFuncNumbParam,",")
		For ee=0 To UBound(ArrProcFuncFromTheWord)
			If GlobalModule.IsProcedure(ArrProcFuncFromTheWord(ee),"","","",TextAll, TheProc) Then
				On Error Resume Next
				'if (ArrProcFuncNumbParam(ee) = "-1")
				If (TheProc.TypeItem = 2) And (ArrProcFuncNumbParam(ee) = "-1") Then
					GetColumnsFromTZ = GetColumnsFromTZA(TheProc.RetValueStr, STipom, TheProc.Text)
				Else
					ParamNumber = TheProc.GetParamNumber(ArrProcFuncNumbParam(ee))
					GetColumnsFromTZ = GetColumnsFromTZA(ParamNumber, STipom, TheProc.Text)
				End If
				If Len(GetColumnsFromTZ)<>0 Then
					Exit Function
				End If
			End If
		Next
	End If
	GlobalStr = OldGlobalStr
	doc = ""
	if Not CheckWindow(doc) then Exit Function

	'����������� ��������� ����� �� ����� � ����� ����
	virajenieCO1 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_.]*[\s]*(�����������������������|UnloadTable)[\s]*[(][\s]*" & ResultWord &"[\s]*[)](\s|^|;)"
	'���� ����� ���������� ������ ���� //������������=��������.������� ��� //������������=�������.�������
	virajenieCO3 = 	"(//)+[\s]*(" & ResultWord &")[\s]*[=]+[\s]*(��������|�������)+[\s]*[.]+[\s]*[a-zA-Z�-��-�0-9_.]+"
	TextAll = ""
	'���� �����, ���� �� � ���������� ������, ����� �������� ������ ����� ������ ���������
	'�����, ������� �������� ������ ����� ���������, �����, ���� �� ��������, �����
	'���� ����� � ������ � �� �����
	'����� �������� ���������� � ��������� ������� � ���� �������
	Rezult = ""
	' ��������� � 2 �������, ���� � ������� ����������� �� ������, ������ �� ���� ������
	EndPosotion = 1
	If Windows.ActiveWnd.Document.Name = "���������� ������" Then
		EndPosotion = 0
	End If
	For yyy = 0 To EndPosotion
		FStartPosition = 0
		IF yyy = 0 Then
			FStartPosition = GlobalStr
		Else
			FStartPosition = Doc.LineCount
		End IF
		For i = FStartPosition To 0 Step -1
			ttext = Trim(Doc.range(i))
			GlobalStr = i
			TextAll = ttext & TextAll
			if inStr(1,ttext,"//")>0 Then
				ttext = Mid(ttext,1,inStr(1,ttext,"//"))
			End If
			IF yyy = 0 Then
				If Len(FindInStrEx("(\s|^)+(���������|�������|Procedure|Function)", ttext))>0 Then
					Exit For
				End If
			End If
			If InStr(1,UC(ttext),UC(ResultWord)) > 0 Then

				'����������� ������ � ������� ��������
				virajenieCO0 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*(���������|Unload)+(\s)*[\(]+(\s)*(" & ResultWord &")+.*[\)]+(\s|^|;)"
				RezultStr = FindInStrEx(virajenieCO0, ttext)
				if Len(RezultStr)> 0 Then
					'�������� ���������� �������
					textVarZapros = FindInStrEx("(\s)*[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*",RezultStr)
					'����� ����� - ������ ��������� ����� ���� � �� �������� :(
					' ���������............
					If Len(textVarZapros)>0 Then
						'��������
						textVarZapros = FindInStrEx("[a-zA-Z�-��-�0-9_]+",textVarZapros)
					End IF
					Patern = "[\(]+(\s)*(" & ResultWord &")+.*[\)]+"
					ttt = FindInStrEx(Patern,RezultStr)
					ttt = Replace(Replace(Replace(ttt,"(",""),")","")," ","")
					if InStr(1,ttt,",")>0 Then
						eeee = Split(ttt,",")
						If Ubound(eeee)<> -1 Then
							ttt = eeee(0)
						End if
					End if
					If UC(ttt) = UC(ResultWord) Then
						TypeVid = GetTypeFromTextRegExps(textVarZapros,"")
						'Stop
						if InStr(1,TypeVid,"+���������������")>0 Then
							TypeVid = Replace(TypeVid,"+���������������","")
						End if
						RezultStr = ""
						if (UC(TypeVid) = "������") Then
							RezultStr = GetVariableAndFunctionZapros(textVarZapros, STipom)
						Elseif (UC(TypeVid) = UC(cnstNameTypeTZ)) Then
							RezultStr = GetColumnsFromTZ(textVarZapros, STipom)
						ElseIf Instr(1,Lcase(TypeVid),LCase("�������."))>0 Then
							eeee = Split(TypeVid,".")
							RezultStr = GetStringRekvizitovFromObj(eeee(0), eeee(1),0,0, "",	0, glRWord)
							'GetStringRekvizitovFromObj(TypeObj, NameObj,SMetodami,AsRecvOfForms, BTMeth,	BTNumberParams, RWord)
						End if
						if Len(RezultStr)>0 Then
							Rezult = Rezult & vbCrLf &  RezultStr
						End if
						Exit For
					End if
				End if

				'����������� ����� �� �������� � ������� ��������
				virajenieCO0 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*(��������������|RetrieveTotals)+(\s)*[\(]+(\s)*(" & ResultWord &")+(\s)*(,)*[" & cnstRExWORD & "\(\)]*(,)*[0-3]*(,)*(\s)*[\)]+(\s|^|;)"
				RezultStr = FindInStrEx(virajenieCO0, ttext)
				if Len(RezultStr)> 0 Then
					'�������� ���������� �������
					textVarRegistr = FindInStrEx("(\s)*[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*",RezultStr)
					If Len(textVarRegistr)>0 Then
						'��������
						textVarRegistr = FindInStrEx("[a-zA-Z�-��-�0-9_]+",textVarRegistr)
					End IF
					If Len(textVarRegistr)>0 Then
						AddWords = ""
						ttipvid = GetTypeFromTextRegExps(textVarRegistr,AddWords)
						if Len(ttipvid)>0 Then
							Arrttipvid = Split(ttipvid,".")
							if UBound(Arrttipvid) = 1 Then
								If STipom = 1 Then
									RezultStr = GetStringRekvizitovFromObj(Arrttipvid(0), Arrttipvid(1),2,0,"",0)
								Else
									RezultStr = GetStringRekvizitovFromObj(Arrttipvid(0), Arrttipvid(1),0,0,"",0)
								End IF
							End IF
						End IF
					End IF
					if Len(RezultStr)>0 Then
						Rezult = Rezult & vbCrLf &  RezultStr
					End if
					Exit for
				End if

				'���� ����� ���������� ������� � ��
				virajenieCO2 = 	"(\s|^|;)+(" & ResultWord &")[\s]*[.]+[\s]*(���������������|InsertColumn|������������|NewColumn)\((.*?)\)"
				RezultStr = FindInStrEx(virajenieCO2, ttext)
				if Len(RezultStr)> 0 Then
					RezultStr = FindInStrEx("\((.*?)\)", RezultStr)
					if Len(RezultStr)>0 Then
						If (InStr(1,RezultStr,"""")>0) And (InStr(1,RezultStr,"+") = 0) And (InStr(1,RezultStr,")") <> 0) And (InStr(1,RezultStr,"(") <> 0) Then
							RezultStr = ReplaceEx(RezultStr,Array("(","", ")","", """",""))
							tttt = split(RezultStr,",")
							if UBound(tttt) > 0 Then
								'� ������� (0) ����� ���, � (1) ����� ���
								IF STipom = 1 Then
									RezultStr = tttt(0) & "=" & tttt(1)
								Else
									RezultStr = tttt(0) '& "#" & tttt(1)
								End if
							elseif UBound(tttt) = 0 Then
								RezultStr = tttt(0)
							else
								RezultStr = ""
							End if
							if (Len(Rezult)>0) And (Len(RezultStr)>0) Then
								Rezult = Rezult & vbCrLf & RezultStr
							ElseIF (Len(RezultStr)>0) Then
								Rezult = RezultStr
							End if
						End if
					End if
				End if
				'����������� ��������� ����� �� ����� � ����� ����
				'virajenieCO1 = 	"(\s|^|;)*[a-zA-Z�-��-�0-9_.]*[\s]*(�����������������������|UnloadTable)[\s]*[(][\s]*" & ResultWord &"[\s]*[)](\s|^|;)"
				RezultStr = FindInStrEx(virajenieCO1, ttext)
				if Len(RezultStr)> 0 Then
					'����� ������ � ��������� �������� ��������� �����
					VidDoca = Array("��������")
					RezultStr = FindInStrEx("(\s|^|;)+[a-zA-Z�-��-�0-9_]+[\s]*[.]+[\s]*(�����������������������|UnloadTable)", RezultStr)
					if Len(RezultStr)>0 Then
						RezultStr = FindInStrEx("[a-zA-Z�-��-�0-9_]+[\s]*[.]+", RezultStr)
						RezultStr = Replace(RezultStr,".","")
						RezultStr = Trim(RezultStr)
						AddWords = ""
						if Len(RezultStr)>0 Then
							OldStartScanLineText = StartScanLineText
							StartScanLineText = i
							tttVidDoca = GetTypeFromTextRegExps(RezultStr,AddWords)
							StartScanLineText = OldStartScanLineText
							if InStr(1,tttVidDoca,".") > 0 Then
								VidDoca = split(tttVidDoca, ".")
							End if
						End if
					Else
						VidDoca = split(Doc.Name, ".")
					End if
					if UBound(VidDoca)>0 Then
						IF ObjectExist("��������",VidDoca(1)) Then
							Set MetaDoc = MetaData.TaskDef.Childs(CStr("��������"))(CStr(VidDoca(1)))
							Set Table = MetaDoc.Childs("����������������������")
							for tbl_cnt = 0 To Table.Count - 1
								set tbl_chld = Table(tbl_cnt)
								if Len(Rezult)>0 Then
									Rezult = Rezult &  vbCrLf & tbl_chld.Name
								else
									Rezult = tbl_chld.Name
								End if
								if STipom = 1 Then
									Rezult = Rezult & "=" & tbl_chld.Type.FullName
								End if
							Next
							Exit for
						End if
					End if
				End if
			End if
		Next
		RezultStr = FindInStrEx(virajenieCO3, TextAll)
		if Len(Rezult) > 0 Then
			Exit For '����� ��� ��������� ������ ������?
		End If
	Next '�������

	RezultStr = FindInStrEx(virajenieCO3, TextAll)
	if Len(RezultStr) > 0 Then
		IF InStr(1,RezultStr,vbCrLf) Then
			nnnn = Split(RezultStr,vbCrLf)
			If Len(nnnn(0))>0 Then
				RezultStr = nnnn(0)
			End IF
		End IF
		patern = 	"[\s|^]+(�������|��������)+[\s]*[.]+[\s]*[a-zA-Z�-��-�0-9_.]+"
		RezultStr = FindInStrEx(patern, RezultStr)
		if Len(RezultStr)>0 Then
			VidDoca = split(UCase(RezultStr), ".")
			if UBound(VidDoca)>0 Then
				if (VidDoca(0) = UCase("��������")) Then
					Set MetaDoc = MetaData.TaskDef.Childs(CStr("��������"))(CStr(VidDoca(1)))
					Set Table = MetaDoc.Childs("����������������������")
		 			for tbl_cnt = 0 To Table.Count - 1
		 				set tbl_chld = Table(tbl_cnt)
		 				if Len(Rezult)>0 Then
		 					Rezult = Rezult &  vbCrLf & tbl_chld.Name
		 				else
		 					Rezult = tbl_chld.Name
		 				End if
		 			Next
			 	ElseIF (VidDoca(0) = UCase("�������")) Then
			 		Rezult = GetStringRekvizitovFromObj("�������", CStr(VidDoca(1)),0,0,"",0)
				End if
			End if
		End If
	End if


	'������������ �� ������������
	if STipom = 0 Then
		if 1=2 Then
			Patern = "(" & ResultWord & ")+[\s]*[.]+[\s]*(��������|GroupBy)+[\s]*[\(]+[\s]*[" & cnstRExWORD & "\""\s,]+[\s]*[\)]+"
			ttt = FindInStrEx(Patern,TextAll)
			IF Len(ttt)>0 Then
				Patern = "[\(]+[\s]*[" & cnstRExWORD & "\""\s,]+[\s]*[\)]+"
				ttt = FindInStrEx(Patern,ttt)
				IF Len(ttt)>0 Then
					Rezult = ReplaceEx(ttt,Array("(","", ")","", """","", vbCrLf,",", ",,",","))
				End if
			End if
		End if
		ArrRezult = Split(Rezult,vbCrLf)
		tmpRezult = ""
		if UBound(ArrRezult)<> -1 Then
			For ee = 0 To UBound(ArrRezult)
				patern = "(\s|^|,)+(" & ArrRezult(ee) & ")+(\s|^|,|$)+"
				if Len(FindInStrEx(patern,tmpRezult)) = 0 Then
					 If Len(tmpRezult)>0 Then
						tmpRezult = tmpRezult & "," & ArrRezult(ee)
						else
						tmpRezult = ArrRezult(ee)
					End IF
				End IF
			Next
		End IF
		Rezult = Replace(tmpRezult,",",vbCrLf)
	End IF

	GetColumnsFromTZ	= Rezult
End Function

'ttextModuleGM = ""
'ttextModuleGM_PriNachRabSys = ""
'�������� ��������, ��������
Private Function GetColumnsFromTZA(ResultWord, STipom, textModule)
	GetColumnsFromTZA = ""
' 	If Windows.ActiveWnd Is Nothing Then
' 		Exit Function
' 	End If
	ArrtextModule = split(textModule, vbCrLf)
	'����������� ��������� ����� �� ����� � ����� ����
	virajenieCO1 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_.]*[\s]*(�����������������������|UnloadTable)[\s]*[(][\s]*" & ResultWord &"[\s]*[)](\s|^|;)"
	'���� ����� ���������� ������ ���� //������������=��������.������� ��� //������������=�������.�������
	virajenieCO3 = 	"(//)+[\s]*(" & ResultWord &")[\s]*[=]+[\s]*(��������|�������)+[\s]*[.]+[\s]*[a-zA-Z�-��-�0-9_.]+"
	TextAll = ""
	'���� �����, ���� �� � ���������� ������, ����� �������� ������ ����� ������ ���������
	'�����, ������� �������� ������ ����� ���������, �����, ���� �� ��������, �����
	'���� ����� � ������ � �� �����
	'����� �������� ���������� � ��������� ������� � ���� �������
	Rezult = ""
	' ��������� � 2 �������, ���� � ������� ����������� �� ������, ������ �� ���� ������
	EndPosotion = 1
	If Windows.ActiveWnd.Document.Name = "���������� ������" Then
		EndPosotion = 0
	End If
	For i = UBound(ArrtextModule) To 0 Step -1
		ttext = Trim(ArrtextModule(i))
		GlobalStr = i
		TextAll = ttext & TextAll
		if inStr(1,ttext,"//")>0 Then
			ttext = Mid(ttext,1,inStr(1,ttext,"//"))
		End If
		If InStr(1,UC(ttext),UC(ResultWord)) > 0 Then

			'����������� ������ � ������� ��������
			virajenieCO0 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*(���������|Unload)+(\s)*[\(]+(\s)*(" & ResultWord &")+(\s)*(,)*[" & cnstRExWORD & "\(\)]*(,)*[0-3]*(,)*(\s)*[\)]+(\s|^|;)"
			RezultStr = FindInStrEx(virajenieCO0, ttext)
			if Len(RezultStr)> 0 Then
				'�������� ���������� �������
				textVarZapros = FindInStrEx("(\s)*[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*",RezultStr)
				If Len(textVarZapros)>0 Then
					'��������
					textVarZapros = FindInStrEx("[a-zA-Z�-��-�0-9_]+",textVarZapros)
				End IF
				Patern = "[\(]+(\s)*(" & ResultWord &")+(\s)*(,)*[" & cnstRExWORD & "\(\)]*(,)*[0-3]*(,)*(\s)*[\)]+"
				ttt = FindInStrEx(Patern,RezultStr)
				ttt = ReplaceEx(ttt,Array("(","",")",""," ",""))
				if InStr(1,ttt,",")>0 Then
					eeee = Split(ttt,",")
					If Ubound(eeee)<> -1 Then
						ttt = eeee(0)
					End if
				End if
				If UC(ttt) = UC(ResultWord) Then
					TypeVid = GetTypeFromTextRegExps(textVarZapros,"")
					if (UC(TypeVid) = "������") Then
						RezultStr = GetVariableAndFunctionZapros(textVarZapros, STipom)
					Elseif (UC(TypeVid) = UC(cnstNameTypeTZ)) Then
						RezultStr = GetColumnsFromTZA(textVarZapros, STipom,textModule)
					End if
					if Len(RezultStr)>0 Then
						Rezult = Rezult & vbCrLf &  RezultStr
					End if
					Exit For
				End if
			End if

			'����������� ����� �� �������� � ������� ��������
			virajenieCO0 = 	"(\s|^|;)+[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*(��������������|RetrieveTotals)+(\s)*[\(]+(\s)*(" & ResultWord &")+(\s)*(,)*[" & cnstRExWORD & "\(\)]*(,)*[0-3]*(,)*(\s)*[\)]+(\s|^|;)"
			RezultStr = FindInStrEx(virajenieCO0, ttext)
			if Len(RezultStr)> 0 Then
				'�������� ���������� �������
				textVarRegistr = FindInStrEx("(\s)*[a-zA-Z�-��-�0-9_]+(\s)*[.]+(\s)*",RezultStr)
				If Len(textVarRegistr)>0 Then
					'��������
					textVarRegistr = FindInStrEx("[a-zA-Z�-��-�0-9_]+",textVarRegistr)
				End IF
				If Len(textVarRegistr)>0 Then
					AddWords = ""
					ttipvid = GetTypeFromTextRegExps(textVarRegistr,AddWords)
					if Len(ttipvid)>0 Then
						Arrttipvid = Split(ttipvid,".")
						if UBound(Arrttipvid) = 1 Then
							RezultStr = GetStringRekvizitovFromObj(Arrttipvid(0), Arrttipvid(1),0,0,"",0)
						End IF
					End IF
				End IF
				if Len(RezultStr)>0 Then
					Rezult = Rezult & vbCrLf &  RezultStr
				End if
				Exit for
			End if

			'���� ����� ���������� ������� � ��
			virajenieCO2 = 	"(\s|^|;)+(" & ResultWord &")[\s]*[.]+[\s]*(���������������|InsertColumn|������������|NewColumn)\((.*?)\)"
			RezultStr = FindInStrEx(virajenieCO2, ttext)
			if Len(RezultStr)> 0 Then
				RezultStr = FindInStrEx("\((.*?)\)", RezultStr)
				if Len(RezultStr)>0 Then
					If (InStr(1,RezultStr,"""")>0) And (InStr(1,RezultStr,"+") = 0) And (InStr(1,RezultStr,")") <> 0) And (InStr(1,RezultStr,"(") <> 0) Then

						RezultStr = ReplaceEx(RezultStr,Array("(","", ")","", """",""))
						tttt = split(RezultStr,",")
						if UBound(tttt) > 0 Then
							'� ������� (0) ����� ���, � (1) ����� ���
							IF STipom = 1 Then
								RezultStr = tttt(0) & "=" & tttt(1)
							Else
								RezultStr = tttt(0) '& "#" & tttt(1)
							End if
						elseif UBound(tttt) = 0 Then
							RezultStr = tttt(0)
						else
							RezultStr = ""
						End if
						if (Len(Rezult)>0) And (Len(RezultStr)>0) Then
							Rezult = Rezult & vbCrLf & RezultStr
						ElseIF (Len(RezultStr)>0) Then
							Rezult = RezultStr
						End if
					End if
				End if
			End if
			'����������� ��������� ����� �� ����� � ����� ����
			'virajenieCO1 = 	"(\s|^|;)*[a-zA-Z�-��-�0-9_.]*[\s]*(�����������������������|UnloadTable)[\s]*[(][\s]*" & ResultWord &"[\s]*[)](\s|^|;)"
			RezultStr = FindInStrEx(virajenieCO1, ttext)
			if Len(RezultStr)> 0 Then
				'����� ������ � ��������� �������� ��������� �����
				VidDoca = Array("��������")
				RezultStr = FindInStrEx("(\s|^|;)+[a-zA-Z�-��-�0-9_]+[\s]*[.]+[\s]*(�����������������������|UnloadTable)", RezultStr)
				if Len(RezultStr)>0 Then
					RezultStr = FindInStrEx("[a-zA-Z�-��-�0-9_]+[\s]*[.]+", RezultStr)
					RezultStr = Replace(RezultStr,".","")
					RezultStr = Trim(RezultStr)
					AddWords = ""
					if Len(RezultStr)>0 Then
						tttVidDoca = GetTypeFromTextRegExps(RezultStr,AddWords)
						if InStr(1,tttVidDoca,".") > 0 Then
							VidDoca = split(tttVidDoca, ".")
						End if
					End if
				Else
					VidDoca = split(Doc.Name, ".")
				End if
				if UBound(VidDoca)>0 Then
					Set MetaDoc = MetaData.TaskDef.Childs(CStr("��������"))(CStr(VidDoca(1)))
					Set Table = MetaDoc.Childs("����������������������")
					for tbl_cnt = 0 To Table.Count - 1
						set tbl_chld = Table(tbl_cnt)
						if Len(Rezult)>0 Then
							Rezult = Rezult &  vbCrLf & tbl_chld.Name
						else
							Rezult = tbl_chld.Name
						End if
						if STipom = 1 Then
							Rezult = Rezult & " " & tbl_chld.Type.FullName
						End if
					Next
					Exit for
				End if
			End if
		End if
	Next
	RezultStr = FindInStrEx(virajenieCO3, TextAll)
	if Len(RezultStr) > 0 Then
		IF InStr(1,RezultStr,vbCrLf) Then
			nnnn = Split(RezultStr,vbCrLf)
			If Len(nnnn(0))>0 Then
				RezultStr = nnnn(0)
			End IF
		End IF
		patern = 	"[\s|^]+(�������|��������)+[\s]*[.]+[\s]*[a-zA-Z�-��-�0-9_.]+"
		RezultStr = FindInStrEx(patern, RezultStr)
		if Len(RezultStr)>0 Then
			VidDoca = split(UCase(RezultStr), ".")
			if UBound(VidDoca)>0 Then
				if (VidDoca(0) = UCase("��������")) Then
					Set MetaDoc = MetaData.TaskDef.Childs(CStr("��������"))(CStr(VidDoca(1)))
					Set Table = MetaDoc.Childs("����������������������")
		 			for tbl_cnt = 0 To Table.Count - 1
		 				set tbl_chld = Table(tbl_cnt)
		 				if Len(Rezult)>0 Then
		 					Rezult = Rezult &  vbCrLf & tbl_chld.Name
		 				else
		 					Rezult = tbl_chld.Name
		 				End if
		 			Next
			 	ElseIF (VidDoca(0) = UCase("�������")) Then
			 		Rezult = GetStringRekvizitovFromObj("�������", CStr(VidDoca(1)),0,0,"",0)
				End if
			End if
		End If
	End if


	'������������ �� ������������
	if STipom = 0 Then
		if 1=2 Then
			Patern = "(" & ResultWord & ")+[\s]*[.]+[\s]*(��������|GroupBy)+[\s]*[\(]+[\s]*[" & cnstRExWORD & "\""\s,]+[\s]*[\)]+"
			ttt = FindInStrEx(Patern,TextAll)
			IF Len(ttt)>0 Then
				Patern = "[\(]+[\s]*[" & cnstRExWORD & "\""\s,]+[\s]*[\)]+"
				ttt = FindInStrEx(Patern,ttt)
				IF Len(ttt)>0 Then
					Rezult = ReplaceEx(ttt,Array("(","", ")","", """","", vbCrLf,",", ",,",","))
				End if
			End if
		End if
		ArrRezult = Split(Rezult,vbCrLf)
		tmpRezult = ""
		if UBound(ArrRezult)<> -1 Then
			For ee = 0 To UBound(ArrRezult)
				patern = "(\s|^|,)+(" & ArrRezult(ee) & ")+(\s|^|,|$)+"
				if Len(FindInStrEx(patern,tmpRezult)) = 0 Then
					 If Len(tmpRezult)>0 Then
						tmpRezult = tmpRezult & "," & ArrRezult(ee)
						else
						tmpRezult = ArrRezult(ee)
					End IF
				End IF
			Next
		End IF
		Rezult = Replace(tmpRezult,",",vbCrLf)
	End IF

	GetColumnsFromTZA	= Rezult
End Function

Function GetConstantExA()
	Set Childs = MetaData.TaskDef.Childs(CStr("���������"))
	GetConstantExA = ""
	StrKO = ""
	For i = 0 To Childs.Count - 1
		Set mdo = Childs(i)
		if Len(StrKO) = 0 Then
			StrKO = mdo.Name
		else
			StrKO = StrKO & vbCrLf & mdo.Name
		End if
	next
	IF Len(StrKO) = 0 Then
		Exit Function
	End If
	GetConstantExA = SelectFrom(StrKO, Caption)

End Function

Function GetConstantEx()
	Dim TypeConstArr
	GetConstantEx = ""
	TypeConstStr = ""
	tree = ""
	Set Childs = MetaData.TaskDef.Childs(CStr("���������"))
	For i = 0 To Childs.Count - 1
		Set mdo = Childs(i)
		if Len(TypeConstStr) = 0 Then
			TypeConstStr = mdo.Type.FullName & "##"
		elseIf InStr(1,TypeConstStr,mdo.Type.FullName&"##")=0 Then
			TypeConstStr = TypeConstStr & mdo.Type.FullName & "##"
		End if
	next
	TypeConstArr = split(TypeConstStr,"##")
	if UBound(TypeConstArr)<>-1 Then
		For i = 0 To UBound(TypeConstArr)-1
			tree = tree & TypeConstArr(i)& vbCrLf
			For tt = 0 To Childs.Count - 1
				Set mdo = Childs(tt)
				if (mdo.Type.FullName = TypeConstArr(i)) Then
					tree = tree & vbTab & mdo.Name& vbCrLf
				End if
			next
		next
	End if
	Set srv=CreateObject("Svcsvc.Service")
	Cmd = srv.SelectInTree(tree,"�������....",false)
	ln = Len(Cmd)
	If Ln = 0 Then
		Exit Function
	else
		rrrr = split(Cmd,"\")
		if UBound(rrrr)<> -1 Then
			GetConstantEx = rrrr(UBound(rrrr))
		End if
	End if
End Function

Sub ViewConstantEx()
	AAA = GetConstantEx()
End Sub

Private Function SelectFrom(VLStr, Caption)
	textTo = ""
	if KakVibiraem = 1 Then
		Set srv=CreateObject("Svcsvc.Service")
		textTo = srv.SelectValue(VLStr,Caption,False)
	Elseif KakVibiraem = 2 Then
		Set srv=CreateObject("Svcsvc.Service")
		textTo = srv.FilterValue(VLStr,1 + 4 + 16,Caption,0,0,1)
	Elseif KakVibiraem = 3 Then
		Set SelObj = CreateObject("SelectValue.SelectVal")
		textTo = SelObj.SelectPopUp(VLStr, Windows.ActiveWnd.HWnd, vbCrLf)
	End If
	SelectFrom = textTo
End Function

Private Function GetListOsnovaniyDoca(NameDoc)
	'18 ��������(��������� � 1) - � ��������� "�������� ���������� ���"
	'�������������, ���� ����� ����� ��������� �� ���������
	'������� �������� ��� - ����� ��������� ��� ���������
	'� ���������� � ��� � ���� ��������
	Rezult = ""
	Set Docs = Metadata.TaskDef.Childs("��������")
	For cntdoc = 0 To Docs.Count - 1
		Set Doc = Docs(cntdoc)
		Set PropDoc = Doc.Props
		patrn = "(��������."&NameDoc&")[,]*"
		text = FindInStrEx (patrn, PropDoc(18))
		if Len(text)>0 Then
			Rezult = Rezult&"��������."&PropDoc(0)&" "
		End if
	Next
	GetListOsnovaniyDoca = "#"&Rezult
End Function
'GetMethodsAndRekvOfZapros 'GetVariableAndFunctionZapros

Private Function GetVariableAndFunctionZapros(NameVariable, STipom)
	GetVariableAndFunctionZapros = ""
	Rezult = ""
	textAllmodule = GetLocationText()
	TextCurentProcFunc = GetTextCurentProcFunc()

	For qq = 0 To 1 '��� �������
		'������� ������� �� ������� ���������-�������, ���� �� ���������, ���� ���� �����...
		'������� ��� �� ���������� �������� �� ����� �������
		IF qq=0 Then
			textmodule = TextCurentProcFunc
		Else
			textmodule = textAllmodule
		End IF
		varTextOfZapros = ""

		strPatrn = "(\s|^)*(" & NameVariable & ")+(\s)*(.���������|.Execute)+(\s)*(\()+(\s)*[a-zA-z�-��-�0-9_]+(\s)*\)"
		text = FindInStrEx (strPatrn, textmodule)
		if Len(text)>0 Then
			'���� ����� ������ "�����" ������� ��������� (�������������������������)
			strPatrn = "(\()(\s)*[a-zA-z�-��-�0-9_]+(\s)*\)"
			text = FindInStrEx (strPatrn, text)
			if Len(text)>0 Then
				'������ ������
				strPatrn = "[a-zA-z�-��-�0-9_]+"
				text = FindInStrEx (strPatrn, text)
				if Len(text)>0 Then
					'���� ��� ������ ��� ����������
					varTextOfZapros = text
					Exit For
				End if
			End if
		End if
	Next
	if Len(varTextOfZapros) = 0 Then
		Exit Function
	End if
	varTextZaprosa = ""
	'������� �������� ����� �������
	strPatrn =	"(\s|^)*" & varTextOfZapros & "(\s)*(=)(\s)*("")(\s)*[^#]+(\s)*("")[a-zA-z�-��-�0-9_|,=;\s^+\(\)/{}]*(\s)*(;)"
	text = FindInStrEx (strPatrn, textmodule)
	if Len(text)>0 Then
		'������ �� ������ ������� ���������� ������ ������ ������� ���������� "������������"
		strPatrn = "(\s|^)*" & varTextOfZapros & "(\s)*(=)(\s)*"
		textVZ = FindInStrEx (strPatrn, text)
		if Len(textVZ)>0 Then
			text = Replace(text,textVZ,"")
			'����� � ��� ������ ����� �������, �� ���� ��� � ���������� � �������������� ����������
			'� ������ �������
			varTextZaprosa = text
		End if
	End if

	'���� �������������� ���������� � ������ �������
	strPatrn = "(\s|^)*" & varTextOfZapros & "(\s)*(=)(\s)*"& varTextOfZapros & "(\s)*(\+)(\s)*("")(\s)*[^#]+(\s)*(;)+(\s)*("")+(\s)*(;)+"
	text = FindInStrEx (strPatrn, textmodule)
	if Len(text)>0 Then
		strPatrn = "(\s|^)*" & varTextOfZapros & "(\s)*(=)(\s)*"& varTextOfZapros & "(\s)*(\+)"
		pptext = FindInStrEx (strPatrn, textmodule)
		if Len(pptext)>0 Then
			Text = replace(text,pptext,"")
			varTextZaprosa = varTextZaprosa & text
		End if
	End if
	varTextZaprosa = Replace(varTextZaprosa,vbTab,"")
	varTextZaprosa = Replace(varTextZaprosa,"|","")

	'��������� ���������� �������
	strPatrn = "[a-zA-z�-��-�0-9_|]+(\s)*(=)(\s)*[a-zA-z�-��-�0-9_,\s^\.]+(;)"
	text = FindInStrEx (strPatrn, varTextZaprosa)
	if Len(text)>0 Then
		'�������� �� ����� "�����" :)
		if STipom = 0 Then
			strPatrn = "[a-zA-z�-��-�0-9_|]+(\s)*(=)"
			text = FindInStrEx (strPatrn, text)
			if Len(text)>0 Then
				'������� ��� "�� ����� �����" :)
				strPatrn = "[a-zA-z�-��-�0-9_|]+"
				text = FindInStrEx (strPatrn, text)
				if Len(text)>0 Then
					'������� ���������� �������
					Rezult = text
				End if
			End if
		else
			'����� ������ ��������� ���������� �� �������
			strPatrn = "[a-zA-z�-��-�0-9_|]+(\s)*(=)(\s)*[a-zA-z�-��-�0-9_.]+(\s)*[,|;]"
			text = FindInStrEx (strPatrn, text)
			if Len(text)>0 Then
				Rezult = text
			End if
		End if
	End if
	if STipom = 0 Then '������� ��� ��� ������������ �� �����
		'��������� ������� �������
		strPatrn = "(�������|Function)(\s)+[a-zA-z�-��-�0-9_]+(\s)*(=)"
		text = FindInStrEx (strPatrn, varTextZaprosa)
		if Len(text)>0 Then
			'������� ����� �������
			strPatrn = "(\s)+[a-zA-z�-��-�0-9_]+(\s)*(=)"
			text = FindInStrEx (strPatrn, text)
			if Len(text)>0 Then
				'������� ������� � �����
				strPatrn = "[a-zA-z�-��-�0-9_]+"
				text = FindInStrEx (strPatrn, text)
				if Len(text)>0 Then
					'�������� � ����������� :)
					Rezult = Rezult & vbCrLf & text
				End if
			End if
		End if
		'���������� (������� �� ���������� �������, ����� ����� ����� � ������������)
		ArrRezult = split(Rezult,vbCrLf)
		Rezult = " "
		If UBound(ArrRezult)<> -1 Then
			For ii = 0 To UBound(ArrRezult)
				ttt = Trim(ArrRezult(ii))
				patern = "(\s)+(" & ttt & ")(\s)+"
				text = FindInStrEx (patern,Rezult)
				if Len(text) = 0 Then
					'�������� � ����������� :)
					Rezult = Rezult & " " & ttt & " "
				End if
			Next
		End If
		Rezult = Trim(Rezult)
		Rezult = Replace(Rezult, " ", vbCrLf)
	End if

	GetVariableAndFunctionZapros = Rezult
End Function


Private Function GetLocationText()
	'�������� ����� �� ������ ��� ��������� �� ������ ���������-�������
	GetLocationText = ""
	Rezult = ""
	doc = ""
	if Not CheckWindow(doc) then Exit Function
	if doc.Name <> "���������� ������" Then
		Rezult = Doc.Text
	Else
		virajenieRgExp = "(\s|^|$)+(���������|�������|Procedure|Function)(\s)+[a-zA-z�-��-�0-9_]+\("
		for lncnt = doc.SelStartLine To 0 Step -1
			if Len(FindInStrEx (virajenieRgExp, doc.Range(lncnt)))>0 Then
				Exit For
			End if
			Rezult = doc.Range(lncnt) & vbCrLf & Rezult
		Next
	End if
	GetLocationText = Rezult
End Function

Private Function GetTextCurentProcFunc()
	'�������� ����� �� ������ ��� ��������� �� ������ ���������-�������
	GetTextCurentProcFunc = ""
	Rezult = ""
	doc = ""
	if Not CheckWindow(doc) then Exit Function
	if doc.Name <> "���������� ������" Then
		'��������� �������� ����� ������� ���������, ������ � �� glRWord.doc_SelStartLine,
		'��� ����� ��������� ������ ��������� � ������� �������� ������������..
		If Not IsEmpty(LocalModule) Then
			IF LocalModule.CounItem<>-1 Then
				For qq = 0 To UBound(LocalModule.ArrItem)
					If Not IsEmpty(LocalModule.ArrItem(qq)) Then
						IF (LocalModule.ArrItem(qq).LineStart<glRWord.doc_SelStartLine) And (LocalModule.ArrItem(qq).LineEnd>glRWord.doc_SelStartLine) Then
							Rezult	= LocalModule.ArrItem(qq).text
							Rezult = DellKommentForText(Rezult)
							Exit For
						End IF
					End IF
				Next
			End IF
		End IF
	End if
	GetTextCurentProcFunc = Rezult
End Function

'=============================================================================================
'doc - ������ �����
Function CheckWindowOnWorkbook(doc)
	CheckWindowOnWorkbook = False
	If Windows.ActiveWnd Is Nothing Then
    	Exit Function
	End If
	Set doc = Windows.ActiveWnd.Document

    if CommonScripts.CheckDocOnExtension(doc, NotIntellisenceExtensions) Then Exit Function

	If doc = docWorkBook Then
		if (Doc.ActivePage <> 1) Then
			Exit Function
		End If
		Set doc=doc.Page(1)
		CheckWindowOnWorkbook = True
	Else
		Exit Function
	End If
	If doc.LineCount = 0 Then
    	Exit Function
	End If
	CheckWindowOnWorkbook = True
End Function

'=============================================================================================
Function CheckWindow(doc)
	CheckWindow = False

	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit Function

    if CommonScripts.CheckDocOnExtension(doc, NotIntellisenceExtensions) Then Exit Function

	If doc.LineCount = 0 Then
    	Exit Function
	End If
	CheckWindow = True
End Function

' ����� ������ ���������
' ���� ��������� ������ �� �������� ���������� � ����������� �� �������������()
Private Sub ChoiceDocMet()
	doc = ""
	if Not CheckWindowOnWorkbook(doc) then Exit Sub

	Vals_Method_Doc=GetDocMet(0)
	Vals_Method_Doc = SortStringForList(Vals_Method_Doc, vbCrLf)
	textTo = SelectFrom(Vals_Method_Doc, "��������, ����")
	if (Len(textTo) = 0) Then
		Exit Sub
	End if
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelStartLine,doc.SelStartCol) = textTo
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+Len(textTo)
End Sub

' ����� ������ �����������.
' ���� ��������� ������ �� �������� ���������� � ����������� �� �������������()
Private Sub ChoiceSprMet()
	doc = ""
	if Not CheckWindowOnWorkbook(doc) then Exit Sub

	Vals_Method_Spr=GetSprMet(1)
	Vals_Method_Spr = SortStringForList(Vals_Method_Spr, vbCrLf)
	textTo = SelectFrom(Vals_Method_Spr, "��������, ����")
	if (Len(textTo) = 0) Then
		Exit Sub
	End if
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelStartLine,doc.SelStartCol) = textTo
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+Len(textTo)
End Sub

Private Function UC(Word)
	UC = UCase(Word)
End Function

'���� � ������� ����������� � ������......
Private Function DellKommentForText(ttextBlock)
	DellKommentForText = ttextBlock
	IF len(ttextBlock) = 0 Then	Exit Function
	patern = "(//)+[\S \t\-]*"
	ttext0 = FindInStrEx(patern,ttextBlock)
	If Len(ttext0)>0 Then
		ArrTText0 = Split(ttext0, vbCrLf)
		For u = 0 To Ubound(ArrTText0)
			'Message ArrTText0(u)
			ttextBlock = Replace(ttextBlock,ArrTText0(u) & vbCrLf,"")
		Next
	End IF
	While InStr(1,ttextBlock, vbCrLf&vbCrLf)>0
		ttextBlock = Replace(ttextBlock,vbCrLf&vbCrLf,vbCrLf)
	Wend
	DellKommentForText = ttextBlock
End Function

'������ - ��������� �������� ������ ��������/������� ������, + ��������� ����������������� �� ���
' ��������� ������������ ������ ���������/�������, ���������� � "�������"  ����������
' � ����-�� ���������
' �������� - �������� ������ ��������/������� ������ � ������������ - ������ ������/�����
' � ���������� �������� �� ����������� ��������� = ������������������������������(��������)
' ������� ������ �� ����������� ������

Class TheModuleItem
	' �������� ���������/������� ������ � � ���������....
	Public TypeItem			' ��� �����:	0 - �������� ���������� � ������������ �������� � �������..
							'				1 - ��������
							'				2 - �������
							'				-1 - �� ���������
	Public Name				' ��� ��������� / �������.
	Public NameFull			' ��� ������ ��������� / �������.
	Public LineStart		' ������, ��� ���������� �������/���������
	Public LineEnd			' ������, ��� �������������

	Public LenthText		' ������ ������

	Public Parameters		' ������ ���������� � ������� ���������� �/� �������

	Public Text				' ������ ����� �������
	Public TextBezKommet	' ������ ����� �������, �� ��� ������������..

	Public RetValueStr		' ��� ���������� ������������ ��������...
	Private Sub Class_Initialize
		TypeItem	= -1
		Name		= ""
		LineStart	= 0
		LineEnd		= 0
		Parameters	= ""
		Text		= ""
		TextBezKommet=""
		RetValueStr = ""
	End Sub
	Function GetParamNumber(NumberP)
		tNumberP = Int(NumberP)
		tNumberP = tNumberP - 1
		GetParamNumber = ""
		if Len(Parameters)>0 And tNumberP>0 Then
			ArrParameters = Split(Parameters,",")
			If UBound(ArrParameters)<>-1 Then
				If UBound(ArrParameters)<=tNumberP Then
					GetParamNumber = ArrParameters(tNumberP)
				End If
			End If
		End If
	End Function
End Class

Class TheModule
	Public ModuleName	' ��� ������
	Public CounItem		'���������� ��������/�������
	Public CounLine		'���������� ����� � ������ ����� �� ������������...
	Public ArrItem		'������ ��������/������� ������ ����� "TheModuleItem"
	Public ModuleDoc	'�������� ������ (���� "Text")
	Public NamesAllProcFunk	' ��� ��������� ������� ������ � ������� /�������/�������2/.../� �.�.
	Public ModuleTextArray	'�������� ������ (���� "Text")

	Private Sub Class_Initialize
		ModuleName		= ""	' ��� ������
		CounItem		= "" '���������� ��������/�������
		ArrItem			= Array("")
		NamesAllProcFunk = "/"
	End Sub

	Function IsProcedure(NameProc, LineS, LineE, TypePF, TextProcFunk,TheProc)
		IsProcedure = False
		if IsArray(ArrItem) Then
			If Len(NameProc)>0 Then
				If InStr(1,LCase(NamesAllProcFunk),LCase("/" & TRim(NameProc) & "/"))>0 Then
					For qqq = 0 To UBound(ArrItem)
						IF LCase(ArrItem(qqq).Name) = LCase(TRim(NameProc)) Then
							'�������� �� ������ �� � ��� ��������� ����� � �����...
							If (InStr(1, LCase(ModuleDoc.Range(ArrItem(qqq).LineStart-1)), LCase(TRim(NameProc)))>0) OR (InStr(1, LCase(ModuleDoc.Range(ArrItem(qqq).LineStart)), LCase(TRim(NameProc)))>0) Then
								IsProcedure = True
								LineS		= ArrItem(qqq).LineStart
								LineE		= ArrItem(qqq).LineEnd
								TextProcFunk = ArrItem(qqq).Text
								Set TheProc = ArrItem(qqq)
							Else
								Message "������ ����������������, ����� �����������������!"
							End IF
						End IF
					Next
				End IF
			End IF
		else
			Message "������ �� ���������������!"
		End IF
		'ArrItem			= Array("")
	End Function

	Sub SetDoc(Doc)
		Set ModuleDoc = Doc
		ModuleName = Doc.Name
		ModuleTextArray = Split(Doc.Text,  vbCrLf)
	End Sub

	Sub InitializeModule(ExtractAll)
		GetAllProcFunc(ModuleDoc.Text)
		if (ExtractAll=1) Then
			ExtractNameAndOther()
		End IF
	End Sub

	Sub ReDimModule (cnt)
		ReDim ArrItem(cnt)
	End Sub

	Sub Listing ()
		For ee = 0 To UBound(ArrItem)
			If Not IsEmpty(ArrItem(ee)) Then
				Message ArrItem(ee).LineStart & " " & ArrItem(ee).LineEnd & " " & ArrItem(ee).Parameters & " "
			End If
		Next
	End Sub

	Sub VerifyPosition(FuncProcName, Position)
		MinPos = 0
		MaxPos = ModuleDoc.LineCount

		SkipPosition = Position
		'Position = -1 '�������� ���������� ������
		For qq = 0 To 50
			If qq = 0 Then
			ElseIf qq/2 = qq\2 Then '������ ���� ��������� ������
				SkipPosition = Position - qq
			Else
				SkipPosition = Position + qq
			End IF
			If (SkipPosition>=MinPos) And (SkipPosition<=MaxPos) Then
ProfilerEnter("ModuleDoc.Range")
' ����� ��������� artbear !!
'				ttext = ModuleDoc.Range(SkipPosition)
				ttext = ModuleTextArray(SkipPosition)
ProfilerLeave("ModuleDoc.Range")
				If InStr(1,LCase(ttext),LCase(FuncProcName))=1 Then
					Position = SkipPosition
					Exit For
				End IF
			End IF
		Next
	End Sub
	'==========================================
	'������� ��������� �������� �������� � �������... �������� ��� ����������� ������..
	Sub VerifyAndRecalcPosition()
		MinPos = 0
		MaxPos = ModuleDoc.LineCount
		AddDeleteLines = ModuleDoc.LineCount - MaxPos
		If CounLine <> ModuleDoc.LineCount And False Then ' ���� �������..
			'������ ������... ��� "���������� �������"
			For ee = 0 To UBound(ArrItem)

				If Not IsEmpty(ArrItem(ee)) Then
					'ArrItem(ee).NameFull
				End IF
			Next
		End IF
		CounLine = ModuleDoc.LineCount
	End Sub


	Sub ExtractNameAndOther()
	'Stop
		For ee = 0 To UBound(ArrItem)
			If Not IsEmpty(ArrItem(ee)) Then
				' �������� ��� � ��� ��������� ��� �������
				ArrItem(ee).NameFull = ArrItem(ee).Name
				tName = FindInStrEx("(�������|Function|���������|Procedure)+[\s]+[" & cnstRExWORD &"]+[\s]*[\(]+",ArrItem(ee).Name)
				tType = FindInStrEx("(�������|Function|���������|Procedure)+[\s]+",tName)
				tName = ReplaceEx(tName, Array(tType,"", "(",""))
				tName = Trim(tName)
				tType = UCAse(Trim(tType))
				ArrItem(ee).Name = tName
				NamesAllProcFunk = NamesAllProcFunk & lCase(tName) & "/"
				Select case tType
				case "�������":					ArrItem(ee).TypeItem = 2
				case "FUNCTION":    			ArrItem(ee).TypeItem = 2
				case "���������"				ArrItem(ee).TypeItem = 1
				case "PROCEDURE"				ArrItem(ee).TypeItem = 1
				end select
				' �������� ���������...
				tName = FindInStrEx("(�������|Function|���������|Procedure)+[\s]+[" & cnstRExWORD &"]+[\s]*[\(]+",ArrItem(ee).NameFull)
				'������ ����
				tName =  "(" &Replace(ArrItem(ee).NameFull,tName,"")
				tName = ReplaceEx(tName, Array("(","", ")",""))
				' ������� ���������� ������� = "���-��"
				tStr = FindInStrEx("\=+\s*("")+\s*[" & cnstRExWORD &",\.\\\/\s]*\s*("")+",tName)
				If Len(tStr)>0 Then
					tArr = Split(tStr,vbCrLf)
					tName = ClearInSring(tName, tArr)
				End If
				tStr = FindInStrEx("\=+\s*[""""]+",tName)
				If Len(tStr)>0 Then
					tArr = Split(tStr,vbCrLf)
					tName = ClearInSring(tName, tArr)
				End If
				tStr = FindInStrEx("\=+\s*[0-9\.]+",tName)
				If Len(tStr)>0 Then
					tArr = Split(tStr,vbCrLf)
					tName = ClearInSring(tName, tArr)
				End If
				tName = LCase(tName)
				tArr = Split(tName,",")
				If UBound(tArr)<>-1 Then
					For qq = 0 To UBound(tArr)
						tArr(qq) = Trim(tArr(qq))
						If Len(tArr(qq))>0 Then
							tArr(qq) = Replace(tArr(qq), "���� ","")
						End If
					Next
					ArrItem(ee).Parameters	= Join(tArr,",")
				End If
			End If
		Next
	End Sub

Dim Profiler
Private Sub ProfilerEnter(name)
'	Set Profiler = CreateObject("SHPCE.Profiler")
'	Profiler.StartPiece(name)
End Sub

Private Sub ProfilerLeave(name)
'	Set Profiler = CreateObject("SHPCE.Profiler")
'	Profiler.EndPiece(name)
End Sub

Private Sub GetAllProcFunc(textModule) ',ArrNameProcFunc, ArrTextProcFunc)
'Set Profiler = CreateObject("SHPCE.Profiler")

ProfilerEnter("GetAllProcFunc")
ProfilerEnter("0")

		CounLine = ModuleDoc.LineCount
		Patern = "(�������|Function|���������|Procedure)+[\s]+[" & cnstRExWORD &"]+[\s]*[\(]+[" & cnstRExWORD &"=, \t""]*[\)]*" '\s*(�������|Expotr)*\s*(�����|Forward)*" '-��������� �������
''		ttextPF = FindInStrEx(Patern,textModule)
''
''		ArrNameProcFunc  = Array("")
''stop
'		ttextPF = FindInStrEx(Patern,textModule)
'			ArrNameProcFunc = Split(ttextPF, vbCrLf)
'Message UBound(ArrNameProcFunc)
'if UBound(ArrNameProcFunc)-2 > -1 then Message ArrNameProcFunc(UBound(ArrNameProcFunc)-2)
'if UBound(ArrNameProcFunc)-1 > -1 then Message ArrNameProcFunc(UBound(ArrNameProcFunc)-1)

		ArrNameProcFunc = FindInStrExArtur(Patern,textModule)
		textModuleSrc = textModule
'Message UBound(ArrNameProcFunc)
'if UBound(ArrNameProcFunc)-2 > -1 then Message ArrNameProcFunc(UBound(ArrNameProcFunc)-2)
'if UBound(ArrNameProcFunc)-1 > -1 then Message ArrNameProcFunc(UBound(ArrNameProcFunc)-1)

		ArrForLocText = ""
		PosithionInModule = 0
		ttextPFCurent = ""	'ttextPF = Replace(ttextPF,"(","")
'		If Len(ttextPF)>0 Then
		If UBound(ArrNameProcFunc)>0 Then
'			ArrNameProcFunc = Split(ttextPF, vbCrLf)
			If UBound(ArrNameProcFunc)>0 Then
				ReDimModule UBound(ArrNameProcFunc)
				Dim ItemsCount
				ItemsCount = UBound(ArrNameProcFunc)
				'For ee = 0 To UBound(ArrNameProcFunc)-1 : Status "����� ����/����.... " & GetProcent(ee, UBound(ArrNameProcFunc))
ProfilerLeave("0")
ProfilerEnter("For")
				For ee = 0 To ItemsCount-1
ProfilerEnter("Status")
					Status "1 ����� ����/����.... " & GetProcent(ee, ItemsCount)
ProfilerLeave("Status")

					'���� ����� ����� ���� N �� ������� � ���� N+1 � ��������� ����� ����� ����.
					'���� ������������ ����� ��������� ���������

					' �������� ����...
					'Message ArrNameProcFunc(ee)
ProfilerEnter("1")
					Set ItemModule = New TheModuleItem
'					Pos1 = InStr(1, textModule,ArrNameProcFunc(ee))
					Pos1 = FindInStrExArturPositionArray(ee)
					if ee = 0 Then
						'������ ��������� � ��������� - �������� �������� ������
'						ttextPFCurent = Mid(textModule,1,pos1)
						ttextPFCurent = Mid(textModuleSrc,1,pos1)

						ArrForLocText = Split(ttextPFCurent,vbCrLf)
						PosithionInModule = UBound(ArrForLocText)
						if (PosithionInModule = -1) Then
							PosithionInModule = 1
						Else
							PosithionInModule = PosithionInModule + 1
						End IF
					End IF
ProfilerLeave("1")
ProfilerEnter("2")
ProfilerEnter("2.0")
'					VerifyPosition ArrNameProcFunc(ee), PosithionInModule
ProfilerLeave("2.0")
ProfilerEnter("2.1")
'					Pos2 = InStr(1, textModule,ArrNameProcFunc(ee+1))
					Pos2 = CLng(FindInStrExArturPositionArray(ee+1))
					if Pos2 = 0 then
						Pos2 = Len(textModuleSrc)
					end if
'					ttextPFCurent = Mid(textModule,pos1,Pos2-pos1)
					ttextPFCurent = Mid(textModuleSrc,pos1,Pos2-pos1)

					ItemModule.LenthText = Pos2-pos1
					ItemModule.Text = ttextPFCurent
ProfilerLeave("2.1")
ProfilerEnter("2.2")
					ArrForLocText = Split(ttextPFCurent,vbCrLf)
ProfilerLeave("2.2")
ProfilerEnter("2.3")
					ItemModule.LineStart = PosithionInModule
					PosithionInModule = PosithionInModule + UBound(ArrForLocText)
					ItemModule.LineEnd	= PosithionInModule-1
					ItemModule.Name = ArrNameProcFunc(ee)
ProfilerLeave("2.3")
ProfilerEnter("2.4")
					Set ArrItem(ee) = ItemModule
ProfilerLeave("2.4")
ProfilerEnter("2.5")

					textModule = Mid(textModule,Pos2)
ProfilerLeave("2.5")
ProfilerLeave("2")
				Next
ProfilerLeave("For")
			End IF
		End if
ProfilerEnter("3")
		Set ItemModule = New TheModuleItem

		ItemModule.Text = textModule
		ArrForLocText = Split(textModule,vbCrLf)
		ItemModule.LineStart = PosithionInModule
		PosithionInModule = PosithionInModule + UBound(ArrForLocText)
		ItemModule.LineEnd	= PosithionInModule
		ItemModule.Name = ArrNameProcFunc( UBound(ArrNameProcFunc))
		Set ArrItem(ee) = ItemModule
		Status ""
ProfilerLeave("3")
ProfilerLeave("GetAllProcFunc")
'Set Results = Profiler.Results
'For Each R In Results
''	Message R & "/" & TypeName(R) & "/" & IsObject(R) & "/" & VarType(R)
'	Message R
'Next
	End Sub
End Class

Sub NewInicializeModule()
'������������ ������������� ������..
	Set LocalModule = New TheModule
	doc = ""
	if Not CheckWindow(doc) then Exit Sub

	LocalModule.SetDoc(doc)
	LocalModule.InitializeModule(1)
	'LocalModule.ExtractNameAndOther
	LocalModule.Listing
End Sub




Sub InicializeModule()
'stop
	Set LocalModule = New TheModule
	TextModuleInTxtFilesProc()
	LocalModule.ExtractNameAndOther
	'LocalModule.Listing
End Sub


Private Sub TextModuleInTxtFilesProc()
	doc = ""
	if Not CheckWindow(doc) then Exit Sub
	'�������� ��������� ���� �������� � ������� ������
	ttext = doc.text
	ArrTtextOfPF  = Array("")
	ArrPF		  =  Array("")
	GetAllProcFunc ttext, ArrPF, ArrTtextOfPF
	For ee = 0 To Ubound(ArrTtextOfPF) ': Status "������ �����������..."
		'ttextProc = DellKommentForText(ArrTtextOfPF(ee))
		ttextProc = ArrTtextOfPF(ee)
	Next :	Status "������ �����������...��!"
End Sub

Private Function GetProcent(p1, p2)
	GetProcent = " %"
	proc = p1 * 100 / p2
	proc = "" & proc
	Arr = Split(proc,",")
	GetProcent = "" & Arr(0) & GetProcent
End Function

Private Function ClearInSring(tStr, tArr)
	ClearInSring = tStr
	if IsArray(tArr) Then
		If UBound(tArr)<>-1 Then
			For qq = 0 To UBound(tArr)
				If Len(tArr(qq))>0 Then
					tStr = Replace(tStr, tArr(qq), "")
				End If
			Next
		End If
	Else
		tStr = Replace(tStr, tArr, "")
	End If
	ClearInSring = tStr
End Function

Private Function GetStrAccessRight()
	ArrstrRekv = Array("������", "�����������������������", "�������������������������", "�����������������������������", _
	 "�������������������������������", "�������", "OLEAutomationServer", "��������������������������", "��������������������", "��������������������������������", _
	 "��������������������������������", "��������������������������", "�����������������������������", "������������������������", _
	 "����������������", "��������������������������", _
	 "�������������", "��������������",	 "����������",	 "��������", "�����������������", _
	 "�����������������������", 	 "�����", 	 "�����������������������������", _
	 "�������������������", 	 "������������������������������", 	 "������������������������������������", _
	 "��������������������������������", "�������������", 	 "��������������", _
	 "�������������������������������",  "�������������������������", 	 "��������������������������", _
	 "��������������������������������",  "��������������������������������������",  "������������������������")
	 GetStrAccessRight = Join(ArrstrRekv,vbCrLf)
	 GetStrAccessRight = SortStringForList(GetStrAccessRight, vbCrLf)

End Function


Private Function FindInStrExA (patrn, strng)
	FindFirstInFindInStrEx = True
	FindInStrExA = FindInStrEx (patrn, strng)
	FindFirstInFindInStrEx = false
End Function

Private Function FindInStrEx (patrn, strng)
  regEx.Pattern = patrn			' Set pattern.
  Set Matches = regEx.Execute(strng)	' Execute search.
  RetStr = ""
  For Each Match in Matches		' Iterate Matches collection.
	if Len(RetStr)>0 Then
		RetStr = RetStr & vbCrLf & Match.Value
	else
		RetStr = Match.Value
    End if
    if (FindFirstInFindInStrEx = True) Then
		Exit For
    End if
  Next
  FindInStrEx = RetStr
End Function

' artbear
Dim FindInStrExDict
'set FindInStrExDict = CreateObject("Scripting.Dictionary")
Dim FindInStrExArturNamesArray, FindInStrExArturPositionArray

Private Function FindInStrExArtur(patrn, strng)
  regEx.Pattern = patrn			' Set pattern.
  Set Matches = regEx.Execute(strng)	' Execute search.
'  RetStr = ""
  Redim FindInStrExArturNamesArray(Matches.Count)
  Redim FindInStrExArturPositionArray(Matches.Count)
  Dim i
  i = 0
  For Each Match in Matches		' Iterate Matches collection.
'	if Len(RetStr)>0 Then
'		RetStr = RetStr & vbCrLf & Match.Value
'	else
'		RetStr = Match.Value
'    End if
		FindInStrExArturNamesArray(i) = Match.Value
		FindInStrExArturPositionArray(i) = Match.FirstIndex + 1
		i = i + 1
    if (FindFirstInFindInStrEx = True) Then
		Exit For
    End if
  Next
'  FindInStrEx = RetStr
  FindInStrExArtur = FindInStrExArturNamesArray
End Function 'FindInStrExArtur


Set regEx	= New RegExp : regEx.IgnoreCase = True : regEx.Global = True
Set GRegExp = New RegExp : GRegExp.IgnoreCase = True : GRegExp.Global = True
FindFirstInFindInStrEx = False

Private Function VibColor()	'������ �����
	VibColor = "-1"
	set CD = CreateObject("MSComDlg.CommonDialog")
	CD.ShowColor()
	VibColor = CD.Color
End Function

Sub AddWordToSlovar()
	Doc = ""
	If CheckWindow(Doc) Then
		ttext = doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine,doc.SelEndCol)
		If (Len(ttext) > 0) Then'And (InStr(1,ttext," ") = 0) Then
			AddToSlovar(ttext)
		End IF
	End IF
End Sub

Sub InsertFromSlovar()
	Doc = ""
	If Not CheckWindow(doc) Then
		Exit sub
	End IF
	ttext  = GetWordFromSlovar()
	If Len(ttext) = 0 Then
		Exit sub
	End IF
	Pos = len(ttext)
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelStartLine,doc.SelStartCol) = ttext
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+Pos
End Sub


Sub AddToSlovar(Word)
	If Len(Word)<1 Then
		Exit Sub
	End IF
	ttext = LoadSlovar()
	If Len(ttext)>0 Then
		If InStr(1,lcase(ttext),lcase(Word))>0 Then
			Exit Sub
		End IF
	End IF

	DictFileName = BinDir + "DictDots.txt"
	FileExists = FSO.FileExists(DictFileName)
	IF Not FileExists Then
		FSO.CreateTextFile(DictFileName)
	End IF
	FileExists = FSO.FileExists(DictFileName)
	IF Not FileExists Then
		Exit Sub
	End IF
	Set Fl = FSO.OpenTextFile(DictFileName,8)
	Fl.WriteLine(Word)
	Fl.Close
End sub

Function LoadSlovar()
	LoadSlovar = ""
	DictFileName = BinDir + "DictDots.txt"
	FileExists = FSO.FileExists(DictFileName)
	IF FileExists Then
		Set Fl = FSO.GetFile(DictFileName)
		Set FileStream = Fl.OpenAsTextStream()
		If Fl.Size<> 0 Then
			LoadSlovar = FileStream.ReadAll()
		End IF
	End IF
End Function

Function GetTypeFromVid(tempNameObj)
	GetTypeFromVid = ""
	ArrTipes = Array("��������","����������","�������","���������", "������������")
	For tttt = 0 To 4
		Set MDObjekts = Metadata.TaskDef.Childs(CStr(ArrTipes(tttt)))
		For cntdoc = 0 To MDObjekts.Count - 1
			Set MDObj = MDObjekts(cntdoc)
			if (UCase(MDObj.Name) = UCase(tempNameObj)) Then
				GetTypeFromVid = ArrTipes(tttt)
				tempNameObj = ""
				Exit Function
			End If
		Next
	Next
End Function

Function GetWordFromSlovar()
	ttext = LoadSlovar()
	ttext = SelectFrom(ttext,"")
	If Len(ttext) <> 0 Then
		GetWordFromSlovar = ttext
	End IF
End Function

Private Sub AddToDict(Dict, ttextAdd, ttextObj, Spliter)
	ttextAdd = lcase(ttextAdd)
	ArrText = Split(ttextAdd,Spliter)
	For qq = 0 To Ubound(ArrText)
		Dict.Add ArrText(qq), ttextObj
	Next
End Sub
Private Function CompareNoCase(str1, str2, IsTrim)
	CompareNoCase = false
	Tstr1 = str1 : 	Tstr2 = str2
	if IsTrim = 1 Then
		Tstr1 = Trim(str1) : 	Tstr2 = Trim(str2)
	End IF
	if LCase(Tstr1) = LCase(Tstr2) Then
		CompareNoCase = true
	End IF

End Function

'���������� ����� ��������� ������ ������� ������ � ������ ������.
Private Function StrCountOccur(str1, str2)
	StrCountOccur = 0
	Tstr1 = lcase(str1) : 	Tstr2 = lcase(str2)
	find = InStr(1,Tstr1,Tstr2)
	Do While find
		StrCountOccur = StrCountOccur + 1
		find = InStr(find + Len(str2),Tstr1,Tstr2)
	Loop
End Function

Sub Configurator_OnActivateWindow(Wnd,bActive) 'As ICfgWindow, ByVal bActive As Boolean)
' artbear - ������� ������, ����������� ��� �������� ���� ���� (����� ����) - ���� ���������� ���� ��� �� ����������
	caption = ""
	on error resume next
	Caption = W.Caption
	on error goto 0
	If Caption = "" Then Exit Sub

	On Error Resume Next
    Set doc134 = Wnd.Document
    iErrNumber = err.Number
	On Error GoTo 0
    If iErrNumber <> 0 Then
        Exit Sub
    End If

	if Not bActive Then
		Exit Sub
	End IF
	on error resume next
	iErrNumber = Err.Number
	on error goto 0
	Set doc = Wnd.Document
	if doc.Type = 0 Then
		Exit Sub
	End IF
	If Doc.Name = "���������� ������" Then
		GlobalModuleParse = 0
	End IF
End Sub

'================================================
Private Function ReplaceEx(ttext, Arr1)
	ReplaceEx = ttext
	For i=0 To UBound(Arr1) / 2
		ttext = Replace(ttext, Arr1(i*2), Arr1(i*2+1))
	Next
	ReplaceEx = ttext
End Function

' ���� �����, ������ ����������� ������� �����:
' = ������������.��������������������.����������������
' ����� �������� � ���������� ���������� ���������� ����
' ������������.�������������������� � �������� �� ��� ����
' ����������. ��������� ����������
Sub ExtraktGlobalVariableName(ttextIns)
	Arrttext = Split(ttextIns,".")
	if UBound(Arrttext) = 2 Then
		ttext = Arrttext(0) & "." & Arrttext(1)
		ttext = lcase(ttext)
		a = GlobalVariableType.Items             ' ���� � ������ ������ ������������
		PosFind =  -1
		For i = 0 To GlobalVariableType.Count -1
			if Trim(lcase(a(i))) = ttext Then
				PosFind = i
				Exit For
			End IF
		Next
		if PosFind<>-1 Then
			s = GlobalVariableNoCase.Items
			ttextIns = s(PosFind) & "." & Arrttext(2)
		End IF
	End IF
End Sub

Function SimpleType(TypeVid1)
	SimpleType = False
	TypeVid = UCase(trim(TypeVid1))
	if Len(TypeVid)>0 Then
		Select case TypeVid
		case "������":	SimpleType = True
		case "STRING":	SimpleType = True
		case "�����":   SimpleType = True
		case "NUMBER":   SimpleType = True
		case "����":	SimpleType = True
		case "DATE":	SimpleType = True
		end select
	End IF

End Function

Function ObjectExist(TType, tVid)
	ObjectExist = false
	if InStr(1,lcase("/���������/��������/����������/�������/�����/���������/"), lcase("/" & TType & "/"))>0 Then
		Set MetaObjcts = MetaData.TaskDef.Childs(CStr(TType))
		For tt = 0 To MetaObjcts.Count-1
			if LCase(MetaObjcts(tt).Name) = LCase(tVid) Then
				ObjectExist = true
				exit for
			End IF
		Next
	End IF
End Function

'====================================
'������ �� ��������� ������� ���������...
Function Get_file_vk_dict(StrObjekta)
	Get_file_vk_dict = ""
	lcStrObjekta = lcase(StrObjekta)
	if vk_dict_CreateObject.Exists(lcStrObjekta) Then
		Get_file_vk_dict = vk_dict_CreateObject.Item(lcStrObjekta)
	End IF
End Function

Sub Loadvk_dict_CreateObject()
	'��������� ������� ������� ���������..
	vk_dict_CreateObject.RemoveAll()
	File_vk_dictName = BinDir & "config\Intell\vk_dict_CreateObject.Dict"
	FileExists = FSO.FileExists(File_vk_dictName)
	if FileExists = True then
		Set Fl = FSO.GetFile(File_vk_dictName)
		Set FileStream = Fl.OpenAsTextStream()
		AllMeth = FileStream.ReadAll()
		AllMethods = Split(AllMeth, vbCrLf)
		for i = 0 to UBound(AllMethods)
			IF Len(AllMethods(i))>0 Then
				Arr2 = Split(AllMethods(i)," ")
				if UBound(Arr2)>0 Then
					Arr2(0) = trim(Arr2(0))
					Arr2(1) = trim(Arr2(1))
					Arr2(1) = Replace(Arr2(1),".ints","")
					if vk_dict_CreateObject.Exists(lcase(Arr2(0))) Then
						message "������������ �������� � �����" & File_vk_dictName
					Else
						vk_dict_CreateObject.Add lcase(Arr2(0)), Arr2(1)
					end if
				end if
			end if
		next
	end if
End Sub

Sub Init(dummy) ' ��������� ��������, ����� ��������� �� �������� � �������
	InitScript(0) ' ������������� ������ ����������

	'*************************��������**********************
	ttextObjekts = "��������"
	ttext = "�������,DocDate,��������,DocNum,���������������,CurrentDocument,����������������,SelectDocuments," & _
		"���������������������������,SelectChildDocs,���������������������������,SelectBySequence,���������������,SelectByNum," & _
		"�����������������,SortLines,��������,MakeActions,��������������������,UnPost,����������,CompareWithAP"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","
	'��� ������� ����� ��������� ��������� :)
	ttext = GetORD()
	ttext = Replace(ttext,vbCrLf,",")
	If Len(ttext)>0 Then
		AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","
	End IF

	'*************************��������������**********************
	ttextObjekts = "��������������"
	ttext = "������������,GetListSize,��������������������������,SortByPresent,����������������������,FromSeparatedString,���������������������," & _
	"ToSeparatedString,�������,Check,����������,RemoveAll,����������������,MoveValue"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","

	'*************************����������**********************
	ttextObjekts = "����������"
	ttext = "�������������������,FindByDescr,����������������,FindByAttribute,������������,FindItem,���������������,GetItem," & _
	"���������������������,UseOwner,�������������������,IncludeChildren,������������,OrderByCode," & _
	"�������������������,OrderByDescr,����������������,OrderByAttribute,�����������,NewGroup," & _
	"�����������,CodePrefix,������������������,SetNewCode"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","

	'*************************�������**********************
	ttextObjekts = "�������"
	ttext = "�������������,PutSection,������������������,AttachSection,�������������,NewPage,�������������,TableHeight," & _
	"������������,SectionWidth,��������������,ReadOnly,������,Protection,�������,Area," & _
	"�������������,PrintRange,�����������������,PageSetup,���������������������,NumberOfCopies,���������������������,CopyesPerPage," & _
	"����������,Print,���������������������,ValueOfCurrentCell"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","


	'*************************���������������**********************
	ttextObjekts = "���������������"
	ttext = "�����������������,ColumnCount,���������������,InsertColumn,��������������DeleteColumn,��������������������������,SetColumnParameters," & _
	"������������������������,GetColumnParameters,��������������,MoveLine,���������,Fill," & _
	"��������,GroupBy,����������������,ColumnVisibility,�����������,Fix,�������������������,ShowImages"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","

	'*************************���������������**********************
	ttextObjekts = "�������"
	ttext = "������,Income,������,Outcome,���������������,SelectActs,������������������������,SelectDocActs,����������������,GetDocAct," & _
	"������������,SelectTotals,������������,GetTotal,���������������,TempCalc,��������������,RetrieveTotals," & _
	"�������������������������,SetFilterValue,������������������,UsePeriod,�����������,ConsolidatedTotal,������������,ConsolidatedTotals," & _
	"�������������,TotalsGet,�������,Rest,��������������,ConsolidatedRest,�������,Rests,��������������,ConsolidatedRests,���������������," & _
	"GetRests,�����������������,LinkLine,��������������,ActIncome,��������������,ActOutcome,�����������������������,DoActIncome," & _
	"�����������������������,DoActOutcome,��������,Act,�����������������,DoAct"
	AddToDict UniMetodsAndRekv, ttext, ttextObjekts, ","
	SoobshitType = False
	Loadvk_dict_CreateObject
End Sub

'�������� ������ �� �������� "�����.<��������>
Sub SyntaxCheckModule()
	Doc = "" : WorkBook = ""
	If Not CheckWindowOnWorkbook(Doc) Then
		Exit Sub
	End IF

	' �������� ���������� ��������� �����.������������
	'������ ������� ������� "�����". �� ���� ��������� �� ��������...
	ArrOFMethods = Array("��������","TabCtrl","��������","Parameter","��������������","ReadOnly","��������","Refresh","��������������������","TabCtrlState",_
	"����������������","UseLayer","���������","Caption","������������������","ToolBar","�����������������","DefButton",_
	"���������������������","ProcessSelectLine","��������������","MakeChoice","�����������","ChoiceMode",_
	"��������������","ModalMode","���������������","GetAttrib","���������������","ActiveControl","��������������","CurrentColumn","�������","Close")

	Patern = "(\s|^|;)+(�����|Form)+\s*\.+[" & cnstRExWORD & "]+"
	'������ ��������� � ����� "�����.��������
	ttext = FindInStrEx(Patern, Doc.text)
	ArrFR = Split(ttext,vbCrLf)
	ttextROF = GetTableRecvFromForms(0,FALSE)
	ArrROF = Split(ttextROF,vbCrLf)
	if (UBound(ArrFR)<>-1) And (UBound(ArrROF)<>-1) Then
		For qq=0 To UBound(ArrFR)
			ArrFR(qq) = ReplaceEX(ArrFR(qq), Array(" ","",";","", vbcr,"",vbTab,""))
			if Len(ArrFR(qq))>0 Then
				Arr = Split(ArrFR(qq),".")
				if UBound(Arr)=1 Then
					IsFind = False
					if Not FindInArray(Arr(1),ArrOFMethods) Then
						For qq2=0 To UBound(ArrROF)
							if UCase(Arr(1)) = UCase(ArrROF(qq2)) Then
								IsFind = True
								Exit For
							End IF
						Next
						If Not IsFind Then
							Message "����������� ���������: """ & ArrFR(qq)&"""", 2
						End IF
					End IF
				End IF
			End IF
		Next
	End If
End Sub

Function FindInArray(TheWord, TheArray)
	FindInArray = False
	For qq = 0 To UBound(TheArray)
		if Trim(UCase(TheArray(qq))) = Trim(UCase(TheWord)) Then
			FindInArray = True
			Exit For
		End IF
	Next
End Function

'
' ��������� ������������� �������
'
Sub InitScript(dummy) ' ��������� ��������, ����� ��������� �� �������� � �������

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

End Sub ' InitScript

Init 0