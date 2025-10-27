$NAME SciColorer

sub ShowSettings() '���������...
	SendCommand 20060
end sub

sub ExpandAll() '���������� ��
	SendCommand 20065
end sub

sub CollapseAll() '�������� ��
	SendCommand 20066
end sub

sub ToggleCurrent() '��������/���������� ������� ���� (Click)
	SendCommand 20067
end sub

sub ToggleCurrentWithSubLevels() '����������/�������� ������� ������ � ����������� (Ctrl+Click)
	SendCommand 20068
end sub

sub SelectCurrentBlock() '�������� ������� ����
	SendCommand 20069
end sub

sub ToggleReadOnlyMode() '������������ ������ "������ ������"
	SendCommand 20070
end sub

sub NextModifiedLine() '������� � ��������� ���������������� ������
	SendCommand 20071
end sub

sub PrevModifiedLine() '������� � ���������� ���������������� ������
	SendCommand 20072
end sub

sub ResetModifiedLines() '����� ������������������ �����
	SendCommand 20073
end sub

sub ShowBookmarksList() '�������� ������ �������� ������
	SendCommand 20074
end sub

sub ShowModifiedLinesList() '�������� ������ ���������������� ����� ������
	SendCommand 20075
end sub

'�������� ���������� ���������. ���������� ������������, ���� �� �����-�� ������� ��������� ���������� ��������� ������
sub RefreshEditor() 
	SendCommand 20076
end sub

sub SetBgColor() '���������� ���� ���� ���������� �����
	SendCommand 20077
end sub

sub ReloadSettings() '���������� ��������� �� ini �����, � ������ ���� ��� ���� �������� ������� �� ����� ������ �������������
	SciColorer.ReloadSettings()
end sub

sub ToggleViewNonPrintable() '��������/��������� ����������� ���������� ��������
	SendCommand 20080
end sub

sub ShowSearchResults() '��������� ����� ���������� � ������� ������ � �������� ���� � ������������ ������
	SendCommand 20081
end sub

Sub TogleSearchResults() '������/�������� ������ ����������� ������ (������ ������, ��� ������ ������ ������)
	Windows.PanelVisible("���������� ������")=Not Windows.PanelVisible("���������� ������")	
End Sub

sub ShowSearchResultsOfSelectedText() '��������� ����� ������, ����������� � ������� ������ (��� ����� � ������� ��������� ������), � �������� ���� � ������������ ������
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit sub
	selText = Doc.Range(Doc.SelStartLine , Doc.SelStartCol , Doc.SelEndLine , Doc.SelEndCol)
	if Trim(selText) = "" then selText = doc.CurrentWord
	if Trim(selText) = "" then exit sub
	
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(Windows.MainWnd.Hwnd,0,"AfxControlBar42",NULL)
	
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llls", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(combo,0,NULL,"�����������")
	
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llss", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(combo,0,"ComboBox","")
	
	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	Wrapper.SendMessage combo,WM_SETFOCUS,0,0
	
	Wrapper.Register "USER32.DLL",   "GetWindow",    "I=ll", "f=s", "r=l"
	edit = Wrapper.GetWindow(combo,GW_CHILD)
	
	Set svc = CreateObject("Svcsvc.Service")
	svc.SetWindowText edit, selText	
	Wrapper.SendMessage edit, EM_SETSEL, 0, len(selText)

	SendCommand 20081
end sub




'======================================================================
sub SciColorer_OnHotSpotClick() '���������� ��� ����� ����� �� "�����������"
	
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit sub
	
	prevLine	= doc.SelStartLine
	prevCol	= doc.SelStartCol
	prevWnd = Windows.ActiveWnd

	'������� ������� � ����������� ���������� ��� ��������� ��� ������ ��������
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	Wrapper.SendMessage Windows.MainWnd.HWND, &H111 ,22503, 0 'WM_COMMAND
	'message "telepat"
	if prevWnd <> Windows.ActiveWnd then '�������� ������������ � ����������
		exit sub 
	end if
	if  (doc.SelStartLine = prevLine) and (doc.SelStartCol = prevCol) then
		'����� ������� ������� � ����������� ���������� ��� ������ �������
		'message "intell"
		set scr = nothing
		on error resume next
		set scr = scripts("���������")
		on error goto 0
		if scr is nothing then
			message "C����� ""���������"" �� ����������"	
		End If
		scr.VarDefJump()
		if  (doc.SelStartLine = prevLine) and (doc.SelStartCol = prevCol) then
			'������ �� ��������, ����� ������ �������� � ��������� ������ ����� �� ������,
			'��� ���� ���������� ���-�� ������������� ��� ����� � ���������� ����� "�����"
			'message "assign"
			curWord = scr.GetObjectName(doc.Range(prevLine),prevCol," .,;:|#=+-*/%?<>\()[]{}!~@$^&'""" & vbTab)
			for line = prevLine-1 to 0 step -1
				str = doc.Range(line,0)
				if CommonScripts.RegExpTest("("+curWord+"\s*=)|(�����\s+"+curWord+")",str) then
					CommonScripts.Jump line
					exit for
				end if
			next
		end if
	end if
end sub


sub SciColorer_OnLineNumbersContextMenu() '���������� ��� ����������� ���� �� ������� � �������� �����
	ShowBookmarksList() '�������� ������ �������� ��������
	'Scripts("Bookmarks").SelectBookMark() '�������� ������ �������� ������� Bookmarks
end sub

'======================================================================
set obj = nothing
on error resume next
set obj = Plugins("SciColorer")
on error goto 0
if obj is nothing then
	MsgBox "SciColorer: ������ �������� ������� SciColorer.dll"	
	Scripts.Unload SelfScript.Name
else
	SelfScript.AddNamedItem "SciColorer", obj, False
End If

set obj = nothing
on error resume next
set obj = CreateObject("OpenConf.CommonServices")
on error goto 0
if obj is nothing then 
	MsgBox "SciColorer: ������ ��� �������� ������� OpenConf.CommonServices, ��������������� ���������� \Config\System\CommonServices.wsc"
	Scripts.UnLoad SelfScript.Name
else
	SelfScript.AddNamedItem "CommonScripts", obj, False
	CommonScripts.SetConfig(Configurator)
end if



