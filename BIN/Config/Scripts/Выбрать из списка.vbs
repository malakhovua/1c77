$NAME Выбрать из списка

'=============================================================================================
'	Скрипт для OpenConf. Выбор значений из выпадающего списка с фильтром при наборе с клавиатуры.
'
'	Версия: $Revision: 1.2 $
'   
'	Автор: Юхлин Антон ICQ: 200709902
'	Рефакторинг, доработка: metaeditor <shotfire@inbox.ru>
'
'=============================================================================================
'ОПИСАНИЕ:
'
'	21.04.2006 - макрос SelectFromComboBox переделан с помощью нового функционала Svcsvc.dll
'
'	Позволяет осуществить быстрый выбор из выпадающего списка, в котором находится фокус ввода.
'	Выбор производится с помощью фильтрующегося списка.
'	Удобно при выборе типа реквизита метаданного в диалоге "Свойства Реквизита" или 
'	реквизита "ПолеВвода" в диалоге "Свойства", когда фокус ввода находится в 
'	выпадающем списке "Тип" (реквизит диалога) или "Тип значения" (реквизит метаданного).
'	Скрипт также работает и в обычных списках, например при поиске в синтаксис-помошнике
'	Если "повесить" макрос на горячую клавишу, то он отрабатывает и в модальных окнах.
'=============================================================================================

'========================================================================================
'версия через Svcsvc
'========================================================================================
Sub SelectFromComboBox()
	Set Svc = CreateObject("Svcsvc.Service")
	strSelItem = Svc.FilterValue("",1 + 128,"",0,0,1)
End Sub

'========================================================================================
'пример использования Svcsvc::GetWindowText
'========================================================================================
Sub TestSvcSvcGetWindowText()
	Set Svc = CreateObject("Svcsvc.Service")
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "GetFocus", "f=s", "r=l"       
	Wrapper.Register "USER32.DLL",   "GetClassName","I=lrl",  "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetParent",   "I=l",    "f=s", "r=l"
	
	strClassName = Space(129)
	Combo = Wrapper.GetFocus
	
	bList = false
	cnt = Wrapper.GetClassName(Combo, strClassName, 128)
	strClassName = lcase(trim(strClassName))
	if strClassName = "edit" then
		ComboPar = Wrapper.GetParent(Combo)
		strClassName = Space(129)
		cnt = Wrapper.GetClassName(ComboPar, strClassName, 128)
		strClassName = lcase(trim(strClassName))
		if strClassName = "combobox" then Combo = ComboPar
	end if	
	
	if (strClassName = "listbox") or (strClassName = "combobox") then
		bList = true
	end if
	
	message Svc.GetWindowText(Combo,bList)
End sub

'========================================================================================
function MakeWParam(l, h)
	MakeWParam = l or (h * 2^16)
end function	

'========================================================================================
'версия через DynamicWrapper
'==========================================================================================
Sub SelectFromComboBoxOriginalVersion()
	Const CB_GETCOUNT = &H146
	Const CB_GETLBTEXT = &H148
	Const CB_SELECTSTRING = &H14D
	Const CB_GETLBTEXTLEN = &H0149
	Const CB_ERR = -1
	
	Const WM_COMMAND =  &H111
	Const CBN_SELCHANGE  = 1
	
	Const LB_GETCOUNT = &H018B
	Const LB_SELECTSTRING = &H018C
	Const LB_GETTEXT = &H0189
	Const LBN_SELCHANGE = 1
	Const LB_GETTEXTLEN = &H018A
	Const LB_ERR = -1
	strItem = Space(129)
	strSelItem = Space(129)
	strClassName = Space(129)

	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "GetFocus",              "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "SetFocus",    "I=l",    "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetParent",   "I=l",    "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetClassName","I=lrl",  "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "SetCaretPos", "I=ll",   "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetDlgCtrlID","I=l",    "f=s", "r=l"
	
	msgGetCount = 0
	msgGetItemText = 0
	msgGetItemTextLen = 0
	msgSelectString = 0
	msgNotifySelChange = 0
	retErr = 0
	
	msgGetCount = CB_GETCOUNT
	msgGetItemText = CB_GETLBTEXT
	msgGetItemTextLen = CB_GETLBTEXTLEN
	msgSelectString = CB_SELECTSTRING
	msgNotifySelChange = CBN_SELCHANGE
	retErr = CB_ERR
	
	Combo = Wrapper.GetFocus
	
	
	bComboWithEdit = false
	bListBox = false
	cnt = Wrapper.GetClassName(Combo, strClassName, 128)
	strClassName = lcase(trim(strClassName))
	if strClassName = "edit" then
		'стиль комбобокса, с эдитом или нет
		Combo = Wrapper.GetParent(Combo)
		bComboWithEdit = true
	elseif strClassName = "listbox" then
		msgGetCount = LB_GETCOUNT
		msgGetItemText = LB_GETTEXT
		msgGetItemTextLen = LB_GETTEXTLEN
		msgSelectString = LB_SELECTSTRING
		msgNotifySelChange = LBN_SELCHANGE
		retErr = LB_ERR 
		bListBox = true
	end if
	
	'иногда в активном комбобоксе не создан курсор (caret), например если перед этим редактировать текст,
	'тогда список появляется в координатах прежнего курсора, поэтому создадим курсор сами
	if not (bComboWithEdit or bListBox) then
		'message "creating cursor"
		Wrapper.Register "USER32.DLL",   "CreateCaret", "I=llll", "f=s", "r=l"
		Wrapper.Register "USER32.DLL",   "SetCaretPos", "I=ll",   "f=s", "r=l"
		Wrapper.CreateCaret Combo, NULL, 1,12
		Wrapper.SetCaretPos 1,1
	end if	

	Wrapper.Register "USER32.DLL",   "SendMessageA", "I=llll", "f=s", "r=l"
	ComboItemsCnt = Wrapper.SendMessageA(Combo, msgGetCount, 0, 0)
	if ComboItemsCnt < 1 then exit sub
	
	ComboParent = Wrapper.GetParent(Combo)
	ComboCtrlID = Wrapper.GetDlgCtrlID(Combo)
	
	Wrapper.Register "USER32.DLL", "SendMessageA", "I=lllr", "f=s", "r=l" 
	list = ""
	Set Svc = CreateObject("Svcsvc.Service")
	for i = 0 to ComboItemsCnt-1
		'обход бага динавраппера при чтении длинных строк
		cnt = Wrapper.SendMessageA(Combo, msgGetItemTextLen, i, 0)
		if cnt > 50 then
			strItem = "</*здесь могла быть ваша реклама*/>"
			message "В позиции " & cStr(i) & " обнаружен слишком длинный для DynamicWrapper пункт длиной " & cStr(cnt) & " символов"
		else
			'strItem = Space(129)
			cnt = Wrapper.SendMessageA(Combo, msgGetItemText, i, strItem)
		end if
		list = list  & cStr(strItem) & vbCrLf
	next
	
	'вывод списка фильтервалуе в позиции комбобокса и с автошириной
	strSelItem = Svc.FilterValue(list,1+4,"",0,0,1)
	
	if strSelItem = "" then Exit Sub
	
	ret = Wrapper.SendMessageA (Combo, msgSelectString, -1, trim(strSelItem))
	Wrapper.SetFocus Combo
	
	if ret = retErr then 
		MsgBox "Выбранный пункт не найден в списке, скорее всего это баг DynamicWrapper"
	else
		'скажем диалогу что в комбобоксе или листбоксе сменился пункт
		Wrapper.Register "USER32.DLL", "SendMessageA", "I=llll", "f=s", "r=l"
		Wrapper.SendMessageA ComboParent, WM_COMMAND, MakeWParam(ComboCtrlID,msgNotifySelChange), Combo
	end if
End Sub

'========================================================================================
Private Sub Init()
	Set c = Nothing
	On Error Resume Next
	Set c = CreateObject("OpenConf.CommonServices")
	On Error GoTo 0
	If c Is Nothing Then
		Message "Не могу создать объект OpenConf.CommonServices", mRedErr
		Message "Скрипт " & SelfScript.Name & " не загружен", mInformation
		Scripts.UnLoad SelfScript.Name
	Exit Sub
	End If
	c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

'========================================================================================
'Init ''При загрузке скрипта выполняем инициализацию