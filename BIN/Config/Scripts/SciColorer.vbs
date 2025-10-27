$NAME SciColorer

sub ShowSettings() 'Настройки...
	SendCommand 20060
end sub

sub ExpandAll() 'Развернуть всё
	SendCommand 20065
end sub

sub CollapseAll() 'Свернуть всё
	SendCommand 20066
end sub

sub ToggleCurrent() 'Свернуть/развернуть текущий блок (Click)
	SendCommand 20067
end sub

sub ToggleCurrentWithSubLevels() 'Развернуть/свернуть текущий вместе с внутренними (Ctrl+Click)
	SendCommand 20068
end sub

sub SelectCurrentBlock() 'Выделить текущий блок
	SendCommand 20069
end sub

sub ToggleReadOnlyMode() 'Переключение режима "только чтение"
	SendCommand 20070
end sub

sub NextModifiedLine() 'Переход к следующей модифицированной строке
	SendCommand 20071
end sub

sub PrevModifiedLine() 'Переход к предыдущей модифицированной строке
	SendCommand 20072
end sub

sub ResetModifiedLines() 'Сброс модифицированности строк
	SendCommand 20073
end sub

sub ShowBookmarksList() 'Показать список закладок модуля
	SendCommand 20074
end sub

sub ShowModifiedLinesList() 'Показать список модифицированных строк модуля
	SendCommand 20075
end sub

'Обновить содержимое редактора. Необходимо использовать, если по какой-то причине произошло визуальное нарушение текста
sub RefreshEditor() 
	SendCommand 20076
end sub

sub SetBgColor() 'Установить цвет фона выделенных строк
	SendCommand 20077
end sub

sub ReloadSettings() 'Перечитать настройки из ini файла, в случае если они были изменены вручную во время работы конфигуратора
	SciColorer.ReloadSettings()
end sub

sub ToggleViewNonPrintable() 'включить/выключить отображение непечатных символов
	SendCommand 20080
end sub

sub ShowSearchResults() 'выполнить поиск введенного в тулбаре текста и показать окно с результатами поиска
	SendCommand 20081
end sub

Sub TogleSearchResults() 'скрыть/показать панель результатов поиска (просто скрыть, без вызова самого поиска)
	Windows.PanelVisible("Результаты поиска")=Not Windows.PanelVisible("Результаты поиска")	
End Sub

sub ShowSearchResultsOfSelectedText() 'выполнить поиск текста, выделенного в текущем модуле (или слова в котором находится курсор), и показать окно с результатами поиска
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit sub
	selText = Doc.Range(Doc.SelStartLine , Doc.SelStartCol , Doc.SelEndLine , Doc.SelEndCol)
	if Trim(selText) = "" then selText = doc.CurrentWord
	if Trim(selText) = "" then exit sub
	
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(Windows.MainWnd.Hwnd,0,"AfxControlBar42",NULL)
	
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llls", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(combo,0,NULL,"Стандартная")
	
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
sub SciColorer_OnHotSpotClick() 'вызывается при клике мышью по "гиперссылке"
	
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit sub
	
	prevLine	= doc.SelStartLine
	prevCol	= doc.SelStartCol
	prevWnd = Windows.ActiveWnd

	'пробуем перейти к определению переменной или процедуры при помощи телепата
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	Wrapper.SendMessage Windows.MainWnd.HWND, &H111 ,22503, 0 'WM_COMMAND
	'message "telepat"
	if prevWnd <> Windows.ActiveWnd then 'возможно перепрыгнули в глобальник
		exit sub 
	end if
	if  (doc.SelStartLine = prevLine) and (doc.SelStartCol = prevCol) then
		'иначе пробуем перейти к определению переменной при помощи интелла
		'message "intell"
		set scr = nothing
		on error resume next
		set scr = scripts("Навигация")
		on error goto 0
		if scr is nothing then
			message "Cкрипт ""Навигация"" не установлен"	
		End If
		scr.VarDefJump()
		if  (doc.SelStartLine = prevLine) and (doc.SelStartCol = prevCol) then
			'интелл не сработал, тогда просто перейдем к ближайшей строке вверх по тексту,
			'где этой переменной что-то присваивается или будет её объявление через "Перем"
			'message "assign"
			curWord = scr.GetObjectName(doc.Range(prevLine),prevCol," .,;:|#=+-*/%?<>\()[]{}!~@$^&'""" & vbTab)
			for line = prevLine-1 to 0 step -1
				str = doc.Range(line,0)
				if CommonScripts.RegExpTest("("+curWord+"\s*=)|(перем\s+"+curWord+")",str) then
					CommonScripts.Jump line
					exit for
				end if
			next
		end if
	end if
end sub


sub SciColorer_OnLineNumbersContextMenu() 'вызывается при контекстном меню на отступе с номерами строк
	ShowBookmarksList() 'показать список закладок колорера
	'Scripts("Bookmarks").SelectBookMark() 'показать список закладок скрипта Bookmarks
end sub

'======================================================================
set obj = nothing
on error resume next
set obj = Plugins("SciColorer")
on error goto 0
if obj is nothing then
	MsgBox "SciColorer: Ошибка загрузки плагина SciColorer.dll"	
	Scripts.Unload SelfScript.Name
else
	SelfScript.AddNamedItem "SciColorer", obj, False
End If

set obj = nothing
on error resume next
set obj = CreateObject("OpenConf.CommonServices")
on error goto 0
if obj is nothing then 
	MsgBox "SciColorer: Ошибка при создании объекта OpenConf.CommonServices, зарегистрируйте библиотеку \Config\System\CommonServices.wsc"
	Scripts.UnLoad SelfScript.Name
else
	SelfScript.AddNamedItem "CommonScripts", obj, False
	CommonScripts.SetConfig(Configurator)
end if



