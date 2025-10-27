Dim flCfgWindowIsOpened
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
Init
'=======================================================================================

'=======================================================================================
Sub Configurator_OnActivateWindow(W,A)
	if not flCfgWindowIsOpened then 
		if Instr (W.Caption,"Конфигурация")=1 then 
			W.Maximized=True
			flCfgWindowIsOpened=true
			'SendCommand(32812) 'закрыть окно сообщений
			SendCommand(45098) 'закрыть синт. помощник 
		end if	
	end if	
	'message "Configurator_OnActivateWindow " & W.Caption & " " & A
End Sub

'=======================================================================================
Sub Configurator_AllPluginsInit()
	'message "Configurator_AllPluginsInit" 
	Set Wrapper = CreateObject("DynamicWrapper")
	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	WM_COMMAND =  &H111
	cmdOpenConfWindow = 33188
	Windows.MainWnd.Caption="Конфигуратор - <" & IBDir & ">"
	'Scripts("Сохранение открытых окон").OpenSavedDocs()
	
	Wrapper.SendMessage Windows.MainWnd.HWND, WM_COMMAND ,cmdOpenConfWindow, NULL
End Sub 

'=======================================================================================
Sub Configurator_ConfigWindowCreate()  
	flCfgWindowIsOpened=false
	'message "Configurator_ConfigWindowCreate"
End Sub 

'=======================================================================================
Sub OpenInDebugger()
  SendCommand(33285)
End Sub   

'=======================================================================================
Sub CallSave()
  SendCommand(57603)
End Sub

'=======================================================================================
Sub GlobalSearch()
  SendCommand(33207) 'поиск во всех текстах
End Sub

'=======================================================================================
Sub SearchSynHelp()
  SendCommand(33879) 'поиск синт. помощнике
End Sub

'=======================================================================================
Sub OpenSynHelp()
  SendCommand(33870) 'открыть синт-помощник
End Sub

'=======================================================================================
Sub CloseSynHelp()
  SendCommand(45098) 'закрыть синт. помощник
End Sub

'=======================================================================================
Sub SelectAll()
 If Windows.ActiveWnd Is Nothing Then
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     Exit Sub
  End If
  
  SendCommand(57642) 'command select all
  'doc.movecaret 0, 0, doc.LineCount - 1, doc.lineLen(doc.LineCount - 1)
End Sub

'=======================================================================================
Sub SyntaxCheck()
 If Windows.ActiveWnd Is Nothing Then
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     Exit Sub
  End If
  
  SendCommand(33297) 'синтаксический контроль
End Sub

'=======================================================================================
Sub CloseMessageWindow()
  SendCommand(32812) 'закрыть окно сообщений
End Sub


'=======================================================================================
Sub TogglePanel(PanelName)
	Windows.PanelVisible(PanelName)=Not Windows.PanelVisible(PanelName)
End Sub

'=======================================================================================
Sub ToggleSynaxHelper()
	TogglePanel "Синтакс-Помощник"
End Sub

'=======================================================================================
Sub ToggleOutPutWindow()
	TogglePanel "Окно сообщений"
End Sub

'=======================================================================================
Sub TogleSearchWindow()
	TogglePanel "Список найденных вхождений"
End Sub

'=======================================================================================
Sub TogleStdToolbar()
	TogglePanel "Стандартная"
End Sub

'=======================================================================================
Sub OpenBinDir()
	Set wshShell = CreateObject("wScript.Shell")
	Set srv = CreateObject("Svcsvc.Service")
    list = "IBDIR" & vbCrLf & "Scripts" & vbCrLf & "BINDIR" & vbCrLf & "System" & vbCrLf & "WinWord"
	dir = srv.FilterValue(list,1,"",0,0,1)
  	if Len(dir) > 0 Then
  		Select case dir  
			case "BINDIR":
				wshShell.Run """" & BinDir & """", 1
			case "IBDIR":    
				wshShell.Run """" & IBDIR & """", 1
			case "Scripts":    
				wshShell.Run """" & BinDir & "Config\Scripts""", 1
			case "System":    
				wshShell.Run """" & BinDir & "Config\System""", 1
			case "WinWord":    
				wshShell.Run "winword.exe", 1
			case "CommonServices":
				Documents.Open(BinDir & "Config\System\CommonServices.wsc")
			case "Server.ini":
				Documents.Open(IbDir & "server.ini")
		End select  
  	End If
End Sub
	
'=======================================================================================
Sub BackupMD()
	Status "Автосохранение метаданных..."
    if not MetaData.SaveMDToFile (IBDir & "1cv7_bak.md", False) then message "Не удалось сделать копию метаданных..."
    Status ""
End Sub 

'=======================================================================================
Sub OpenGM()
	Documents("ГлобальныйМодуль").Open
End Sub 

'=======================================================================================
Sub PasteMore()
	Set Doc = CommonScripts.GetTextDocIfOpened()
	if Doc is nothing then exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = ">"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 2, doc.SelEndLine, doc.SelEndCol + 2
End Sub

'=======================================================================================
Sub PasteLess()
	Set Doc = CommonScripts.GetTextDocIfOpened()
	if Doc is nothing then exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = "<"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 2, doc.SelEndLine, doc.SelEndCol + 2
End Sub

'=======================================================================================
Sub PasteCommentLine()
	Set Doc = CommonScripts.GetTextDocIfOpened()
	if Doc is nothing then exit Sub
	doc.Range(doc.SelStartLine, 0) = "//" + String(70, "=")
End Sub
