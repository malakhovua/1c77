$NAME Запуск 1С
function Run1C(mono)
{
	var shell = new ActiveXObject('WScript.Shell')
	shell.Run(CommandLine.match(/^"[^"]*"|^[^ $]*/)[0] +
		' enterprise /d"' + IBdir + '" /n"' + AppProps(appUserName) +
		'" /p"'+AppProps(15)+'"' + (mono ? '/m' : ''), 1, false)
}

function ЗапуститьНеМонопольно(){Run1C(false)}
function ЗапуститьМонопольно(){Run1C(true)}
	
function ЗапуститьНеМонопольноСДругимПользователем()
{
	var shell = new ActiveXObject('WScript.Shell')
	shell.Run(CommandLine.match(/^"[^"]*"|^[^ $]*/)[0] + ' enterprise /d"' + IBdir, 1, false);
	
}
	
