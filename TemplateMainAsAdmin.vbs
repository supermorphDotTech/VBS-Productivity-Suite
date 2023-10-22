'+	Initialisierung

	Set objShell = CreateObject("Shell.Application")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	strPath = FSO.GetParentFolderName (WScript.ScriptFullName)

'+	Pfad zum Skript

	strFile = "\TemplateMain.vbs"

'+	Ausführung

	If FSO.FileExists(strPath & strFile) Then

		 objShell.ShellExecute "wscript.exe", _
			Chr(34) & strPath & strFile & Chr(34), "", "runas", 1
			
	Else

		 MsgBox "Skript mit Namen " & Chr(34) & strFile & Chr(34) & " nicht gefunden!"
		 
	End If