'---------------------------------------------------------------------------

' Bibliothek:

' fctXlsChkInstance			-> Prüfe, ob das Excel-Programm bereits im Hintergrund geöffnet ist
' fctXlsOpenInstWarning		-> Gibt eine Warnung aus, wenn bereits eine Excelinstanz geöffnet ist
' fctXlsWbkOpenHandler		-> Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte
' fctXlsWbkCloseHandler		-> Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte
' fctXlsWriteCmd			-> Formatiert Excel-Schreibcodes
' fctXlsReadVal				-> Zellwert eines Excelarbeitsblatts lesen
' fctXlsWriteVal			-> Zelle eines Excelarbeitsblatts mit einem Wert füllen
' fctXlsCopyWrite			-> Quell-Exceltabelle kopieren und Zellen beschreiben
' fctXlsCopyPaste			-> Excel-Zellbereich kopieren (optional mit Format)
' fctXlsNewWksht			-> Ein neues Arbeitsblatt erstellen und benennen
' fctXlsCopyWksht			-> Kopiere Arbeitsblatt innerhalb einer Exceltabelle
' fctXlsDelWksht			-> Lösche ein Arbeitsblatt einer Exceltabelle
' fctXlsPrintPDF			-> Exceltabelle als PDF drucken

'---------------------------------------------------------------------------

'Allgemeine Variablen
Public objExcel				'Objekt für das Excelprogramm
Public objWorkbookSrc		'Objekt für eine Exceltabelle (Quelle lesen)
Public objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
Public objWorkbookDest		'Objekt für eine Exceltabelle (Ziel schreiben)
Public objWorksheetDest		'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'---------------------------------------------------------------------------

'+	Deklaration fctXlsChkInstance

Public xlsInstExists		'True = Das Excelprogramm war bereits geöffnet

'+	Prüfe, ob das Excel-Programm bereits im Hintergrund geöffnet ist

	'CALL:
	'nameMsgBox					'Titel der Messagebox
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctXlsChkInstance	nameMsgBox , sysMessaging , sysDebug
	
	'RETURN:objExcel			'Objekt für das Excelprogramm
	'		xlsInstExists		'True = Das Excelprogramm war bereits geöffnet

Function fctXlsChkInstance(ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsChkInstance (" & nameMsgBox & ")"
	
	If sysDebug = True Then
		
		'Log vordefinieren
		lineLog	=	"Prüfe, ob das Excel-Programm bereits im Hintergrund geöffnet ist."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
		
'+	Starte neue Excelinstanz

	If IsObject(objExcel) = False Then
	
		Set objExcel	= CreateObject("Excel.Application")
		objExcel.EnableEvents	=	False
		objExcel.DisplayAlerts	=	False
		xlsInstExists			= False
			
		If sysDebug = True Then
			
			'Log vordefinieren
			lineLog	=	"Neue Excelinstanz wurde gestartet." & vbCrLf & _
						"xlsInstExists: " & xlsInstExists
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
		End If
		
		Set objExcel	= GetObject(, "Excel.Application")
		objExcel.EnableEvents	=	False
		objExcel.DisplayAlerts	=	False
		
'+	Übernehme vorhandene Excelinstanz

	Else
	
		Set objExcel	= GetObject(, "Excel.Application")
		objExcel.EnableEvents	=	False
		objExcel.DisplayAlerts	=	False
		xlsInstExists			= True
			
		If sysDebug = True Then
			
			'Log vordefinieren
			lineLog	=	"Vorhandene Excelinstanz wurde übernommen." & vbCrLf & _
						"xlsInstExists: " & xlsInstExists
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
		End If
	
	End If
	
'+	Parameter des Excelobjekts

	objExcel.DisplayAlerts	= False
	objExcel.Visible		= False
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsOpenInstWarning

Public xlsAppRunning			'True = Das Excel-Programm läuft

'+	Gebe eine Warnung aus, wenn bereits eine Excelinstanz geöffnet ist

	'CALL:
	'nameMsgBox					'Titel der Messagebox
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctXlsOpenInstWarning	nameMsgBox , sysMessaging , sysDebug
	
	'RETURN:VOID

Function fctXlsOpenInstWarning(ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsOpenInstWarning (" & nameMsgBox & ")"
	
	If sysDebug = True Then
		
		'Log vordefinieren
		lineLog	=	"Gebe eine Warnung aus, wenn bereits eine Excelinstanz geöffnet ist."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
		
'+	Prüfe, ob das Excel-Programm bereits läuft

	Set objExcel	= GetObject(, "Excel.Application")
	objExcel.EnableEvents	=	False
	objExcel.DisplayAlerts	=	False
	
	If Not TypeName(objExcel) = "Empty" Then
	
		xlsAppRunning = True
		
	Else
	
		xlsAppRunning = False
		
	End If

'+	Programm beenden

	If xlsAppRunning = True Then
	
		'Log vordefinieren
		lineLog	=	"Sie müssen die Excelanwendung erst schließen, bevor Sie fortfahren können." & vbCrLf & vbCrLf & _
					"Die Ausführung des Skripts wird abgebrochen..."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
		MsgBox lineLog,vbExclamation,sysProgramName
		
		fctSysEndXls	sysProgramName , _
						sysMessaging , _
						sysDebug
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctSysEndXls

'Public PLATZHALTER

'+	Script beenden

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctSysEndXls	nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug


	'RETURN:VOID
	
Function fctSysEndXls(ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked		= "fctSysEndXls (" & nameMsgBox & ")"

'+	Startbedingung

	If sysEndTriggered = False Then

'+	Error-Logs speichern

		If Err.Number <> 0 Then
		
			'Log vordefinieren
			lineLog	= 	"Das Script wurde mit Fehler(n) ausgeführt." & vbCrLf & _
						"(" & Err.Number & ") " & Err.Description
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBox , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
			boxRetVal	=	MsgBox(	"Das Script wurde mit Fehler(n) ausgeführt. Bitte kontrollieren Sie die Logs auf Hinweise." & vbCrLf & vbCrLf & _
									"Logdatei:" & vbCrLf & thisPath & "/" & nameLogfile,vbCritical,nameMsgBoxStacked)
							
		End If

'+	Laufzeit messen

		'CALL:
		'dayStart					'Startparameter Zeitformat T
		'hourStart					'Startparameter Zeitformat hh
		'minuteStart				'Startparameter Zeitformat mm
		'secondStart				'Startparameter Zeitformat ss
		'msgRuntime					'Laufzeit-Rückmeldung (True = EIN)
		'nameMsgBox					'Name des Systemnachrichtenfensters
		'sysMessaging				'Systemnachrichten: True = EIN

		fctRuntime		dayStart , _
						hourStart , _
						minuteStart , _
						secondStart , _
						False , _
						sysProgramName , _
						sysMessaging

		'RETURN:dateStamp				'Datumsformat JJJJ-MM-TT
		'		timeStamp				'Zeitformat hhmmss
		'		dateTimeStamp			'Stempel JJJJ-MM-TT_hhmmss
		'		timeStampRuntime		'Laufzeit: #d hh:mm:ss

'+	Ende der Logdatei schreiben

		lineLog		=	sysProgramName & vbCrLf & _
						"============= ENDE ============" & vbCrLf & _
						"Zeitstempel: " & dateTimeStamp & vbCrLf & _
						"Laufzeit:    " & timeStampRuntime & vbCrLf

'+	Logdatei schreiben

		'CALL:
		'lineLog			'Zu schreibende Zeile
		'nameMsgBox			'Name des Systemnachrichtenfensters
		'thisPath			'Pfad zum aktuell ausgeführten Script
		'nameLogfile		'Pfad der Logdatei
		'sysMessaging		'Systemnachrichten: True = EIN

		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

		'RETURN:VOID

		If sysQuiet = False Then
			
			boxRetVal	=	msgBox	(lineLog,vbOKOnly,sysProgramName)
		
		End If
		
'+	Fehlerspeicher loggen

		If errString <> "" AND IsEmpty(errString) = False Then

			'CALL:
			'lineLog			'Zu schreibende Zeile
			'nameMsgBox			'Name des Systemnachrichtenfensters
			'thisPath			'Pfad zum aktuell ausgeführten Script
			'nameLogfile		'Pfad der Logdatei
			'sysMessaging		'Systemnachrichten: True = EIN

			fctLogfile		errString , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							False

			'RETURN:VOID
			
		End If

		WScript.Quit

	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsWbkOpenHandler

Public xlsSourceOpen			'True = Quell-Exceltabelle soll geöffnet werden
Public xlsSourceFile			'Pfad zur Quell-Exceltabelle
Public xlsSourceWksh			'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
Public xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
Public xlsDestFile				'Pfad zur Ziel-Exceltabelle
Public xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen

'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

	'CALL:
	'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
	'xlsSourceFile				'Pfad zur Quell-Exceltabelle
	'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
	'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
	'xlsDestFile				'Pfad zur Ziel-Exceltabelle
	'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
	'nameMsgBox					'Titel der Messagebox
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctXlsWbkOpenHandler	xlsSourceOpen , xlsSourceFile , xlsSourceWksh , _
	'						xlsDestOpen , xlsDestFile , xlsDestWksh , _
	'						nameMsgBox , sysMessaging , sysDebug
	'fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
	
	'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
	'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
	'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
	'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

Function fctXlsWbkOpenHandler(ByVal xlsSourceOpen,ByVal xlsSourceFile,ByVal xlsSourceWksh,ByVal xlsDestOpen,ByVal xlsDestFile,ByVal xlsDestWksh,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsWbkOpenHandler (" & nameMsgBox & ")"
	
	If sysDebug = True Then
		
		'Log vordefinieren
		lineLog	=	"Festlegung der Quell- und Ziel-Exceltabellen-Objekte."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
	
'+	Ziel-Exceltabelle öffnen
		
	If xlsDestOpen = True Then
	
		If IsObject(objWorkbookDest) = True Then
		
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Ziel-Exceltabelle bereits geöffnet. Übernehme Fenster." & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookDest) = " & IsObject(objWorkbookDest) & vbCrLf & _
							"xlsDestFile: " & xlsDestFile
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
			
			If Not objWorkbookDest Is Nothing Then
			
				objExistent	=	True
				
			Else
			
				objExistent	=	False
				
			End If
			
		End If
		
		If IsObject(objWorkbookDest) = False OR objExistent = False Then
		
'+	Prüfe auf Existenz einer Datei

			'CALL:
			'filePath				'Pfad und Name der Datei
			'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
			'nameMsgBox				'Name des Systemnachrichtenfensters
			'sysMessaging			'Systemnachrichten: True = EIN
			'sysDebug				'Exzessives Logging

			fctChkFile	xlsDestFile , _
						True , _
						nameMsgBoxStacked , _
						sysMessaging , _
						sysDebug

			'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

			If fileExist = True Then
				
'+	Prüfe, ob das Excel-Programm bereits im Hintergrund geöffnet ist.

				'CALL:
				'nameMsgBox					'Titel der Messagebox
				'sysMessaging				'Systemnachrichten: True = EIN
				'sysDebug					'Exzessives Logging

				fctXlsChkInstance	nameMsgBox , sysMessaging , sysDebug
				
				'RETURN:objExcel			'Objekt für das Excelprogramm
				'		xlsInstExists		'True = Das Excelprogramm war bereits geöffnet

				Set objWorkbookDest	= objExcel.Workbooks.Open(xlsDestFile)

				If sysDebug = True Then
					
					'Log vordefinieren
					lineLog	=	"Ziel-Exceltabelle noch nicht geöffnet. Öffne Tabelle." & vbCrLf & vbCrLf & _
								"IsObject(objWorkbookDest) = " & IsObject(objWorkbookDest) & vbCrLf & _
								"xlsDestFile: " & xlsDestFile
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
					
				End If
				
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
			End If
			
		End If
			
'+	Objekt für Ziel-Excel-Arbeitsblatt festlegen

		If IsObject(objWorkbookDest) = True Then
		
			'Ziel-Arbeitsblatt Zwangssetzen
			If xlsDestWksh = "" Or IsEmpty(xlsDestWksh) = True Then
			
				xlsDestWksh	= 1
				
			End If
			
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Ziel-Excel-Arbeitsblatt Parameter:" & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookDest): " & IsObject(objWorkbookDest) & vbCrLf & _
							"IsObject(objWorksheetDest): " & IsObject(objWorksheetDest) & vbCrLf & _
							"xlsDestFile: " & xlsDestFile & vbCrLf & _
							"xlsDestWksh: " & xlsDestWksh
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
			Set objWorksheetDest	= Nothing
			Set objWorksheetDest	= objWorkbookDest.Worksheets(xlsDestWksh)

			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
	
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Ziel-Excel-Arbeitsblatt festelegt." & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookDest): " & IsObject(objWorkbookDest) & vbCrLf & _
							"IsObject(objWorksheetDest): " & IsObject(objWorksheetDest) & vbCrLf & _
							"xlsDestFile: " & xlsDestFile & vbCrLf & _
							"xlsDestWksh: " & xlsDestWksh
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
		End If
		
	End If
			
	fileExist	= False
		
'+	Quell-Exceltabelle öffnen

	If xlsSourceOpen = True Then
	
		If IsObject(objWorkbookSrc) = True Then
		
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Quell-Exceltabelle bereits geöffnet. Übernehme Fenster." & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookSrc) = " & IsObject(objWorkbookSrc) & vbCrLf & _
							"xlsSourceFile: " & xlsSourceFile
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
			If Not objWorkbookSrc Is Nothing Then
			
				objExistent	=	True
				
			Else
			
				objExistent	=	False
				
			End If
			
		End If
		
		If IsObject(objWorkbookSrc) = False OR objExistent = False Then
		
'+	Prüfe auf Existenz einer Datei

			'CALL:
			'filePath				'Pfad und Name der Datei
			'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
			'nameMsgBox				'Name des Systemnachrichtenfensters
			'sysMessaging			'Systemnachrichten: True = EIN
			'sysDebug				'Exzessives Logging

			fctChkFile	xlsSourceFile , _
						True , _
						nameMsgBoxStacked , _
						sysMessaging , _
						sysDebug

			'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

			If fileExist = True Then
				
'+	Prüfe, ob das Excel-Programm bereits im Hintergrund geöffnet ist.

				'CALL:
				'nameMsgBox					'Titel der Messagebox
				'sysMessaging				'Systemnachrichten: True = EIN
				'sysDebug					'Exzessives Logging

				fctXlsChkInstance	nameMsgBox , sysMessaging , sysDebug
				
				'RETURN:objExcel			'Objekt für das Excelprogramm
				'		xlsInstExists		'True = Das Excelprogramm war bereits geöffnet

				Set objWorkbookSrc	= objExcel.Workbooks.Open(xlsSourceFile)

				If sysDebug = True Then
					
					'Log vordefinieren
					lineLog	=	"Quell-Exceltabelle noch nicht geöffnet. Öffne Tabelle." & vbCrLf & vbCrLf & _
								"IsObject(objWorkbookSrc) = " & IsObject(objWorkbookSrc) & vbCrLf & _
								"xlsSourceFile: " & xlsSourceFile
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
					
				End If
				
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
			End If
			
		End If
		
'+	Objekt für Quell-Excel-Arbeitsblatt festlegen

		If IsObject(objWorkbookSrc) = True Then
		
			'Quell-Arbeitsblatt Zwangssetzen
			If xlsSourceWksh = "" Or IsEmpty(xlsSourceWksh) = True Then
			
				xlsSourceWksh	= 1
				
			End If
			
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
	
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Quell-Excel-Arbeitsblatt Parameter:" & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookSrc): " & IsObject(objWorkbookSrc) & vbCrLf & _
							"IsObject(objWorksheetSrc): " & IsObject(objWorksheetSrc) & vbCrLf & _
							"xlsSourceFile: " & xlsSourceFile & vbCrLf & _
							"xlsSourceWksh: " & xlsSourceWksh
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
			Set objWorksheetSrc	= Nothing
			Set objWorksheetSrc	= objWorkbookSrc.Worksheets(xlsSourceWksh)
			
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
	
			If sysDebug = True Then
				
				'Log vordefinieren
				lineLog	=	"Quell-Excel-Arbeitsblatt festelegt." & vbCrLf & vbCrLf & _
							"IsObject(objWorkbookSrc): " & IsObject(objWorkbookSrc) & vbCrLf & _
							"IsObject(objWorksheetSrc): " & IsObject(objWorksheetSrc) & vbCrLf & _
							"xlsSourceFile: " & xlsSourceFile & vbCrLf & _
							"xlsSourceWksh: " & xlsSourceWksh
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
			End If
			
		End If
			
	End If
	
	fileExist	= False
		
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsWbkCloseHandler

Public closeSourceOnExit		'True = Quell-Exceltabelle schließen am Ende der Funktion
Public saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
Public closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

	'CALL:
	'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
	'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox					'Titel der Messagebox
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctXlsWbkCloseHandler	closeSourceOnExit , _
	'						saveDestOnExit , _
	'						closeDestOnExit , _
	'						nameMsgBox , sysMessaging , sysDebug
	'fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
	
	'RETURN:VOID

Function fctXlsWbkCloseHandler(ByVal closeSourceOnExit,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsWbkCloseHandler (" & nameMsgBox & ")"
	
	If sysDebug = True Then
		
		'Log vordefinieren
		lineLog	=	"Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If

'+	Zieltabelle speichern

	If saveDestOnExit = True Then

		'Systemdebugging
		If sysDebug = True Then
		
			lineLog	=	"Zieldatei speichern." & vbCrLf & _
						"xlsDestFile: " & xlsDestFile
			
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
		objWorkbookDest.Save
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
	End If
		
	If closeDestOnExit = True Then

'+	Zieltabelle schließen

		'Systemdebugging
		If sysDebug = True Then
		
			lineLog	=	"Zieldatei schließen." & vbCrLf & _
						"xlsDestFile: " & xlsDestFile
			
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
		objWorkbookDest.Close
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
		Set objWorksheetDest	= Nothing
		Set objWorkbookDest		= Nothing
		xlsDestWksh		= ""
		
	End If
		
'+	Quelltabelle schließen

	If closeSourceOnExit = True Then

		'Systemdebugging
		If sysDebug = True Then
		
			lineLog	=	"Quelldatei schließen." & vbCrLf & _
						"xlsSourceFile: " & xlsSourceFile
			
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
		objWorkbookSrc.Close
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
		Set objWorksheetSrc	= Nothing
		Set objWorkbookSrc	= Nothing
		xlsSourceWksh		= ""
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsWriteCmd

Public xlsWorksheet				'Gibt das Tabellenblatt an (z.B. 2)
Public xlsCells					'Gibt die Zelle(n) an (z.B. "A5")
Public xlsCellContent			'Gibt den Zellinhalt an (z.B. "ZELLINHALT")
Public xlsTempCmd				'Temporäre Codeformatierung
Public xlsWriteCommands			'Variable, die den Rückgabewert von fctXlsWriteCmd hält

'+	Formatierung eines Excel-Schreibcodes

	'CALL:
	'xlsWorksheet			'Gibt das Tabellenblatt an (z.B. 2)
	'xlsCells				'Gibt die Zelle(n) an (z.B. "A5")
	'xlsCellContent			'Gibt den Zellinhalt an (z.B. "ZELLINHALT")
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'xlsWriteCommands	= fctXlsWriteCmd(	xlsWorksheet , xlsCells , xlsCellContent , _
	'										nameMsgBox , sysMessaging , sysDebug )

Function fctXlsWriteCmd(ByVal xlsWorksheet,ByVal xlsCells,ByVal xlsCellContent,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsWriteCmd (" & nameMsgBox & ")"
	
'+	Engaben als Code formatieren

	'Log vordefinieren
	lineLog	=	"Excel-Schreibcode wird formatiert." & vbCrLf & _
				"xlsWorksheet: " & xlsWorksheet & vbCrLf & _
				"xlsCells: " & xlsCells & vbCrLf & _
				"xlsCellContent: " & xlsCellContent
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	xlsTempCmd	= "objExcel.Worksheets(" & xlsWorksheet & ").Range(" & Chr(34) & xlsCells & Chr(34) & ").Value = " & Chr(34) & xlsCellContent & Chr(34)

	fctXlsWriteCmd	= xlsTempCmd

	'Log vordefinieren
	lineLog	=	"Excel-Schreibcode generiert." & vbCrLf & _
				"fctXlsWriteCmd: " & fctXlsWriteCmd
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsReadVal

Public xlsCell					'Gibt die Zelle an (z.B. "A5")
Public xlsReadValue				'Variable, die den Rückgabewert von fctXlsReadVal hält
Public xlsTempVal				'Hält den Rückgabewert der Zelle(n), die gelesen wird

'+	Zellwert eines Excelarbeitsblatts lesen

	'CALL:
	'xlsSourceFile			'Pfad zur Quell-Exceltabelle
	'xlsWorksheet			'Gibt das Tabellenblatt an (z.B. 2)
	'xlsCell				'Gibt die Zelle an (z.B. "A5")
	'closeSourceOnExit		'True = Quell-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'xlsReadValue	= fctXlsReadVal(xlsSourceFile , xlsWorksheet , xlsCell , closeSourceOnExit , nameMsgBox , sysMessaging , sysDebug)
	'fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

Function fctXlsReadVal(ByVal xlsSourceFile,ByVal xlsWorksheet,ByVal xlsCell,ByVal closeSourceOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsReadVal (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Zellwert eines Excelarbeitsblatts lesen." & vbCrLf & _
				"xlsSourceFile: " & xlsSourceFile & vbCrLf & _
				"xlsWorksheet: " & xlsWorksheet & vbCrLf & _
				"xlsCell: " & xlsCell
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Prüfe, ob Quelle vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsSourceFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then
		
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	True , xlsSourceFile , xlsSourceWksh , _
								False , xlsDestFile , xlsDestWksh , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

		xlsTempVal = objWorksheetSrc.Range(xlsCell).Value

		fctXlsReadVal	= xlsTempVal

		'Log vordefinieren
		lineLog	=	"Excelzelle gelesen." & vbCrLf & _
					"fctXlsReadVal: " & fctXlsReadVal
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	closeSourceOnExit , _
								False , _
								False , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID

	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsWriteVal

Public xlsWriteValue			'Wert, mit dem eine Zelle gefüllt werden soll

'+	Zelle eines Excelarbeitsblatts mit einem Wert füllen

	'CALL:
	'xlsDestFile			'Pfad zur Ziel-Exceltabelle
	'xlsWorksheet			'Gibt das Tabellenblatt an (z.B. 2)
	'xlsCell				'Gibt die Zelle an (z.B. "A5")
	'xlsWriteValue			'Wert, mit dem eine Zelle gefüllt werden soll
	'saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit		'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsWriteVal		xlsDestFile , xlsWorksheet , xlsCell , xlsWriteValue , saveDestOnExit , closeDestOnExit , nameMsgBox , sysMessaging , sysDebug
	'fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

Function fctXlsWriteVal(ByVal xlsDestFile,ByVal xlsWorksheet,ByVal xlsCell,ByVal xlsWriteValue,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsWriteVal (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Zelle eines Excelarbeitsblatts mit einem Wert füllen." & vbCrLf & _
				"xlsDestFile: " & xlsDestFile & vbCrLf & _
				"xlsWorksheet: " & xlsWorksheet & vbCrLf & _
				"xlsCell: " & xlsCell & vbCrLf & _
				"xlsWriteValue: " & xlsWriteValue
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Prüfe, ob Ziel vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsDestFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then
		
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
								True , xlsDestFile , xlsWorksheet , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

		objWorksheetDest.Range(xlsCell).Value = xlsWriteValue

		'Log vordefinieren
		lineLog	=	"Excelzelle mit Wert gefüllt." & vbCrLf & _
					"fctXlsWriteVal: " & fctXlsWriteVal
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	False , _
								saveDestOnExit , _
								closeDestOnExit , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsCopyWrite

'Public PLATZHALTER

'+	Quelldatei kopieren und Zellen beschreiben

	'CALL:
	'xlsSourceFile			'Pfad zur Quell-Exceltabelle
	'xlsDestFile			'Pfad zur Ziel-Exceltabelle
	'xlsWriteCommands		'Variable, die den Rückgabewert von fctXlsWriteCmd hält
	'saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit		'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'Overwrite				'True = Überschreibe Zieldatei, wenn sie bereits existiert
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsCopyWrite	xlsSourceFile , _
	'					xlsDestFile , _
	'					xlsWriteCommands , _
	'					saveDestOnExit , _
	'					closeDestOnExit , _
	'					nameMsgBox , _
	'					Overwrite , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsCopyWrite(ByVal xlsSourceFile,ByVal xlsDestFile,ByVal xlsWriteCommands,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal Overwrite,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsCopyWrite (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Quelldatei kopieren und Zellen beschreiben."
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Prüfe, ob Quelle vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsSourceFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Exceltabelle kopieren und schreiben:" & vbCrLf & _
					"Overwrite: " & Overwrite & vbCrLf & _
					"Quelle: " & xlsSourceFile & vbCrLf & _
					"Ziel: " & xlsDestFile & vbCrLf & vbCrLf & _
					"Befehl:" & vbCrLf & xlsWriteCommands
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
	
'+	Von Quelle kopieren und schreiben

		'CALL:
		'filePath				'Pfad und Name der Datei
		'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
		'nameMsgBox				'Name des Systemnachrichtenfensters
		'sysMessaging			'Systemnachrichten: True = EIN
		'sysDebug				'Exzessives Logging

		fctChkFile	xlsDestFile , _
					False , _
					nameMsgBoxStacked , _
					sysMessaging , _
					sysDebug

		'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

		If fileExist = False Then
			
			objFSO.CopyFile xlsSourceFile, xlsDestFile, Overwrite
				
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
			'xlsSourceFile				'Pfad zur Quell-Exceltabelle
			'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
			'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
			'xlsDestFile				'Pfad zur Ziel-Exceltabelle
			'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
									True , xlsDestFile , 1 , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
			'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
			'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
			'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'+	Exceltabelle beschreiben
			
			Execute xlsWriteCommands
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

			'Log vordefinieren
			lineLog	=	"Datei kopiert und beschrieben:" & vbCrLf & _
						"Pfad: " & xlsDestFile & vbCrLf & vbCrLf & _
						"Code:" & vbCrLf & xlsWriteCommands
							
'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
			'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
			'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkCloseHandler	False , _
									saveDestOnExit , _
									closeDestOnExit , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:VOID

'+	Overwrite-Handling
						
		ElseIf fileExist = True AND Overwrite = False Then
			
			'Log vordefinieren
			lineLog	=	"Datei existiert bereits und wurde nicht überschrieben." & vbCrLf & _
						"Pfad: " & xlsDestFile
				
'+	Exceltabelle beschreiben
				
		ElseIf fileExist = True AND Overwrite = True Then
			
			objFSO.CopyFile xlsSourceFile, xlsDestFile, Overwrite
				
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
			'xlsSourceFile				'Pfad zur Quell-Exceltabelle
			'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
			'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
			'xlsDestFile				'Pfad zur Ziel-Exceltabelle
			'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
									True , xlsDestFile , 1 , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
			'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
			'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
			'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

			Execute xlsWriteCommands
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
			'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
			'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkCloseHandler	False , _
									saveDestOnExit , _
									closeDestOnExit , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:VOID

		End If
			
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsCopyPaste

Public xlsCopyCells				'Zu kopierender Zellbereich (z.B. "A1:C3")
Public xlsPasteCells			'Zelle oben links des Zellbereichs, auf den kopiert werden soll (z.B. "A1")

'+	Excel-Zellbereich kopieren (optional mit Format)

	'CALL:
	'xlsSourceFile				'Pfad zur Quell-Exceltabelle
	'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden (z.B. "Tabelle1")
	'xlsCopyCells				'Zu kopierender Zellbereich (z.B. "A1:C3")
	'xlsDestFile				'Pfad zur Ziel-Exceltabelle (z.B. "A1")
	'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen (z.B. "Tabelle1")
	'xlsPasteCells				'Zelle oben links des Zellbereichs, auf den kopiert werden soll (z.B. "A1")
	'pasteAll					= True	'OPTION: TRUE = Alles wird eingefügt (schließt andere Optionen aus)
	'pasteAllExceptBorders		= False	'OPTION: TRUE = Alles mit Ausnahme der Rahmen wird eingefügt
	'pasteAllMergingCondF		= False	'OPTION: TRUE = Alles wird eingefügt, und bedingte Formate werden zusammengeführt
	'pasteAllUsingSrcTheme		= False	'OPTION: TRUE = Alles wird mithilfe des Quelldesigns eingefügt
	'pasteColumnWidths			= False	'OPTION: TRUE = Die kopierte Spaltenbreite wird eingefügt
	'pasteComments				= False	'OPTION: TRUE = Kommentare werden eingefügt
	'pasteFormats				= False	'OPTION: TRUE = Das kopierte Quellformat wird eingefügt
	'pasteFormulas				= False	'OPTION: TRUE = Formeln werden eingefügt
	'pasteFormAndNmbFormats		= False	'OPTION: TRUE = Formeln und Zahlenformate werden eingefügt
	'pasteValidation			= False	'OPTION: TRUE = Überprüfungen werden eingefügt
	'pasteValues				= False	'OPTION: TRUE = Werte werden eingefügt
	'pasteValAndNmbFormats		= False	'OPTION: TRUE = Werte und Zahlenformate werden eingefügt
	'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
	'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox					'Titel der Messagebox
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctXlsCopyPaste	xlsSourceFile , xlsSourceWksh , xlsCopyCells , _
	'					xlsDestFile , xlsDestWksh , xlsPasteCells , _
	'					pasteAll , _
	'					pasteAllExceptBorders , _
	'					pasteAllMergingCondF , _
	'					pasteAllUsingSrcTheme , _
	'					pasteColumnWidths , _
	'					pasteComments , _
	'					pasteFormats , _
	'					pasteFormulas , _
	'					pasteFormAndNmbFormats , _
	'					pasteValidation , _
	'					pasteValues , _
	'					pasteValAndNmbFormats , _
	'					closeSourceOnExit , _
	'					saveDestOnExit , _
	'					closeDestOnExit , _
	'					nameMsgBox , sysMessaging , sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsCopyPaste(ByVal xlsSourceFile,ByVal xlsSourceWksh,ByVal xlsCopyCells,ByVal xlsDestFile,ByVal xlsDestWksh,ByVal xlsPasteCells, _
						ByVal pasteAll,ByVal pasteAllExceptBorders,ByVal pasteAllMergingCondF,ByVal pasteAllUsingSrcTheme,ByVal pasteColumnWidths,ByVal pasteComments, _
						ByVal pasteFormats,ByVal pasteFormulas,ByVal pasteFormAndNmbFormats,ByVal pasteValidation,ByVal pasteValues,ByVal pasteValAndNmbFormats, _
						ByVal closeSourceOnExit,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsCopyPaste (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Excel-Zellbereich kopieren (optional mit Format)."
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Prüfe, ob Quelle und Ziel vorhanden

	If objFSO.Fileexists(xlsSourceFile) = True AND  objFSO.Fileexists(xlsDestFile) = True Then
	
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	True , xlsSourceFile , xlsSourceWksh , _
								True , xlsDestFile , xlsDestWksh , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'+	Übergebene Parameter loggen

		lineLog	=	"Excel-Zellbereich kopieren:" & vbCrLf & _
					"xlsSourceFile: " & xlsSourceFile & vbCrLf & _
					"xlsSourceWksh: " & xlsSourceWksh & vbCrLf & _
					"xlsCopyCells: " & xlsCopyCells & vbCrLf & _
					"xlsDestFile: " & xlsDestFile & vbCrLf & _
					"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
					"xlsPasteCells: " & xlsPasteCells
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
		If sysDebug = True Then
			
			'Log vordefinieren
			lineLog	=	"OPTIONEN:" & vbCrLf & _
						"pasteAll:               " & pasteAll & vbCrLf & _
						"pasteAllExceptBorders:  " & pasteAllExceptBorders & vbCrLf & _
						"pasteAllMergingCondF:   " & pasteAllMergingCondF & vbCrLf & _
						"pasteAllUsingSrcTheme:  " & pasteAllUsingSrcTheme & vbCrLf & _
						"pasteColumnWidths:      " & pasteColumnWidths & vbCrLf & _
						"pasteComments:          " & pasteComments & vbCrLf & _
						"pasteFormats:           " & pasteFormats & vbCrLf & _
						"pasteFormulas:          " & pasteFormulas & vbCrLf & _
						"pasteFormAndNmbFormats: " & pasteFormAndNmbFormats & vbCrLf & _
						"pasteValidation:        " & pasteValidation & vbCrLf & _
						"pasteValues:            " & pasteValues & vbCrLf & _
						"pasteValAndNmbFormats:  " & pasteValAndNmbFormats
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
						
		End If
		
'+	Bereich kopieren
		
		objWorksheetSrc.Range(xlsCopyCells).Copy
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
	
		'Systemdebugging
		If sysDebug = True AND errString = "" Then
		
			lineLog	=	"Zellbereich im Zwischenspeicher." & vbCrLf & _
						"xlsSourceWksh: " & xlsSourceWksh & vbCrLf & _
						"xlsCopyCells: " & xlsCopyCells
			
		ElseIf sysDebug = True AND errString <> "" Then
		
			lineLog	=	"Zellbereich konnte NICHT in den Zwischenspeicher kopiert werden." & vbCrLf & _
						"xlsSourceWksh: " & xlsSourceWksh & vbCrLf & _
						"xlsCopyCells: " & xlsCopyCells
						
		End If
		
		If sysDebug = True Then
		
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
'+	Bereich einfügen

		If xlsPasteAll = True Then
			
			'Systemdebugging
			If sysDebug = True Then
			
				lineLog	=	"Ausführen Parameter: xlsPasteAll" & vbCrLf & _
							"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
							"xlsPasteCells: " & xlsPasteCells
				
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
								
			End If
			
			'Alles einfügen (schließt andere Optionen aus)
			objWorksheetDest.Range(xlsPasteCells).PasteSpecial -4104, True, False
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
			'Log vordefinieren
			lineLog	=	"Alle Zellinformationen übertragen." & vbCrLf & _
						"xlsPasteAll: " & xlsPasteAll
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
						
		ElseIf xlsPasteAll = False Then
			
			If pasteAllExceptBorders = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteAllExceptBorders" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
						
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Alles mit Ausnahme der Rahmen einfügen
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 7, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Alles mit Ausnahme der Rahmen eingefügt." & vbCrLf & _
								"pasteAllExceptBorders: " & pasteAllExceptBorders
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteAllMergingCondF = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteAllMergingCondF" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Alles wird eingefügt, und bedingte Formate werden zusammengeführt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 14, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Alles eingefügt, und bedingte Formate zusammengeführt." & vbCrLf & _
								"pasteAllMergingCondF: " & pasteAllMergingCondF
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteAllUsingSrcTheme = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteAllUsingSrcTheme" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Alles wird mithilfe des Quelldesigns eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 13, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Alles mithilfe des Quelldesigns eingefügt." & vbCrLf & _
								"pasteAllUsingSrcTheme: " & pasteAllUsingSrcTheme
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteColumnWidths = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteColumnWidths" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Die kopierte Spaltenbreite wird eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 8, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Die kopierte Spaltenbreite wurde übertragen." & vbCrLf & _
								"pasteColumnWidths: " & pasteColumnWidths
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteComments = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteComments" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Kommentare werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial -4144, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Kommentare eingefügt." & vbCrLf & _
								"pasteComments: " & pasteComments
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteFormats = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteFormats" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Das kopierte Quellformat wird eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial -4122, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Das kopierte Quellformat wurde eingefügt." & vbCrLf & _
								"pasteFormats: " & pasteFormats
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteFormulas = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteFormulas" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Formeln werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial -4123, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Formeln eingefügt." & vbCrLf & _
								"pasteFormulas: " & pasteFormulas
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteFormAndNmbFormats = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteFormulas" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Formeln und Zahlenformate werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 11, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Formeln und Zahlenformate eingefügt." & vbCrLf & _
								"pasteFormAndNmbFormats: " & pasteFormAndNmbFormats
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteValidation = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteValidation" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Überprüfungen werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 6, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Überprüfungen eingefügt." & vbCrLf & _
								"pasteValidation: " & pasteValidation
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteValues = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteValues" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Werte werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial -4163, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Werte eingefügt." & vbCrLf & _
								"pasteValues: " & pasteValues
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
			If pasteValAndNmbFormats = True Then
				
				'Systemdebugging
				If sysDebug = True Then
				
					lineLog	=	"Ausführen Parameter: pasteValAndNmbFormats" & vbCrLf & _
								"xlsDestWksh: " & xlsDestWksh & vbCrLf & _
								"xlsPasteCells: " & xlsPasteCells
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
									
				End If
				
				'Werte und Zahlenformate werden eingefügt
				objWorksheetDest.Range(xlsPasteCells).PasteSpecial 12, True, False
				fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
				
				If errString = "" Then
					
					'Log vordefinieren
					lineLog	=	"Werte und Zahlenformate eingefügt." & vbCrLf & _
								"pasteValAndNmbFormats: " & pasteValAndNmbFormats
					
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
				End If
				
			End If
			
		End If

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	closeSourceOnExit , _
								saveDestOnExit , _
								closeDestOnExit , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID

'+	Quelle existiert nicht

	Else
	
		'Log vordefinieren
		lineLog	=	"Die Datei wurde nicht gefunden!" & vbCrLf & _
					"xlsSourceFile ex.: " & objFSO.Fileexists(xlsSourceFile) & vbCrLf & _
					"xlsDestFile ex.: " & objFSO.Fileexists(xlsDestFile) & vbCrLf & vbCrLf & _
					"'OK' zum fortfahren oder 'Abbrechen' zum beenden drücken..."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
		
'+	Nutzerabfrage

		boxRetVal = MsgBox(	lineLog,vbOKCancel,sysProgramName & " - Quelle nicht gefunden!")
						
'+	Programm fortfahren

		If boxRetVal = vbOK Then
		
			'Log vordefinieren
			lineLog	=	"Programm fortfahren" & vbCrLf & _
						"Nutzereingabe: " & boxRetVal
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
				
'+	Programm beenden

		ElseIf boxRetVal = vbCancel Then
		
			'Log vordefinieren
			lineLog	=	"Programm beenden" & vbCrLf & _
						"Nutzereingabe: " & boxRetVal
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
			fctSysEnd		sysProgramName , _
							sysMessaging , _
							sysDebug
				
		End If
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctXlsNewWksht

Public wkshtNameNew		'Name des Neuen Arbeitsblatts
Public wkshtPosition	'Neue Position des Arbeitsblatts in der Exceltabelle

'+	Füge Arbeitsblatt zu einer Exceltabelle hinzu

	'CALL:
	'xlsDestFile			'Pfad zur Ziel-Exceltabelle
	'wkshtNameNew			'Name des Neuen Arbeitsblatts
	'wkshtPosition			'Arbeitsblat an Position ## verschieben
	'saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit		'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsNewWksht		xlsDestFile , _
	'					wkshtNameNew , _
	'					wkshtPosition , _
	'					saveDestOnExit , _
	'					closeDestOnExit , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsNewWksht(ByVal xlsDestFile,ByVal wkshtNameNew,ByVal wkshtPosition,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsNewWksht (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Öffne Exceltabelle und erstelle neues Arbeitsblatt." & vbCrLf & _
				"xlsDestFile: " & xlsDestFile
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Prüfe, ob Ziel vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsDestFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Arbeitsblatt hinzufügen." & vbCrLf & _
					"xlsDestFile: " & xlsDestFile & vbCrLf & _
					"wkshtNameNew: " & wkshtNameNew
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
					
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
								True , xlsDestFile , xlsDestWksh , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'+	Arbeitsblatt am Ende hinzufügen

		objExcel.Sheets.Add.Name = wkshtNameNew
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

		'Log vordefinieren
		lineLog	=	"Excel-Arbeitsblatt hinzugefügt." & vbCrLf & _
					"wkshtNameNew: " & wkshtNameNew
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

'+	Verschiebe das erstellte Arbeitsblatt ans Ende

		objExcel.Sheets(wkshtNameNew).Move , objExcel.Sheets(wkshtPosition)
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

		If sysDebug = True Then
			'Log vordefinieren
			lineLog	=	"Excel-Arbeitsblatt verschoben." & vbCrLf & _
						"wkshtNameNew: " & wkshtNameNew & vbCrLf & _
						"Position: " & wkshtPosition
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
		End If

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	False , _
								saveDestOnExit , _
								closeDestOnExit , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID

	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function
						
'---------------------------------------------------------------------------

'+	Deklaration fctXlsCopyWksht

'Public PLATZHALTER

'+	Kopiere Arbeitsblatt innerhalb einer Exceltabelle

	'CALL:
	'xlsDestFile			'Pfad zur Ziel-Exceltabelle
	'wkshtNameOrg			'Name des zu kopierenden Arbeitsblatts
	'wkshtNameNew			'Name des Neuen Arbeitsblatts
	'wkshtPosition			'Arbeitsblat an Position ## verschieben
	'saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit		'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsCopyWksht	xlsDestFile , _
	'					wkshtNameOrg , _
	'					wkshtNameNew , _
	'					wkshtPosition , _
	'					saveDestOnExit , _
	'					closeDestOnExit , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsCopyWksht(ByVal xlsDestFile,ByVal wkshtNameOrg,ByVal wkshtNameNew,ByVal wkshtPosition,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsCopyWksht (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Kopiere Arbeitsblatt innerhalb einer Exceltabelle." & vbCrLf & _
				"xlsDestFile: " & xlsDestFile
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Prüfe, ob Ziel vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsDestFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Arbeitsblatt kopieren." & vbCrLf & _
					"xlsDestFile: " & xlsDestFile & vbCrLf & _
					"wkshtNameOrg: " & wkshtNameOrg & vbCrLf & _
					"wkshtNameNew: " & wkshtNameNew & vbCrLf & _
					"wkshtPosition: " & wkshtPosition
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
		
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
								True , xlsDestFile , xlsDestWksh , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'+	Arbeitsblatt kopieren und mit anderem Namen speichern

		objExcel.Sheets(wkshtNameOrg).Copy , objExcel.Sheets(wkshtNameOrg)
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

		objExcel.ActiveSheet.Name		= wkshtNameNew
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

'+	Verschiebe das kopierte Arbeitsblatt an die vorgegebene Position

		objExcel.Sheets(wkshtNameNew).Move , objExcel.Sheets(wkshtPosition)
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

		If sysDebug = True Then
		
			'Log vordefinieren
			lineLog	=	"Excel-Arbeitsblatt verschoben." & vbCrLf & _
						"wkshtNameNew: " & wkshtNameNew & vbCrLf & _
						"Position: " & wkshtPosition
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
		End If

		'Log vordefinieren
		lineLog	=	"Excel-Arbeitsblatt kopiert." & vbCrLf & _
					"wkshtNameNew: " & wkshtNameNew
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
							
'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	False , _
								saveDestOnExit , _
								closeDestOnExit , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID

	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function
						
'---------------------------------------------------------------------------

'+	Deklaration fctXlsDelWksht

Public wkshtToDelete		'Name des zu löschenden Arbeitsblatts

'+	Lösche ein Arbeitsblatt einer Exceltabelle

	'CALL:
	'xlsDestFile			'Pfad zur Ziel-Exceltabelle
	'wkshtToDelete			'Name des zu löschenden Arbeitsblatts
	'saveDestOnExit			'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
	'closeDestOnExit		'True = Ziel-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsDelWksht		xlsDestFile , _
	'					wkshtToDelete , _
	'					saveDestOnExit , _
	'					closeDestOnExit , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsDelWksht(ByVal xlsDestFile,ByVal wkshtToDelete,ByVal saveDestOnExit,ByVal closeDestOnExit,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsDelWksht (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Öffne Exceltabelle." & vbCrLf & _
				"xlsDestFile: " & xlsDestFile
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Prüfe, ob Ziel vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsDestFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Arbeitsblatt löschen." & vbCrLf & _
					"xlsDestFile: " & xlsDestFile & vbCrLf & _
					"wkshtToDelete: " & wkshtToDelete
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
					
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
		'xlsSourceFile				'Pfad zur Quell-Exceltabelle
		'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
		'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
		'xlsDestFile				'Pfad zur Ziel-Exceltabelle
		'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkOpenHandler	False , xlsSourceFile , xlsSourceWksh , _
								True , xlsDestFile , xlsDestWksh , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
		'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
		'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
		'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

'+	Arbeitsblatt löschen

		objExcel.Sheets(wkshtToDelete).Delete
		fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

		If sysDebug = True Then
			'Log vordefinieren
			lineLog	=	"Excel-Arbeitsblatt gelöscht." & vbCrLf & _
						"wkshtToDelete: " & wkshtNameNew
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
		End If

'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

		'CALL:
		'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
		'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
		'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsWbkCloseHandler	False , _
								saveDestOnExit , _
								closeDestOnExit , _
								nameMsgBoxStacked , sysMessaging , sysDebug
		fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
		
		'RETURN:VOID

	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function
						
'---------------------------------------------------------------------------

'+	Deklaration fctXlsPrintPDF

Public xlsPdfPath			'Pfad zum PDF-Ausgabedokument
Public printFrom			'Erste zu druckende Seite
Public printTo				'Letzte zu druckende Seite

'+	Exceltabelle als PDF drucken

	'CALL:
	'xlsSourceFile			'Pfad zur Quell-Exceltabelle
	'xlsPdfPath				'Pfad zum PDF-Ausgabedokument
	'printFrom				'Erste zu druckende Seite
	'printTo				'Letzte zu druckende Seite
	'closeSourceOnExit		'True = Quell-Exceltabelle schließen am Ende der Funktion
	'nameMsgBox				'Titel der Messagebox
	'Overwrite				'True = Überschreibe Zieldatei, wenn sie bereits existiert
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctXlsPrintPDF		xlsSourceFile , _
	'					xlsPdfPath , _
	'					printFrom , _
	'					printTo , _
	'					closeSourceOnExit , _
	'					nameMsgBox , _
	'					Overwrite , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctXlsPrintPDF(ByVal xlsSourceFile,ByVal xlsPdfPath,ByVal printFrom,ByVal printTo,ByVal closeSourceOnExit,ByVal nameMsgBox,ByVal Overwrite,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctXlsPrintPDF (" & nameMsgBox & ")"
	
	'Log vordefinieren
	lineLog	=	"Öffne und exportiere Exceltabelle. Excelinstanz geöffnet."
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
'+	Standarddrucker festlegen

	Set WSHNetwork = CreateObject("WScript.Network")
	WSHNetwork.SetDefaultPrinter "Microsoft Print to PDF"
	
'+	Prüfe, ob Quelle vorhanden

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	xlsSourceFile , _
				True , _
				nameMsgBoxStacked , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Exceltabelle öffnen und als PDF exportieren:" & vbCrLf & _
					"Overwrite = " & Overwrite & vbCrLf & _
					"Quelle: " & xlsSourceFile & vbCrLf & _
					"Export: " & xlsPdfPath
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
'+	Exceltabelle als PDF exportieren

		'CALL:
		'filePath				'Pfad und Name der Datei
		'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
		'nameMsgBox				'Name des Systemnachrichtenfensters
		'sysMessaging			'Systemnachrichten: True = EIN
		'sysDebug				'Exzessives Logging

		fctChkFile	xlsPdfPath , _
					False , _
					nameMsgBoxStacked , _
					sysMessaging , _
					sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

'+	Overwrite-Handling
						
		If fileExist = True AND Overwrite = False Then
			
			'Log vordefinieren
			lineLog	=	"Datei existiert bereits und wurde nicht überschrieben." & vbCrLf & _
						"Pfad: " & xlsPdfPath & vbCrLf & _
						"Overwrite: " & Overwrite
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
'+	PDF überschreiben
				
		ElseIf fileExist = False OR Overwrite = True Then
			
'+	Regelt die Festlegung der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'xlsSourceOpen				'True = Quell-Exceltabelle soll geöffnet werden
			'xlsSourceFile				'Pfad zur Quell-Exceltabelle
			'xlsSourceWksh				'Arbeitsblatt, auf dem sich die zu kopierenden Zellen befinden
			'xlsDestOpen				'True = Ziel-Exceltabelle soll geöffnet werden
			'xlsDestFile				'Pfad zur Ziel-Exceltabelle
			'xlsDestWksh				'Arbeitsblatt, auf das die Zellen kopiert werden sollen
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkOpenHandler	True , xlsSourceFile , 1 , _
									False , xlsDestFile , 1 , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:objWorkbookDest		'Objekt für ein Excel Arbeitsblatt (Ziel schreiben)
			'		objWorkbookSrc		'Objekt für ein Excel Arbeitsblatt (Quelle lesen)
			'		objWorksheetSrc		'Objekt für ein Excel-Arbeitsblatt (Quelle lesen)
			'		objWorksheetDest	'Objekt für ein Excel-Arbeitsblatt (Ziel schreiben)

			'PARAMETER:
			'Type (0=xlTypePDF, 1=xltypexps)
			'FileName
			'Quality (0=xlQualityStandard, 1=xlQualityMinimum)
			'IncludeDocProperties (True, False)
			'IgnorePrintAreas (True, False)
			'From (INT)
			'To (INT)
			'OpenAfterPublish (True, False)
			'FixedFormatExtClassPtr (Zeiger auf die FixedFormatExt-Klasse)
			
			objWorkbookSrc.ExportAsFixedFormat	0 , _
												xlsPdfPath , _
												0 , _
												True , _
												False , _
												printFrom , _
												printTo , _
												True
			fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug
		
			If fileExist = False Then
			
				'Log vordefinieren
				lineLog	=	"Exceltabelle veröffentlicht." & vbCrLf & _
							"Pfad: " & xlsPdfPath
							
			ElseIf fileExist = True AND Overwrite = True Then
			
				'Log vordefinieren
				lineLog	=	"Exceltabelle veröffentlicht. Datei wurde überschrieben." & vbCrLf & _
							"Pfad: " & xlsPdfPath
							
			End If
							
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
'+	Regelt das Schließen/Speichern der Quell- und Ziel-Exceltabellen-Objekte

			'CALL:
			'closeSourceOnExit			'True = Quell-Exceltabelle schließen am Ende der Funktion
			'saveDestOnExit				'True = Ziel-Exceltabelle wird am Ende der Funktion gespeichert
			'closeDestOnExit			'True = Ziel-Exceltabelle schließen am Ende der Funktion
			'nameMsgBox					'Titel der Messagebox
			'sysMessaging				'Systemnachrichten: True = EIN
			'sysDebug					'Exzessives Logging

			fctXlsWbkCloseHandler	True , _
									False , _
									False , _
									nameMsgBoxStacked , sysMessaging , sysDebug
			fctErrorHandling		nameMsgBoxStacked , sysMessaging , sysDebug
			
			'RETURN:VOID

		End If

	End If

	nameMsgBoxStacked	= nameMsgBox
		
End Function
