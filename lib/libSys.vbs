'---------------------------------------------------------------------------

' Bibliothek:

' fctAdminHandling		-> Ermittelt Admin-Status und vergleicht mit Vorgabe
' fctDateTimeNow		-> Ermittelt Datum und Uhrzeit
' fctRuntime			-> Misst die Laufzeit zwischen jetzt und einem anderen Wert
' fctSystemInit			-> Ermittelt Systemparameter wie Nutzername und Ausführungspfad
' fctLogfile			-> Schreibt die Logdatei und generiert Meldungen
' fctSysEnd				-> Beendet das Skript
' fctChkDir				-> Prüft auf Existenz eines Verzeichnisses und legt es ggf. an
' fctChkFile			-> Prüft auf Existenz einer Datei
' fctCopyFile			-> Datei kopieren
' fctDeleteFile			-> Löschen einer Datei
' BETA fctSelFolderDialog	-> Öffnet einen Auswahldialog (kann auch Ordner erstellen)
' fctInputBoxTest		-> Fenster für Benutzereingabe mit Autokorrektur aufrufen
' fctReplaceString		-> Durchsucht eine Textdatei auf eine Zeichenkette und ersetzt diese
' fctReplaceCharacter	-> Durchsucht einen String auf Zeichen und ersetzt diese
' fctRemoveSpecial		-> Durchsucht einen String auf Sonderzeichen und entfernt diese
' fctRegEx				-> Eine Zeichenfolge auf Muster untersuchen
' fctCmdWrite			-> Schreibt eine Batchdatei (bzw. Textdatei)
' fctCreateFolder		-> Ordner erstellen
' fctTestFolder			-> Prüfe Ordner auf Existenz, lege Overwrite fest und erstelle Verzeichnis bedingt
' fctWeekNumber			-> Gibt die Kalenderwoche eines Datums zurück
' fctErrorHandling		-> Prüft ob ein ERROR registriert wurde und speichert ihn

'---------------------------------------------------------------------------

'+	Deklaration Allgemeiner Variablen

Public objFSO												'File System Object
Set objFSO			= Wscript.CreateObject("Scripting.FileSystemObject")
Public objShell												'Shell Object
Set objShell		= CreateObject("WScript.Shell")
Public objGroup												'WinNT-Gruppenobjekt
Public objNetwork											'WinNT-Netzwerkobjekt
Set objNetwork		= CreateObject("Wscript.Network")

Public boxRetVal											'Ergebnis einer Nutzereingabe
Public loopActive											'Schleifenkriterium
Public tempString											'Temporäre Variable zur Stringverarbeitung
Public tempCounter											'Temporäre Variable zur Zählung (Integer)
Public tempBool												'Temporäre binäre Variable

Public Overwrite											'True = Überschreibe Zieldatei, wenn sie bereits existiert

Set objRegEx		= CreateObject("VBScript.RegExp")		'Stringmuster prüfen
Public regExpMatches										'Methode zur Ausführung

'---------------------------------------------------------------------------

'+	Deklaration fctAdminHandling

Public fctNeedsAdmin					'Administratorrechte benötigt
Public isAdmin							'Aktuelle Administratorberechtigung
Public userPermission					'Aktuelle Berechtigungsstufe
Public testPermOK						'True = Berechtigungen vorhanden / OK
Public nameMsgBoxStacked				'Name des Systemnachrichtenfensters (zusammengesetzt)

'+	Admin-Status erfassen und vergleichen mit Vorgabe

	'CALL:
	'sysNeedsAdmin		'Global: Administratorrechte benötigt
	'fctNeedsAdmin		'Administratorrechte benötigt
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctAdminHandling	sysNeedsAdmin , _
	'					fctNeedsAdmin , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug

	'RETURN:	isAdmin				'Wahrheitswert über aktuelle Administratorrechte
	'			userPermission		'Aktuelle Berechtigungsstufe
	'			testPermOK			'True = Berechtigungen vorhanden / OK

Function fctAdminHandling(ByVal sysNeedsAdmin,ByVal fctNeedsAdmin, byVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctAdminHandling (" & nameMsgBox & ")"

	If IsEmpty(fctNeedsAdmin) = True Then
	
		fctNeedsAdmin	=	False
		
	End If
	
	testPermOK	= False

'+	Aktuelle Administratorrechte ermitteln

	Set objExec		= objShell.Exec("WHOAMI /Groups")
	
	tempString		= objExec.StdOut.ReadAll( )
	tempInt			= objExec.ExitCode
	
	If InStr(tempString, "12288") Then
	
		isAdmin			= True
		userPermission	= "Elevated"
		
	Else
	
		isAdmin		= False
		userPermission	= "NOT elevated"
		
	End If
	
'+	Funktionsstart loggen	

	'Log vordefinieren
	lineLog	=	"Administratorrechte ermitteln" & vbCrLf  & vbCrLf & _
				"sysNeedsAdmin: " & sysNeedsAdmin & vbCrLf & _
				"fctNeedsAdmin: " & fctNeedsAdmin & vbCrLf & _
				"isAdmin: " & isAdmin & vbCrLf & _
				"userPermission: " & userPermission
				
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Erhöhte Rechte erforderlich	

	If sysNeedsAdmin = True Or fctNeedsAdmin = True Then
		
		'Log vordefinieren
		lineLog	=	"Administratorrechte benötigt" & vbCrLf & _
					"sysNeedsAdmin: " & sysNeedsAdmin & vbCrLf & _
					"fctNeedsAdmin: " & fctNeedsAdmin
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
'+	Teste Berechtigung

		If isAdmin = True Then
			
			testPermOK	= True
		
			'Log vordefinieren
			lineLog	=	"Erhöhte Rechte vorhanden" & vbCrLf  & vbCrLf & _
						"isAdmin: " & isAdmin & vbCrLf & _
						"userPermission: " & userPermission	& vbCrLf & _
						"testPermOK: " & testPermOK
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
						
		ElseIf isAdmin = False Then
		
			testPermOK	= False
		
			'Log vordefinieren
			lineLog	=	"Abbruch! Erhöhte Rechte benötigt."  & vbCrLf & vbCrLf & _
						"isAdmin: " & isAdmin & vbCrLf & _
						"userPermission: " & userPermission	& vbCrLf & _
						"testPermOK: " & testPermOK
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
					
			MsgBox	lineLog
			
			fctSysEnd		sysProgramName , _
							sysMessaging , _
							sysDebug

		End If
		
'+	Erhöhte Rechte nicht erforderlich	

	Else
	
		'Log vordefinieren
		lineLog	=	"Administratorrechte nicht benötigt"  & vbCrLf & vbCrLf & _
					"sysNeedsAdmin: " & sysNeedsAdmin & vbCrLf & _
					"fctNeedsAdmin: " & fctNeedsAdmin
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
		
		testPermOK	= True
		
	End If

	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctDateTimeNow

Public thisDay						'Datumsformat TT
Public thisMonth					'Datumsformat MM
Public dateStamp					'Datumsformat JJJJ-MM-TT
Public hourNow						'Zeitformat hh
Public minuteNow					'Zeitformat mm
Public secondNow					'Zeitformat ss
Public timeStamp					'Zeitformat hh:mm:ss
Public dateTimeStamp				'Stempel JJJJ-MM-TT_hhmmss

'+	Datum und Uhrzeit erfassen

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN

	'fctDateTimeNow		nameMsgBox , _
	'					sysMessaging

	'RETURN:dateStamp				'Datumsformat JJJJ-MM-TT
	'		timeStamp				'Zeitformat hhmmss
	'		dateTimeStamp			'Stempel JJJJ-MM-TT_hhmmss

Function fctDateTimeNow(ByVal nameMsgBox,ByVal sysMessaging)
	
	nameMsgBoxStacked		= "fctDateTimeNow (" & nameMsgBox & ")"
		
	If day(now) < 10 Then
		thisDay = "0" & day(now)
	Else
		thisDay = day(now)
	End If

	If Month(now) < 10 Then
		thisMonth = "0" & Month(now)
	Else
		thisMonth = Month(now)
	End If

'+	Datum zusammensetzen

	dateStamp = year(now) & "-" & thisMonth & "-" & thisDay

	If hour(now) < 10 Then
		hourNow = "0" & hour(now)
	Else
		hourNow = hour(now)
	End If

	If minute(now) < 10 Then
		minuteNow = "0" & minute(now)
	Else
		minuteNow = minute(now)
	End If

	If second(now) < 10 Then
		secondNow = "0" & second(now)
	Else
		secondNow = second(now)
	End If

'+	Zeit zusammensetzen

	timeStamp = hourNow & minuteNow & secondNow

'+	Datums- und Zeitstempel erstellen

	dateTimeStamp	=	dateStamp & "_" & timeStamp
	
	'Script-Benachrichtigungen
	If sysMessaging	= True Then
	
		boxRetVal	=	MsgBox("dateStamp = " & dateStamp & vbCrLf & _
							"timeStamp = " & timeStamp & vbCrLf & _
							"dateTimeStamp = " & dateTimeStamp & vbCrLf & vbCrLf & _
							"Zum beenden 'Abbrechen' drücken",vbOKCancel,nameMsgBoxStacked)
							
		If boxRetVal = vbCancel Then
		
			WScript.Quit
			
		End If
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctRuntime

Public dayRuntime			'Anzahl Tage
Public hourRuntime			'Zeitformat hh
Public minuteRuntime		'Zeitformat mm
Public secondRuntime		'Zeitformat ss
Public timeStampRuntime		'Laufzeit: #d hh:mm:ss
Public msgRuntime			'Legt fest, ob die Laufzeit gemeldet wird (True = EIN)

'+	Laufzeit messen

	'CALL:
	'dayStart			'Startparameter Zeitformat T
	'hourStart			'Startparameter Zeitformat hh
	'minuteStart		'Startparameter Zeitformat mm
	'secondStart		'Startparameter Zeitformat ss
	'msgRuntime			'Laufzeit-Rückmeldung (True = EIN)
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN

	'fctRuntime		dayStart , _
	'				hourStart , _
	'				minuteStart , _
	'				secondStart , _
	'				msgRuntime , _
	'				nameMsgBox , _
	'				sysMessaging

	'RETURN:dateStamp				'Datumsformat JJJJ-MM-TT
	'		timeStamp				'Zeitformat hhmmss
	'		dateTimeStamp			'Stempel JJJJ-MM-TT_hhmmss
	'		timeStampRuntime		'Laufzeit: #d hh:mm:ss

Function fctRuntime(ByVal dayStart,ByVal hourStart,ByVal minuteStart,ByVal secondStart,ByVal msgRuntime,ByVal nameMsgBox,ByVal sysMessaging)
	
	nameMsgBoxStacked		= "fctRuntime (" & nameMsgBox & ")"
			
'+	Zeitdifferenz messen
		
	dayRuntime				= 0
	hourRuntime				= hour(now)
	minuteRuntime			= minute(now)
	secondRuntime			= second(now)

'+	Laufzeit (Sekunden) berechnen

	If secondRuntime >= CInt(secondNow) Then
		secondRuntime 	= secondRuntime - CInt(secondNow)
	Else
		secondRuntime	= secondRuntime + 60 - CInt(secondNow)
		minuteRuntime	= minuteRuntime - 1
	End If
	
	If secondRuntime < 0 Then
		secondRuntime	= 60 + secondRuntime
	End If
	
'+	Laufzeit (Minuten) berechnen

	If minuteRuntime >= CInt(minuteNow) Then
		minuteRuntime 	= minuteRuntime - CInt(minuteNow)
	Else
		minuteRuntime 	= minuteRuntime + 60 - CInt(minuteNow)
		hourRuntime		= hourRuntime	- 1
	End If
	
	If minuteRuntime < 0 Then
		minuteRuntime	= 60 + minuteRuntime
	End If
	
'+	Laufzeit (Stunden) berechnen

	If hourRuntime >= CInt(hourNow) Then
		hourRuntime 	= hourRuntime - CInt(hourNow)
	Else
		hourRuntime 	= hourRuntime + 24 - CInt(hourNow)
	End If
	
	If hourRuntime < 0 Then
		hourRuntime		= 24 + hourRuntime
	End If
	
'+	Laufzeit (Tage) berechnen

	dayRuntime	= day(now) - thisDay
	
'+	Laufzeit (Sekunden) formatieren

	If secondRuntime < 10 Then
		secondRuntime	 = "0" & secondRuntime
	End If

'+	Laufzeit (Minuten) formatieren

	If minuteRuntime < 10 Then
		minuteRuntime	= "0" & minuteRuntime
	End If

'+	Laufzeit (Stunden) formatieren

	If hourRuntime < 10 Then
		hourRuntime 	= "0" & hourRuntime
	End If
	
'+	Zeit zusammensetzen

	timeStampRuntime = dayRuntime & "d " & hourRuntime & ":" & minuteRuntime & ":" & secondRuntime

	'Log vordefinieren
	lineLog	=	"secondRuntime: " & secondRuntime & vbCrLf & _
				"minuteRuntime: " & minuteRuntime & vbCrLf & _
				"hourRuntime: " & hourRuntime & vbCrLf & _
				"dayRuntime: " & dayRuntime & vbCrLf & _
				"timeStampRuntime: " & timeStampRuntime
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Datum und Uhrzeit erfassen

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN

	fctDateTimeNow		nameMsgBox , _
						sysMessaging

	'RETURN:dateStamp				'Datumsformat JJJJ-MM-TT
	'		timeStamp				'Zeitformat hhmmss
	'		dateTimeStamp			'Stempel JJJJ-MM-TT_hhmmss

	'Log vordefinieren
	lineLog	=	"dateStamp: " & dateStamp & vbCrLf & _
				"timeStamp: " & timeStamp & vbCrLf & _
				"dateTimeStamp: " & dateTimeStamp
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					False
						
'+	Ausgabe der Laufzeit

	If msgRuntime	= True Then
	
		boxRetVal	=	MsgBox("Zeitstempel: " & dateTimeStamp & vbCrLf & _
							"Laufzeit: " & timeStampRuntime & vbCrLf & vbCrLf & _
							"Zum beenden 'Abbrechen' drücken",vbOK,nameMsgBoxStacked)
							
		If boxRetVal = vbCancel Then
		
			WScript.Quit
			
		End If
		
	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctSystemInit

Public WinComputerName				'Computername
Public WinUserName					'Benutzername der aktiven Sitzung
Public WinUserPath					'Benutzerverzeichnis
Public thisFile						'Pfad des aktuell ausgeführten Scripts
Public thisPath						'Pfad zum aktuell ausgeführten Script
Public dirTemp						'Pfad des temporären Verzeichnisses

'+	Systemparameter Initialisieren

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctSystemInit	nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:WinComputerName			'Computername
	'		WinUserName				'Benutzername der aktiven Sitzung
	'		WinUserPath				'Benutzerverzeichnis
	'		thisFile				'Pfad des aktuell ausgeführten Scripts
	'		thisPath				'Pfad zum aktuell ausgeführten Script
	'		dirTemp					'Pfad des temporären Verzeichnisses

Function fctSystemInit(ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	nameMsgBoxStacked		= "fctSystemInit (" & nameMsgBox & ")"

'+	Rechnername erfassen

	WinComputerName = CreateObject("Wscript.Network").ComputerName
		
'+	Benutzername erfassen

	WinUserName		= CreateObject("WScript.Network").UserName

'+	Benutzerverzeichnis erfassen

	WinUserPath		= "C:\Users\" & WinUserName

'+	Scriptname erfassen

	thisFile		= WScript.ScriptFullName

'+	Scriptpfad erfassen

	thisPath		= CreateObject("Scripting.FileSystemObject").GetParentFolderName(thisFile)

'+	Temporären Ordner ermitteln

	dirTemp			= WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
	
'+	Script-Benachrichtigungen

	If sysMessaging	= True Then
	
		boxRetVal	=	MsgBox("WinComputerName = " & WinComputerName & vbCrLf & _
								"WinUserName = " & WinUserName & vbCrLf & _
								"WinUserPath = " & WinUserPath & vbCrLf & _
								"thisFile = " & thisFile & vbCrLf & _
								"thisPath = " & thisPath & vbCrLf & vbCrLf & _
								"Zum beenden 'Abbrechen' drücken",vbOKCancel,nameMsgBoxStacked)
								
		If boxRetVal = vbCancel Then
		
			WScript.Quit
			
		End If
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctLogfile

Public tempDir						'Temporärer Pfadmerker
Public tempFile						'Temporärer Dateimerker
Public createNewDir					'Zu erstellendes Verzeichnis (Objekt)
Public objLogfile					'Text File Object
Public nameLogfile					'Pfad der Logdatei
Public lineLog						'Zu schreibende Zeile
Public nameMsgBox					'Titel der Messagebox
Public countLogs					'Logzähler

'+	Logdatei schreiben

	'CALL:
	'lineLog			'Zu schreibende Zeile
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'thisPath			'Pfad zum aktuell ausgeführten Script
	'nameLogfile		'Pfad der Logdatei
	'sysMessaging		'Systemnachrichten: True = EIN

	'fctLogfile		lineLog , _
	'				nameMsgBox , _
	'				thisPath , _
	'				nameLogfile , _
	'				sysMessaging

	'RETURN:VOID

Function fctLogfile(ByVal lineLog,ByVal nameMsgBox,ByVal thisPath,ByVal nameLogfile,ByVal sysMessaging)

	nameMsgBoxStacked		= nameMsgBox
	
	If sysLoggingOn = True Then
		
		'Dateisystemobjekt definieren
		Set objFSO	= Wscript.CreateObject("Scripting.FileSystemObject")
		
		countLogs	=	countLogs + 1
		
'+	Logdateiverzeichnis festlegen

		tempDir		=	thisPath & "\logs\"
		
		If objFSO.FolderExists(tempDir) = False Then
		
			nameMsgBox	=	"fctLogfile"
			Set createNewDir = objFSO.CreateFolder(tempDir)
			
			'Script-Benachrichtigungen
			If sysMessaging	= True Then
			
				boxRetVal	= MsgBox(	"Verzeichnis anlegen:" & vbCrLf & tempDir & vbCrLf & vbCrLf & _
										"Zum beenden 'Abbrechen' drücken",vbOKCancel,nameMsgBoxStacked)
									
				If boxRetVal = vbCancel Then
				
					WScript.Quit
					
				End If
			End If
		End If

'+	Logdatei vorbereiten

		tempFile	= tempDir & nameLogfile

		If objFSO.FileExists(tempFile) = False Then
		
			'Logdatei erstellen/überschreiben (2)
			Set objLogfile	=	objFSO.OpenTextFile(tempFile, 2, True)
			
			'Script-Benachrichtigungen
			If sysMessaging	= True Then
			
				boxRetVal	= MsgBox(	"Logdatei anlegen:" & vbCrLf & tempFile & vbCrLf & vbCrLf & _
										"Zum beenden 'Abbrechen' drücken",vbOKCancel,nameMsgBoxStacked)
									
				If boxRetVal = vbCancel Then
							
					WScript.quit

				End If
			End If
			
		Else
			
			'Zeile anfügen (8)
			Set objLogfile	=	objFSO.OpenTextFile(tempFile, 8, True)
			
		End If

'+	Log formatieren

		lineLog	=	"+	Log[#" & countLogs & "] " & nameMsgBox & vbCrLf & vbCrLf & lineLog & vbCrLf
		
'+	Log schreiben

		objLogfile.WriteLine lineLog
		objLogfile.Close
		
		'Script-Benachrichtigungen
		If sysMessaging	= True Then
		
			boxRetVal	=	MsgBox(lineLog & vbCrLf & vbCrLf & _
							"'OK' zum fortfahren oder 'Abbrechen' zum beenden drücken...",vbOKCancel,nameMsgBox & " #sysMessaging")
								
'+	Programm fortfahren

			If boxRetVal = vbOK Then
			
				'Log vordefinieren
				lineLog	=	"Programm fortfahren." & vbCrLf & _
							"Nutzereingabe: " & boxRetVal
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBox , _
								thisPath , _
								nameLogfile , _
								False
				
'+	Programm beenden

			ElseIf boxRetVal = vbCancel Then
			
				'Log vordefinieren
				lineLog	=	"Programm beenden." & vbCrLf & _
							"Nutzereingabe: " & boxRetVal
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBox , _
								thisPath , _
								nameLogfile , _
								False
				
				fctSysEnd		nameMsgBoxStacked & " (sysMessaging)" , _
								False , _
								sysDebug
				
			End If
		
		End If
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctSysEnd

Public sysEndTriggered				'Verhindert unendliche Schleife bei sysMessaging = True in fctLogfile
Public sysQuiet						'Verhindert Systemmeldungen, die rein informativ sind

'+	Script beenden

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctSysEnd	nameMsgBox , _
	'			sysMessaging , _
	'			sysDebug


	'RETURN:VOID
	
Function fctSysEnd(ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked		= "fctSysEnd (" & nameMsgBox & ")"

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

'+	Excelinstanz beenden
		
		If IsObject(objExcel) = True Then
		
			objExcel.Application.Quit
			Set objExcel	= Nothing
			
		End If
		
'+	Wordinstanz beenden
		
		If IsObject(objWorddoc) = True Then
		
			objWorddoc.Application.Quit
			Set objWorddoc	= Nothing
			
		End If
		
		WScript.Quit

	End If
	
	sysEndTriggered	= True
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctChkDir

Public dirInputPath				'Auszuwertender Ordner (Pfad)
Public dirCreate					'True = Verzeichnis wird angelegt, wenn es noch nicht existiert
Public dirExist					'Antwort: True = Verzeichnis ex.; False = Verzeichnis ex. nicht

'+	Auf Verzeichnis prüfen

	'CALL:
	'dirInputPath		'Auszuwertender Ordner (Pfad)
	'dirCreate			'True = Verzeichnis wird angelegt, wenn es noch nicht existiert
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctChkDir	dirInputPath , _
	'			dirCreate , _
	'			nameMsgBox , _
	'			sysMessaging , _
	'			sysDebug

	'RETURN:dirExist			'Rückgabewert: True = Verzeichnis existiert bereits; False = Verzeichnis ex. nicht und wurde nicht angelegt

Function fctChkDir(ByVal dirInputPath,ByVal dirCreate,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	nameMsgBoxStacked		= "fctChkDir (" & nameMsgBox & ")"
	
	If objFSO.FolderExists(dirInputPath) = False And dirCreate = True Then
	
		Set createNewDir = objFSO.CreateFolder(dirInputPath)
		
		dirExist	=	True
		
'+	Verzeichnsi angelegt
		
		If objFSO.FolderExists(dirInputPath) = True Then
			
			'Log vordefinieren
			lineLog	=	"Verzeichnis angelegt:" & vbCrLf & dirInputPath & vbCrLf & "dirExist = " & dirExist & vbCrLf & "dirCreate = " & dirCreate
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
'+	Verzeichnsi konnte nicht angelegt werden
							
		Else
		
			'Log vordefinieren
			lineLog	=	"Fehler! Verzeichnis nicht angelegt:" & vbCrLf & dirInputPath & vbCrLf & "dirExist = " & dirExist & vbCrLf & "dirCreate = " & dirCreate
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
	ElseIf objFSO.FolderExists(dirInputPath) = False And dirCreate = False Then
	
		dirExist	=	False
		
		'Log vordefinieren
		lineLog	=	"Verzeichnis nicht gefunden:" & vbCrLf & dirInputPath & vbCrLf & "dirExist = " & dirExist & vbCrLf & "dirCreate = " & dirCreate
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	ElseIf objFSO.FolderExists(dirInputPath) = True Then
	
		dirExist	=	True
		
		'Log vordefinieren
		lineLog	=	"Verzeichnis gefunden:" & vbCrLf & dirInputPath & vbCrLf & "dirExist = " & dirExist & vbCrLf & "dirCreate = " & dirCreate
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctChkFile			-> Prüft auf Existenz einer Datei

Public filePath				'Pfad und Name der Datei
Public fileExist			'Antwort: True = Datei ex.; False = Datei ex. nicht
Public fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert

'+	Prüfe auf Existenz einer Datei

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctChkFile	filePath , _
	'			fileExistWarning , _
	'			nameMsgBox , _
	'			sysMessaging , _
	'			sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

Function fctChkFile(ByVal filePath,ByVal fileExistWarning,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	nameMsgBoxStacked		= "fctChkFile (" & nameMsgBox & ")"

	If sysDebug = True Then
		
		'Log vordefinieren
		lineLog	=	"Prüfe auf Existenz einer Datei."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
					
	If objFSO.FileExists(filePath) = True Then
		
		fileExist	= True
		
		If sysDebug = True Then
			
			'Log vordefinieren
			lineLog	=	"Datei wurde gefunden." & vbCrLf & _
						"filePath: " & filePath & vbCrLf & _
						"fileExist: " & fileExist
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
	Else
		
		fileExist	= False
		
'+	Meldungen Generieren
		
		If fileExistWarning = False Then
			
			'Log vordefinieren
			lineLog	=	"Datei wurde nicht gefunden." & vbCrLf & _
						"filePath: " & filePath & vbCrLf & _
						"fileExist: " & fileExist
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
		Else
			
			'Log vordefinieren
			lineLog	=	"Datei wurde nicht gefunden!" & vbCrLf & _
						"filePath: " & filePath & vbCrLf & _
						"fileExist: " & fileExist & vbCrLf & vbCrLf & _
						"'OK' zum fortfahren oder 'Abbrechen' zum beenden drücken..."
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
'+	Nutzerabfrage

			boxRetVal = MsgBox(	lineLog,52,sysProgramName & " - Datei nicht gefunden!")
						
'+	Programm fortfahren

			If boxRetVal = vbOK Then
			
				'Log vordefinieren
				lineLog	=	"Programm fortfahren." & vbCrLf & _
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
				lineLog	=	"Programm beenden." & vbCrLf & _
							"Nutzereingabe: " & boxRetVal
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
				
				fctSysEnd		nameMsgBoxStacked , _
								sysMessaging , _
								sysDebug
					
			End If
			
		End If
		
	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctCopyFile

Public sourceFile		'Pfad und Dateiname der zu kopierenden Datei
Public destFile			'Pfad und Dateiname der zu erstellenden Datei

'+	Datei kopieren

	'CALL:
	'sourceFile			'Pfad und Dateiname der zu kopierenden Datei
	'destFile			'Pfad und Dateiname der zu erstellenden Datei
	'Overwrite			'True = Überschreibe Zieldatei, wenn sie bereits existiert
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctCopyFile	sourceFile , _
	'				destFile , _
	'				Overwrite , _
	'				nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:VOID

Function fctCopyFile(ByVal sourceFile,ByVal destFile,ByVal Overwrite,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	On Error Resume Next

	nameMsgBoxStacked		= "fctCopyFile (" & nameMsgBox & ")"

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	sourceFile , _
				True , _
				nameMsgBox , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then
	
		'Log vordefinieren
		lineLog	=	"Datei wird kopiert." & vbCrLf & _
					"sourceFile: " & sourceFile & vbCrLf & _
					"destFile: " & destFile
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
								
		'CALL:
		'filePath				'Pfad und Name der Datei
		'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
		'nameMsgBox				'Name des Systemnachrichtenfensters
		'sysMessaging			'Systemnachrichten: True = EIN
		'sysDebug				'Exzessives Logging

		fctChkFile	destFile , _
					False , _
					nameMsgBox , _
					sysMessaging , _
					sysDebug

		'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

		If fileExist = False Then
		
			objFSO.CopyFile sourceFile, destFile, Overwrite
			
			'Log vordefinieren
			lineLog	=	"Datei kopiert." & vbCrLf & _
						"sourceFile: " & sourceFile & vbCrLf & _
						"destFile: " & destFile
						
		ElseIf fileExist = True AND Overwrite = False Then
		
			'Log vordefinieren
			lineLog	=	"Datei existiert bereits und wurde NICHT überschrieben." & vbCrLf & _
						"sourceFile: " & sourceFile & vbCrLf & _
						"destFile: " & destFile
						
		ElseIf fileExist = True AND Overwrite = True Then
		
			objFSO.CopyFile sourceFile, destFile, Overwrite
			
			'Log vordefinieren
			lineLog	=	"Datei existiert und wurde überschrieben." & vbCrLf & _
						"sourceFile: " & sourceFile & vbCrLf & _
						"destFile: " & destFile
						
		End If
	
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
		
	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctDeleteFile

Public	deleteFile			'Pfad zur Datei, die gelöscht werden soll

'+	Löschen einer Datei

	'CALL:
	'deleteFile			'Pfad zur Datei, die gelöscht werden soll
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctDeleteFile	deleteFile , _
	'				nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:VOID

Function fctDeleteFile(ByVal deleteFile,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	On Error Resume Next

	nameMsgBoxStacked		= "fctDeleteFile (" & nameMsgBox & ")"

'+	Datei löschen

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	deleteFile , _
				False , _
				nameMsgBox , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then
	
		'Log vordefinieren
		lineLog	=	"Datei wird gelöscht." & vbCrLf & _
					"deleteFile: " & deleteFile
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
							
		objFSO.GetFile(deleteFile).delete
		
	Else
	
		'Log vordefinieren
		lineLog	=	"Datei wurde nicht gefunden. Löschvorgang abgebrochen." & vbCrLf & _
					"deleteFile: " & deleteFile
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
							
	End If
	
'+	Überprüfe, ob Aktion erfolgreich war

	'CALL:
	'filePath				'Pfad und Name der Datei
	'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	fctChkFile	deleteFile , _
				False , _
				nameMsgBox , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then
	
		'Log vordefinieren
		lineLog	=	"Datei wurde NICHT gelöscht! Überprüfen Sie Parameter und Berechtigungen." & vbCrLf & _
					"deleteFile: " & deleteFile
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
		boxRetVal	=	MsgBox(lineLog,vbExclamation,nameMsgBoxStacked)
							
	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctSelFolderDialog

Public dirSelRequest		'True = Zeige Nutzerabfrage mit Ordnerpfad
Public	dirSelected			'Verzeichnispfad der Auswahl

'+	Verzeichnis wählen

	'CALL:
	'dirSelRequest		'True = Zeige Nutzerabfrage mit Ordnerpfad
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctSelFolderDialog	dirSelRequest , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug

	'RETURN:dirSelected			'Verzeichnispfad der Auswahl

Function fctSelFolderDialog(ByVal dirSelRequest,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	On Error Resume Next

	nameMsgBoxStacked		= "fctSelFolderDialog (" & nameMsgBox & ")"
	loopActive	=	True

'+	Schleife der Ordnerauswahl starten

	Do While loopActive = True
	
		Set objShell = CreateObject("Shell.Application").BrowseForFolder(0,sysProgramName & " Verzeichnisauswahl",1,17) 
		
		Set dirSelected = objShell.Self
		
		'Log vordefinieren
		lineLog	=	"Pfad ausgewählt:" & vbCrLf & dirSelected
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
'+	Überprüfung durch Benutzer (optional)

		If dirSelRequest = True And objFSO.FolderExists(dirSelected) = True Then
		
			'Log vordefinieren
			lineLog	=	"Ist die Pfadangabe korrekt?" & vbCrLf & vbCrLf & dirSelected
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
				
			boxRetVal	=	MsgBox (lineLog,vbYesNo,sysProgramName & " Verzeichnisauswahl")
			
			If boxRetVal = vbYes Then
			
				loopActive	=	False
				'Log vordefinieren
				lineLog		=	"Verzeichnis übernommen:" & vbCrLf & dirSelected
				
			Else
			
				'Log vordefinieren
				lineLog		=	"Verzeichnis verworfen:" & vbCrLf & dirSelected & vbCrLf & vbCrLf & _
								"Erneute Auswahl..."
				
			End If
			
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		ElseIf dirSelRequest = False And objFSO.FolderExists(dirSelected) = True Then
		
			loopActive	=	False
		
'+	Ausgewählter Ordner ist kein gültiges Verzeichnis

		ElseIf objFSO.FolderExists(dirSelected) = False Then
		
			loopActive	=	True
			
			'Log vordefinieren
			lineLog	=	"Der Pfad existiert nicht:" & vbCrLf & dirSelected & vbCrLf & vbCrLf & _
						"Bitte wählen Sie ein gültiges Verzeichnis."
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
				
			boxRetVal	=	MsgBox (lineLog,vbOKCancel,sysProgramName & " Verzeichnisauswahl")
			
		End If
		
'+	Programm Abbrechen durch Benutzer

		If boxRetVal = vbCancel Then
		
			'Log vordefinieren
			lineLog	=	"Programm beendet durch Nutzereingabe"
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
						
			fctSysEnd	nameMsgBox , _
						sysMessaging , _
						sysDebug

		End If
			
		Set boxRetVal = Nothing
		
	Loop
	
	Set objShell = Nothing
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctInputBoxTest

Public inputText					'Text des Eingabefensters
Public inputBoxCaption				'Überschrift des Eingabefensters
Public inputAutoFill				'Angezeigter Textvorschlag
Public inputAutoCorrectTrue		'True = Autokorrektur zulassen (Sonderzeichen entfernen)
Public inputResult					'Ergebnis der Eingabe
Public tempSpecialChar				'Merker für ein Sonderzeichen
Public tempMsgActive				'Mehr für Aktivierung einer Messagebox (BOOL)

'+	Nutzereingabe

	'CALL:
	'inputText					'Text im Eingabefenster
	'inputBoxCaption			'Überschrift des Eingabefensters
	'inputAutoFill				'Angezeigter Textvorschlag
	'inputAutoCorrectTrue		'Autokorrektur zulassen (True/False)
	'nameMsgBox					'Name des Systemnachrichtenfensters
	'sysMessaging				'Systemnachrichten: True = EIN
	'sysDebug					'Exzessives Logging

	'fctInputBoxTest	inputText , _
	'					inputBoxCaption	, _
	'					inputAutoFill , _
	'					inputAutoCorrectTrue , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug

	'RETURN:inputResult			'Ergebnis der Eingabe/Autokorrektur

Function fctInputBoxTest (ByVal inputText,ByVal inputBoxCaption,ByVal inputAutoFill,ByVal inputAutoCorrectTrue,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)
	
	On Error Resume Next

	nameMsgBoxStacked		= "fctInputBoxTest (" & nameMsgBox & ")"
	loopActive	= True
	
	Do While loopActive = True
		
		loopActive	= False

		'Benutzerabfrage
		inputResult	=	InputBox(inputText,inputBoxCaption,inputAutoFill)
		
		'Log vordefinieren
		lineLog	=	"Nutzereingabe:" & vbCrLf & inputResult
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
							
		If inputAutoCorrectTrue = True Then
	
			'Sonderzeichen entfernen
			tempString = ""
			tempCounter = 1
			
			For tempCounter = 1 To Len(inputResult)
			
				Select Case Mid(inputResult, tempCounter, 1)
				
					Case "$"
						tempSpecialChar	=	"$"
						tempMsgActive	=	True
					Case "\"
						tempSpecialChar	=	"\"
						tempMsgActive	=	True
					Case """"
						tempSpecialChar	=	""""
						tempMsgActive	=	True
					Case "/"
						tempSpecialChar	=	"/"
						tempMsgActive	=	True
					Case ":"
						tempSpecialChar	=	":"
						tempMsgActive	=	True
					Case "*"
						tempSpecialChar	=	"*"
						tempMsgActive	=	True
					Case "?"
						tempSpecialChar	=	"?"
						tempMsgActive	=	True
					Case "<"
						tempSpecialChar	=	"<"
						tempMsgActive	=	True
					Case ">"
						tempSpecialChar	=	">"
						tempMsgActive	=	True
					Case "|"
						tempSpecialChar	=	"|"
						tempMsgActive	=	True
					Case ","
						tempSpecialChar	=	","
						tempMsgActive	=	True
					Case ";"
						tempSpecialChar	=	";"
						tempMsgActive	=	True
					Case "+"
						tempSpecialChar	=	"+"
						tempMsgActive	=	True
					Case "*"
						tempSpecialChar	=	"*"
						tempMsgActive	=	True
					Case "~"
						tempSpecialChar	=	"~"
						tempMsgActive	=	True
					Case "'"
						tempSpecialChar	=	"'"
						tempMsgActive	=	True
					Case "#"
						tempSpecialChar	=	"#"
						tempMsgActive	=	True
					Case "!"
						tempSpecialChar	=	"!"
						tempMsgActive	=	True
					Case "§"
						tempSpecialChar	=	"§"
						tempMsgActive	=	True
					Case "."
						tempSpecialChar	=	"."
						tempMsgActive	=	True
					Case "´"
						tempSpecialChar	=	"´"
						tempMsgActive	=	True
					Case "`"
						tempSpecialChar	=	"`"
						tempMsgActive	=	True
					Case "{"
						tempSpecialChar	=	"{"
						tempMsgActive	=	True
					Case "}"
						tempSpecialChar	=	"}"
						tempMsgActive	=	True
					Case "%"
						tempSpecialChar	=	"%"
						tempMsgActive	=	True
					Case Else
						tempString = tempString + Mid(inputResult, tempCounter, 1)
						tempMsgActive	=	False
					
				End Select
				
				'Sonderzeichen gefunden
				If tempMsgActive	=	True Then
						
					'Log vordefinieren
					lineLog	=	"Sonderzeichen gefunden: " & tempSpecialChar & vbCrLf & _
								"Zeichenfolge: " & inputResult & vbCrLf & _
								"Position: " & tempCounter
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
						
				End If
				
				tempSpecialChar	= ""
				
			Next
			
			'Nutzerabfrage Autokorrektur
			If tempString <> inputResult Then
			
				'Log vordefinieren
				lineLog	=	"Sonderzeichen können Probleme verursachen." & vbCrLf & _
							"Soll die Eingabe korrigiert werden?" & vbCrLf & vbCrLf & _
							"Vorher: " & inputResult & vbCrLf & _
							"Nachher: " & tempString & vbCrLf & vbCrLf & _
							"Ja ... Autokorrektur annehmen " & vbCrLf & _
							"Nein ... Autokorrektur ablehnen " & vbCrLf & _
							"Abbruch ... Programm beenden"
				'Logdatei schreiben
				fctLogfile		lineLog , _
								nameMsgBoxStacked , _
								thisPath , _
								nameLogfile , _
								sysMessaging
					
				boxRetVal = MsgBox(	lineLog,vbYesNoCancel,sysProgramName & " - Autokorrektur")
								
				'Autokorrektur angenommen
				If boxRetVal = vbYes Then
				
					'Log vordefinieren
					lineLog	=	"Autokorrektur: " & vbCrLf & tempString & vbCrLf & vbCrLf & _
								"Originaltext: " & vbCrLf & inputResult
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
						
					inputResult = tempString
					loopActive = False
								
				'Autokorrektur abgelehnt
				ElseIf boxRetVal = vbNo Then
				
					'Log vordefinieren
					lineLog	=	"Autokorrektur abgelehnt: " & vbCrLf & tempString & vbCrLf & vbCrLf & _
								"Originaltext: " & vbCrLf & inputResult
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
						
					loopActive = False
								
				'Programm beenden
				ElseIf boxRetVal = vbCancel Then
				
					'Log vordefinieren
					lineLog	=	"Programm beendet"
					'Logdatei schreiben
					fctLogfile		lineLog , _
									nameMsgBoxStacked , _
									thisPath , _
									nameLogfile , _
									sysMessaging
						
					loopActive = False
									
					fctSysEnd	nameMsgBox , _
								sysMessaging , _
								sysDebug

				End If
			Else
				
				loopActive = False
				
			End If
		End If
		
		If IsEmpty(inputResult) Or inputResult = "" Then

			boxRetVal	= MsgBox ("Ungültige Eingabe!" & vbCrLf & "Wiederholen...",vbOK,nameMsgBox)
			
			'Log vordefinieren
			lineLog	=	"Ungültige Eingabe!" & vbCrLf & "Wiederholen..."
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
				
			loopActive = True
				
		End If
	Loop
	
	'Leerzeichen an Anfang und Ende entfernen
	inputResult = Trim(inputResult)
	Set loopActive	=	Nothing

	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctReplaceString

'Public VARIABLE

'+	Zeichenkette in einer Textdatei ändern

	'CALL:
	'sourceFile				'Zu durchsuchende Zeichenkette
	'inputReplace			'Diese Zeichenkette wird ausgetauscht
	'inputWith				'Diese Zeichenkette ersetzt die alte
	'nameMsgBox				'Name des Systemnachrichtenfensters
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctReplaceString	sourceFile , _
	'					inputReplace , _
	'					inputWith , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	
	'RETURN:VOID

Function fctReplaceString (ByVal sourceFile,ByVal inputReplace,ByVal inputWith,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctReplaceString (" & nameMsgBox & ")"

'+	Textdatei lesen

	tempString			= objFSO.OpenTextFile(sourceFile, 1).ReadAll

	'Log vordefinieren
	lineLog	=	"Zeichenfolge wird gesucht und ersetzt." & vbCrLf & _
				"sourceFile: " & sourceFile & vbCrLf & _
				"inputReplace: " & inputReplace & vbCrLf & _
				"inputWith: " & inputWith
				
	'Exzessives Logging
	If sysDebug = True Then
	
		lineLog	=	lineLog & vbCrLf & vbCrLf & _
					"sysDebug: " & sysDebug & vbCrLf & vbCrLf & _
					tempString
						
	End If
				
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
			
'+	Zeichenkette suchen und ersetzen

	outputString	= Replace(tempString, inputReplace, inputWith)
			
'+	Ausgabetext schreiben
	
	If tempString <> outputString Then
	
	fctCmdWrite		sourceFile , _
					outputString , _
					True , _
					nameMsgBoxStacked , _
					sysMessaging , _
					sysDebug

	'Log vordefinieren
	lineLog	=	"Zeichenfolge wurde gefunden und ersetzt." & vbCrLf & _
				"sourceFile: " & sourceFile & vbCrLf & _
				"inputReplace: " & inputReplace & vbCrLf & _
				"inputWith: " & inputWith
					
		'Exzessives Logging
		If sysDebug = True Then
		
			lineLog	=	lineLog & vbCrLf & vbCrLf & _
						"sysDebug: " & sysDebug & vbCrLf & vbCrLf & _
						"Neuer Dateiinhalt:" & vbCrLf & tempString
							
		End If
				
	Else
	
		'Log vordefinieren
		lineLog	=	"Zeichenfolge zum ersetzen wurde nicht gefunden." & vbCrLf & _
					"inputReplace: " & inputReplace
	
	End If
	
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
			
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctReplaceCharacter

Public inputString				'Zu durchsuchender String
Public inputReplace				'Dieses Zeichen wird ausgetauscht
Public inputWith				'Dieses Zeichen ersetzt das alte
Public outputString				'Geänderte Zeichenkette

'+	Zeichen eines Strings korrigieren/ändern

	'CALL:
	'inputString					'Zu durchsuchender String
	'inputReplace					'Dieses Zeichen wird ausgetauscht
	'inputWith						'Dieses Zeichen ersetzt das alte
	'nameMsgBox						'Name des Systemnachrichtenfensters
	'sysMessaging					'Systemnachrichten: True = EIN
	'sysDebug						'Exzessives Logging

	'inputString	= fctReplaceCharacter(	inputString , _
	'										inputReplace , inputWith , _
	'										nameMsgBox , sysMessaging , sysDebug)

Function fctReplaceCharacter (ByVal inputString,ByVal inputReplace,ByVal inputWith,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctReplaceCharacter (" & nameMsgBox & ")"
	
'+	Schleifen-Startparameter

	tempString = ""
	tempCounter = 1
	
'+	Schleife für Symbol suchen und ersetzen

	For tempCounter = 1 To Len(inputString)
	
		Select Case Mid(inputString, tempCounter, 1)
		
			Case inputReplace
				tempSpecialChar	=	inputReplace
				tempString = tempString & inputWith
			Case Else
				tempString = tempString + Mid(inputString, tempCounter, 1)
				tempSpecialChar	= ""
			
		End Select
		
		If sysMessaging = True And tempSpecialChar <> "" Then
			
			'Log vordefinieren
			lineLog	=	"Zeichen gefunden: " & tempSpecialChar & vbCrLf & _
						"Zeichenfolge: " & inputString & vbCrLf & _
						"Position: " & tempCounter
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
				
	Next
	
'+	Ausgabetext schreiben
	
	fctReplaceCharacter	= tempString
	outputString		= tempString

	If inputString <> outputString Then
	
		'Log vordefinieren
		lineLog	=	"Zeichenfolge korrigiert: " & vbCrLf & _
					"Original: " & inputString & vbCrLf & _
					"Korrektur: " & inputReplace & " mit " & inputWith & vbCrLf & _
					"Output: " & outputString
					
	Else
	
		'Log vordefinieren
		lineLog	=	"Zeichenfolge nicht korrigiert: " & vbCrLf & _
					"Original: " & outputString
	
	End If
	
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
			
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctRemoveSpecial

Public loopActiveVal					'True = Die aktuelle Eingabe wird wiederholt

'+	Sonderzeichen entfernen

	'CALL:
	'inputString					'Zu durchsuchender String
	'nameMsgBox						'Name des Systemnachrichtenfensters
	'sysMessaging					'Systemnachrichten: True = EIN
	'sysDebug						'Exzessives Logging

	'inputString	= fctRemoveSpecial(	inputString , _
	'									nameMsgBox , sysMessaging , sysDebug)
	
	'RETURN:loopActiveVal			'True = Die aktuelle Eingabe wird wiederholt
	'		outputString			'Geänderte Zeichenkette

Function fctRemoveSpecial (ByVal inputString,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctRemoveSpecial (" & nameMsgBox & ")"

	'Log vordefinieren
	lineLog	=	"Sonderzeichen werden gesucht. " & tempSpecialChar & vbCrLf & _
				"inputString: " & inputString
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
					
'+	Sonderzeichen entfernen

	tempString = ""
	i = 1
	
	For i = 1 To Len(inputString)
	
		Select Case Mid(inputString, i, 1)
		
			Case "^"
			Case "°"
			Case "!"
			Case """"
			Case "§"
			Case "$"
			Case "%"
			Case "/"
			Case "="
			Case "?"
			Case "\"
			Case "´"
			Case "`"
			Case "*"
			Case "~"
			Case "'"
			Case "#"
			Case "<"
			Case ">"
			Case "|"
			Case ";"
			Case ","
			Case ":"
			Case "."
			Case "("
			Case ")"
			Case Else tempString = tempString + Mid(inputString, i, 1)
			
		End Select	
		
	Next
	
	inputString = Trim(inputString)

'+	Nutzerabfrage Autokorrektur

	If tempString <> inputString AND sysQuiet = False Then
		
		'Ereignis schreiben
		lineLog	= 	"Es wurden Sonderzeichen gefunden." & vbCrLf & _
					"inputString: " & inputString
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
			
		BoxRetValue = MsgBox(	"Sonderzeichen können Probleme verursachen." & vbCrLf & _
								"Soll die Zeichenkette angepasst werden?" & vbCrLf & vbCrLf & _
								"Vorher: " & inputString & vbCrLf & _
								"Nachher: " & tempString,vbYesNo,"Autokorrektur")
		
		If BoxRetValue = vbYes Then
		
			fctRemoveSpecial	= tempString
			outputString		= tempString
		
			'Ereignis schreiben
			lineLog	= 	"Die Korrektur wurde übernommen." & vbCrLf & _
						"inputString: " & inputString & vbCrLf & _
						"outputString: " & outputString
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
			loopActiveVal = False
			
		ElseIf BoxRetValue = vbNo Then

			fctRemoveSpecial	= inputString
			outputString		= inputString
		
			'Ereignis schreiben
			lineLog	= 	"Die Korrektur wurde nicht übernommen." & vbCrLf & _
						"inputString: " & inputString & vbCrLf & _
						"outputString: " & outputString
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
			loopActiveVal = False
			
		End If
	ElseIf tempString <> inputString AND sysQuiet = True Then
		
			outputString		= tempString
			
	Else
		
		fctRemoveSpecial	= inputString
		outputString		= inputString
	
		'Ereignis schreiben
		lineLog	= 	"Es wurden keine Sonderzeichen gefunden." & vbCrLf & _
					"inputString: " & inputString & vbCrLf & _
					"tempString: " & tempString
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
			
		loopActiveVal = False

	End If
	
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctRegEx

Public regExGlobal			'False = Stoppt Suche bei erster Übereinstimmung
Public regExPattern		'Muster zur Überprüfung (Beachte Vorschrift)
Public regExTrue			'Muster zutreffend
Public countRegEx			'Anzahl gefundener Übereinstimmungen

'+	Eine Zeichenfolge auf Muster untersuchen

	'CALL:
	'regExGlobal		'False = Stoppt Suche bei erster Übereinstimmung
	'regExPattern		'Muster zur Überprüfung (Beachte Vorschrift)
	'inputString		'Zeichenkette, die untersucht wird
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctRegEx		regExGlobal , _
	'				regExPattern , _
	'				inputString , _
	'				nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:regExTrue			'True = Muster zutreffend
	'		countRegEx			'Anzahl gefundener Übereinstimmungen
	
Function fctRegEx (ByVal regExGlobal,ByVal regExPattern,ByVal inputString,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked		= "fctRegEx (" & nameMsgBox & ")"
	
'+	Variablenübergabe

	objRegEx.Global		= regExGlobal
	objRegEx.Pattern	= regExPattern
	inputString			= Trim(inputString)
	Set regExpMatches	= objRegEx.Execute(inputString)		'Übereinstimmung String mit Muster prüfen


'+	Operation ankündigen
	
	'Log vordefinieren
	lineLog	=	"Untersuche Zeichenkette auf Muster:" & vbCrLf & _
				"regExGlobal: " & regExGlobal & vbCrLf & _
				"regExPattern: " & regExPattern & vbCrLf & _
				"inputString: " & inputString
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Übereinstimmungen finden

	countRegEx			= regExpMatches.Count
	
	If countRegEx > 0 Then
	
		regExTrue	= True
		lineLog			=	"Übereinstimmungen gefunden!" & vbCrLf & vbCrLf & _
							"countRegEx: " & countRegEx & vbCrLf & _
							"regExTrue: " & regExTrue
		
	Else
	
		regExTrue	= False
		lineLog			=	"Keine Übereinstimmungen gefunden!" & vbCrLf & vbCrLf & _
							"countRegEx: " & countRegEx & vbCrLf & _
							"regExTrue: " & regExTrue
		
	End If
		
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctCmdWrite

Public	objCmd				'Objekt der Batch-/Textdatei
Public cmdPath				'Pfad und Name der zu schreibenden Datei
Public cmdCode				'Inhalt der Batch-/Textdatei

'+	Batch-/Textdatei anlegen und schreiben

	'CALL:
	'cmdPath			'Pfad und Name der zu schreibenden Datei
	'cmdCode			'Inhalt der Batch-/Textdatei
	'Overwrite			'True = Überschreibe Zieldatei, wenn sie bereits existiert
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctCmdWrite	cmdPath , _
	'				cmdCode , _
	'				Overwrite , _
	'				nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:VOID

Function fctCmdWrite(ByVal cmdPath,ByVal cmdCode,ByVal Overwrite,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked		= "fctCmdWrite (" & nameMsgBox & ")"

'+	Dateimodus festlegen

	Const ForWriting	= 2
	Const Create		= True

'+	Operation ankündigen	

	'Log vordefinieren
	lineLog	=	"Schreibe Batch-/Textdatei:" & vbCrLf & _
				"Overwrite = " & Overwrite & vbCrLf & _
				"Pfad: " & cmdPath
				
	'Exzessives Logging
	If sysDebug = True Then
	
		lineLog	=	lineLog & vbCrLf & vbCrLf & _
					"sysDebug: " & sysDebug & vbCrLf & vbCrLf & _
					cmdCode
						
	End If
				
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Datei schreiben/überschreiben/übersprinen	
				
	If objFSO.FileExists(cmdPath) = False Then
		
		Set objCmd	= objFSO.OpenTextFile(cmdPath, ForWriting, Create)
		
		objCmd.WriteLine cmdCode
		objCmd.close
		
		'Log vordefinieren
		lineLog	=	"Datei geschrieben:" & vbCrLf & _
					"Pfad: " & cmdPath
					
	ElseIf objFSO.FileExists(cmdPath) = True And Overwrite = False Then
		
		'Log vordefinieren
		lineLog	=	"Datei existiert bereits und wurde nicht überschrieben." & vbCrLf & _
					"Pfad: " & cmdPath
					
	ElseIf objFSO.FileExists(cmdPath) = True And Overwrite = True Then
	
		Set objCmd			= objFSO.OpenTextFile(cmdPath, ForWriting, Create)
		
		objCmd.WriteLine cmdCode
		objCmd.close
		
		'Log vordefinieren
		lineLog	=	"Datei wurde überschrieben." & vbCrLf & _
					"Pfad: " & cmdPath
					
	End If

	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
					
	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctCreateFolder

'Public	PLATZHALTERVARIABLE

'+	Verzeichnis erstellen

	'CALL:
	'tempDir			'Temporärer Pfadmerker
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctCreateFolder	tempDir , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug

	'RETURN:VOID

Function fctCreateFolder(ByVal tempDir,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctCreateFolder (" & nameMsgBox & ")"

	If objFSO.FolderExists(tempDir) = False Then
	
		'Log vordefinieren
		lineLog	=	"Verzeichnis wird angelegt." & vbCrLf & _
					"tempDir: " & tempDir & vbCrLf & _
					"createNewDir: " & createNewDir
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
		Set createNewDir	= objFSO.CreateFolder(tempDir)

	Else
	
		'Log vordefinieren
		lineLog	=	"Abbruch, Verzeichnis existiert bereits." & vbCrLf & _
					"tempDir: " & tempDir & vbCrLf & _
					"createNewDir: " & createNewDir
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If

	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctTestFolder

'Public	PLATZHALTERVARIABLE

'+	fctTestFolder			Prüfe Ordner auf Existenz, lege Overwrite fest und erstelle Verzeichnis bedingt

	'CALL:
	'tempDir						'Temporärer Pfadmerker
	'nameMsgBox						'Name des Systemnachrichtenfensters
	'sysMessaging					'Systemnachrichten: True = EIN
	'sysDebug						'Exzessives Logging
	
	'fctTestFolder	tempDir , _
	'				nameMsgBox , _
	'				sysMessaging , _
	'				sysDebug

	'RETURN:Overwrite				'True = Überschreibe Zieldatei, wenn sie bereits existiert
	
Function fctTestFolder(ByVal tempDir,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctTestFolder (" & nameMsgBox & ")"

	'Ereignis schreiben
	lineLog	= 	"Prüfe Verzeichnis auf Existenz." & vbCrLf & _
				"tempDir: " & tempDir
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

'+	Verzeichnis erstellen

	If objFSO.FolderExists(tempDir) = False then
	
		fctCreateFolder		tempDir , _
							nameMsgBoxStacked , _
							sysMessaging , _
							sysDebug

		'Ereignis schreiben
		lineLog	= 	"Verzeichnis wurde neu angelegt." & vbCrLf & _
					"tempDir: " & tempDir
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

'+	Abfrage, ob Verzeichnis überschrieben werden soll

	Else
	
		'Ereignis schreiben
		lineLog	= 	"Verzeichnis existiert bereits." & vbCrLf & _
					"tempDir: " & tempDir & vbCrLf & vbCrLf & _
					"Ja...Verzeichnis integrieren und Dateien überschreiben" & vbCrLf & _
					"Nein...Verzeichnis integrieren und Dateien nicht überschreiben"
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging

		'Abfrage wenn Verzeichnis bereits existiert
		BoxRetValue = MsgBox(lineLog,vbYesNoCancel Or vbExclamation, "Verzeichniskonflikt")
		
		If BoxRetValue = vbCancel Then
		
			'Ereignis schreiben
			lineLog	= 	"Abbruch durch den Benutzer." & vbCrLf & _
						"BoxRetValue: " & BoxRetValue
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging

			'Script beenden
			fctSysEnd		nameMsgBoxStacked , _
							sysMessaging , _
							sysDebug
			
		ElseIf BoxRetValue = vbYes Then
		
			Overwrite = True
			
			'Ereignis schreiben
			lineLog	= 	"Überschreiben vorhandener Dateien/Verzeichnisse." & vbCrLf & _
						"Overwrite: " & Overwrite & vbCrLf & _
						"tempDir: " & tempDir & vbCrLf & _
						"createNewDir: " & createNewDir
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging

'+	Verzeichnis überschreiben

			fctCreateFolder		tempDir , _
								nameMsgBoxStacked , _
								sysMessaging , _
								sysDebug

		ElseIf BoxRetValue = vbNo Then
		
			Overwrite = False
			
			'Ereignis schreiben
			lineLog	= 	"Integrieren von Dateien/Verzeichnissen." & vbCrLf & _
						"Overwrite: " & Overwrite & vbCrLf & _
						"tempDir: " & tempDir
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
							
		End If
		
	End If

	fctErrorHandling	nameMsgBoxStacked , sysMessaging , sysDebug

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctWeekNumber

Public weekNr			'Kalenderwochennummer
Public tempDate			'Datum im Format TT.MM.JJJJ

'+	Gibt die Kalenderwoche eines Datums zurück

	'CALL:
	'tempDate			'Datum im Format TT.MM.JJJJ
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctWeekNumber		tempDate , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:weekNr		'Kalenderwochennummer

Function fctWeekNumber(ByVal tempDate,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	nameMsgBoxStacked	= "fctWeekNumber (" & nameMsgBox & ")"
	
	If sysDebug = True Then

		'Ereignis schreiben
		lineLog	= 	"Ermittle Kalenderwochennummer." & vbCrLf & _
					"tempDate: " & tempDate
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
	
	weekNr	= DatePart("ww",tempDate)
	
	If sysDebug = True Then

		'Ereignis schreiben
		lineLog	= 	"Kalenderwochennummer ermittelt." & vbCrLf & _
					"tempDate: " & tempDate & vbCrLf & _
					"weekNr: " & weekNr
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctErrorHandling

Public errString		'Sammlung der Fehlermeldungen
Public errCnt			'Zählung der Fehlermeldungen

'+	Errorhandling

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:errString		'Sammlung der Fehlermeldungen

Function fctErrorHandling(byVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	nameMsgBoxStacked	= "fctErrorHandling (" & nameMsgBox & ")"
	
	If sysDebug = True Then

		'Ereignis schreiben
		lineLog	= 	"Überprüfe Fehlerstatus."
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If
	
'+	Vorformatierung der Fehlersammlung

'+	Errorhandling

	If Err.Number <> 0 Then
	
		errCnt		= errCnt + 1
		
		'Log vordefinieren
		lineLog	= 	"ERROR[" & errCnt & "] Der Befehl wurde nicht ausgeführt." & vbCrLf & vbCrLf & _
					"Err.Number: " & Err.Number  & vbCrLf & _
					"Err.Description: " & Err.Description & vbCrLf & _
					"Err.Source: " & Err.Source & vbCrLf & _
					"Modul: " & nameMsgBox & vbCrLf & vbCrLf & _
					"'OK' zum fortfahren oder 'Abbrechen' zum beenden drücken..."
		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
							
'+	Fehlersammlung generieren
							
		If errString = "" Or IsEmpty(errString) = True Then
		
			errString	= "=== OCCURED ERRORS ==="
			
		End If

		errString	= errString & vbCrLf & "ERROR[" & errCnt & "] " & nameMsgBox & ": " & "(" & Err.Number & ") " & Err.Description & " - " & Err.Source
		
		
'+	Nutzerabfrage

		boxRetVal = MsgBox(	lineLog,17,sysProgramName & " - Fehler!")
		
'+	Programm fortfahren

		If boxRetVal = vbOK Then
		
			'Log vordefinieren
			lineLog	=	"Programm fortfahren." & vbCrLf & _
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
			lineLog	=	"Programm beenden." & vbCrLf & _
						"Nutzereingabe: " & boxRetVal
			'Logdatei schreiben
			fctLogfile		lineLog , _
							nameMsgBoxStacked , _
							thisPath , _
							nameLogfile , _
							sysMessaging
			
'+	Excelinstanz beenden
			
			If IsObject(objExcel) = True Then
			
				objExcel.Application.Quit
				Set objExcel	= Nothing
				
			End If
			
'+	Wordinstanz beenden
					
			If IsObject(objWorddoc) = True Then
			
				objWorddoc.Application.Quit
				Set objWorddoc	= Nothing
				
			End If
					
			fctSysEnd		nameMsgBoxStacked , _
							sysMessaging , _
							sysDebug
				
		End If
		
'+	Fehlerspeicher leeren
		
		Err.Clear
		
	End If
	
	nameMsgBoxStacked	= nameMsgBox
		
End Function
