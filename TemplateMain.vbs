'----------------------------------------------------------------------------
' =========================== INITIALISIERUNG ==============================
'----------------------------------------------------------------------------

'+	Deklaration main

	Public sysProgramName			'Programmname
	Public sysDebug					'Exzessives Logging: True = EIN
	Public sysMessaging				'Systemnachrichten: True = EIN
	Public sysDirLib				'Pfad zur Bibliothek
	Public sysNeedsAdmin			'Global: Administratorrechte benötigt
	Public sysLoggingOn				'True = Logdatei erstellen

'+	Deklaration fctLibInvoke

	Public libInvoke				'Dateiname der aufzurufenden Bibliothek
	Public objFile					'Datei-Objekt
	Public libArr					'Übergabevariable des Bibliotheksinhalts
	Public forReading				'Parameter zum Öffnen einer Textdatei (1 = nur lesen, 2 = lesen/schreiben)

'+	Deklaration Startzeitparameter

	Public yearStart				'Startzeit (Jahr) Format JJJJ
	Public monthStart				'Startzeit (Monat) Format MM
	Public dayStart					'Startzeit (Tag) Format TT
	Public hourStart				'Startzeit (Stunde) Format hh
	Public minuteStart				'Startzeit (Minute) Format mm
	Public secondStart				'Startzeit (Sekunde) Format ss

'---------------------------------------------------------------------------

'+	Programmname festlegen

	sysProgramName	=	"VBS Productivity Suite"
	
'+	Laufzeitparameter

	'On Error Resume Next			'Syntax-/Logikfehler ignorieren
	sysDebug		=	False		'Exzessives Logging: True = EIN
	sysMessaging	=	False		'Systemnachrichten: True = EIN
	sysNeedsAdmin	=	False		'Global: Administratorrechte benötigt
	sysLoggingOn	=	True		'True = Logdatei erstellen
	sysQuiet		=	False		'Verhindert Systemmeldungen, die rein informativ sind

'---------------------------------------------------------------------------

'+	Pfad der Bibliotheken festlegen

	sysDirLib		=	CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\lib"

'+	Funktion für Bibliotheksaufruf

	Function fctLibInvoke(ByVal libInvoke)
	
		forReading	= 1

		Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sysDirLib & "\" & libInvoke, forReading)
		libArr = libArr & vbCrLf & objFile.ReadAll
		objFile.Close
		
	End Function

'+	Bibliotheken definieren

	fctLibInvoke "libSys.vbs"			'Systembibliotheken laden
	fctLibInvoke "libExcel.vbs"			'Excelbibliotheken laden
	fctLibInvoke "libWord.vbs"			'Wordbibliotheken laden

'+	Bibliotheken laden

	ExecuteGlobal libArr

'---------------------------------------------------------------------------

'+	Datum und Uhrzeit erfassen

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN

	fctDateTimeNow		sysProgramName , _
						sysMessaging

	'RETURN:dateStamp				'Datumsformat JJJJ-MM-TT
	'		timeStamp				'Zeitformat hhmmss
	'		dateTimeStamp			'Stempel JJJJ-MM-TT_hhmmss

'+	Startzeit sichern

	yearStart	= year(now)
	monthStart	= thisMonth
	dayStart	= thisDay
	hourStart	= hourNow
	minuteStart	= minuteNow
	secondStart	= secondNow
	
'---------------------------------------------------------------------------

'+	Systemparameter Initialisieren

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	fctSystemInit	sysProgramName , _
					sysMessaging , _
					sysDebug

	'RETURN:WinComputerName			'Computername
	'		WinUserName				'Benutzername der aktiven Sitzung
	'		WinUserPath				'Benutzerverzeichnis
	'		thisFile				'Pfad des aktuell ausgeführten Scripts
	'		thisPath				'Pfad zum aktuell ausgeführten Script
	'		dirTemp					'Pfad des temporären Verzeichnisses

'---------------------------------------------------------------------------

'+	Logdatei-Name festlegen

	nameLogfile	=	dateTimeStamp &" "& sysProgramName &".log"

'+	Kopf der Logdatei vorbereiten

	lineLog		=	"Autor: Bastian Neuwirth " & Chr(169) & " 2021" & vbCrLf & _
					"URL: https://www.supermorph.tech/vbs-productivity-suite" & vbCrLf & vbCrLf & _
					"Zeitstempel:   "& dateTimeStamp & vbCrLf & _
					"Rechner:       "& WinComputerName & vbCrLf & _
					"Benutzer:      "& WinUserName & vbCrLf & _
					"Skript:        "& thisFile & vbCrLf & _
					"Temp. Verz.:   "& dirTemp & vbCrLf & vbCrLf & _
					"============ BEGINN ============" & vbCrLf

'+	Logdatei schreiben

	'CALL:
	'lineLog			'Zu schreibende Zeile
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'thisPath			'Pfad zum aktuell ausgeführten Script
	'nameLogfile		'Pfad der Logdatei
	'sysMessaging		'Systemnachrichten: True = EIN

	fctLogfile		lineLog , _
					sysProgramName , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	'RETURN:VOID

'+	Admin-Status erfassen und vergleichen mit Vorgabe

	'CALL:
	'sysNeedsAdmin		'Global: Administratorrechte benötigt
	'fctNeedsAdmin		'Administratorrechte benötigt
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	fctAdminHandling	sysNeedsAdmin , _
						fctNeedsAdmin , _
						nameMsgBox , _
						sysMessaging , _
						sysDebug

	'RETURN:	isAdmin				'Wahrheitswert über aktuelle Administratorrechte
	'			userPermission		'Aktuelle Berechtigungsstufe
	'			testPermOK			'True = Berechtigungen vorhanden / OK

'---------------------------------------------------------------------------
	
'+	Bibliotheken loggen

	If sysDebug = True Then
	
		'Log vordefinieren
		lineLog	=	"=== START === BIBLIOTHEKEN LADEN ===" & vbCrLf & _
					libArr & vbCrLf & _
					"=== ENDE === BIBLIOTHEKEN LADEN ==="
		'Logdatei schreiben
		fctLogfile		libArr , _
						sysProgramName & " (Bibliotheken)" , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If

'+	Gebe eine Warnung aus, wenn bereits eine Excelinstanz geöffnet ist

	If indicatorExcel = True Then
	
		'CALL:
		'nameMsgBox					'Titel der Messagebox
		'sysMessaging				'Systemnachrichten: True = EIN
		'sysDebug					'Exzessives Logging

		fctXlsOpenInstWarning	sysProgramName , sysMessaging , sysDebug
		
		'RETURN:VOID
		
	End If

'----------------------------------------------------------------------------
' ============================ HAUPTPROGRAMM ===============================
'----------------------------------------------------------------------------

'+	Laufzeitparameter

	sysDebug		=	False		'Exzessives Logging: True = EIN
	sysMessaging	=	False		'Systemnachrichten: True = EIN
	Overwrite		=	False		'True = Überschreibe Zieldatei, wenn sie bereits existiert
	sysQuiet		=	False		'Verhindert Systemmeldungen, die rein informativ sind

'---------------------------------------------------------------------------

'+	ABSCHNITTSBESCHREIBUNG

	#! CODE

'----------------------------------------------------------------------------
' ================================ ENDE ====================================
'----------------------------------------------------------------------------

'+	Laufzeitparameter

	sysDebug		=	False		'Exzessives Logging: True = EIN
	sysMessaging	=	False		'Systemnachrichten: True = EIN
	sysQuiet		=	False		'Verhindert Systemmeldungen, die rein informativ sind

'+	Script beenden

	'CALL:
	'nameMsgBox			'Name des Systemnachrichtenfensters
	'sysMessaging		'Systemnachrichten: True = EIN
	'sysDebug			'Exzessives Logging

	fctSysEnd		sysProgramName , _
					sysMessaging , _
					sysDebug

	'RETURN:VOID
	
WScript.Quit