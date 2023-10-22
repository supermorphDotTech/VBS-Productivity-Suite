'---------------------------------------------------------------------------

' Bibliothek:

' fctWordFormField			-> Formatierung eines Formularfeld-Schreibcodes
' fctWordBookmark			-> Formatierung eines Textmarken-Schreibcodes
' fctWordCopyWrite			-> Quell-Worddokument kopieren und Formularfelder füllen

'---------------------------------------------------------------------------

'Allgemeine Variablen

Public objWorddoc				'Objekt für eine Wordinstanz

'---------------------------------------------------------------------------

'+	Deklaration fctWordFormField

Public wordFormField			'Name des Formularfelds
Public wordFormFieldCont		'Inhalt des Formularfelds
Public wordTempWriteCommands	'Zwischenspeicher für Formularfeld-Schreibcode

'+	Formatierung eines Formularfeld-Schreibcodes

	'CALL:
	'wordFormField			'Name des Formularfelds
	'wordFormFieldCont		'Inhalt des Formularfelds
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'wordWriteCommands	= fctWordFormField(wordFormField , wordFormFieldCont , _
	'										nameMsgBox , sysMessaging , sysDebug )

Function fctWordFormField(ByVal wordFormField,ByVal wordFormFieldCont,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctWordFormField (" & nameMsgBox & ")"
	
'+	Engaben als Code formatieren

	'Log vordefinieren
	lineLog	=	"Formularfeld-Schreibcode wird formatiert." & vbCrLf & _
				"wordFormField: " & wordFormField & vbCrLf & _
				"wordFormFieldCont: " & wordFormFieldCont
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	wordTempWriteCommands	= "objWorddoc.FormFields(" & Chr(34) & wordFormField & Chr(34) & ").Result = " & Chr(34) & wordFormFieldCont & Chr(34)

	fctWordFormField	= wordTempWriteCommands

	'Log vordefinieren
	lineLog	=	"Formularfeld-Schreibcode generiert." & vbCrLf & _
				"wordTempWriteCommands: " & wordTempWriteCommands
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	nameMsgBoxStacked	= nameMsgBox
	
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctWordBookmark

Public wordBookmark				'Name der Textmarke
Public wordBookmarkCont			'Zeichenkette zur Textmarke

'+	Formatierung eines Textmarken-Schreibcodes

	'CALL:
	'wordBookmark			'Name der Textmarke
	'wordBookmarkCont		'Zeichenkette zur Textmarke
	'nameMsgBox				'Titel der Messagebox
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'wordWriteCommands	= fctWordBookmark(wordBookmark , wordBookmarkCont , _
	'										nameMsgBox , sysMessaging , sysDebug )

Function fctWordBookmark(ByVal wordBookmark,ByVal wordBookmarkCont,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctWordBookmark (" & nameMsgBox & ")"
	
'+	Engaben als Code formatieren

	'Log vordefinieren
	lineLog	=	"Textmarken-Schreibcode wird formatiert." & vbCrLf & _
				"wordBookmark: " & wordBookmark & vbCrLf & _
				"wordBookmarkCont: " & wordBookmarkCont
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	wordTempWriteCommands	= "objWorddoc.Bookmarks(" & Chr(34) & wordBookmark & Chr(34) & ").Range = " & Chr(34) & wordBookmarkCont & Chr(34)

	fctWordBookmark	= wordTempWriteCommands

	'Log vordefinieren
	lineLog	=	"Textmarken-Schreibcode generiert." & vbCrLf & _
				"wordTempWriteCommands: " & wordTempWriteCommands
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging

	nameMsgBoxStacked	= nameMsgBox
		
End Function

'---------------------------------------------------------------------------

'+	Deklaration fctWordCopyWrite

Public wordSourceFile			'Pfad zum Quell-Dokument
Public wordDestFile			'Pfad zum neuen Dokument
Public wordWriteCommands		'Formatierte Formularfeld-Schreibkommandos (siehe fctWordFormField)

'+	Quelldatei kopieren und Formularfelder füllen

	'CALL:
	'wordSourceFile			'Pfad zum Quell-Dokument
	'wordDestFile			'Pfad zum neuen Dokument
	'wordWriteCommands		'Formatierte Formularfeld-Schreibkommandos (siehe fctWordFormField)
	'Overwrite				'True = Überschreibe Zieldatei, wenn sie bereits existiert
	'sysMessaging			'Systemnachrichten: True = EIN
	'sysDebug				'Exzessives Logging

	'fctWordCopyWrite	wordSourceFile , _
	'					wordDestFile , _
	'					wordWriteCommands , _
	'					Overwrite , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VOID

Function fctWordCopyWrite(ByVal wordSourceFile,ByVal wordDestFile,ByVal wordWriteCommands,ByVal Overwrite,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "fctWordCopyWrite (" & nameMsgBox & ")"
	
	Set objWord = CreateObject("Word.Application")

	'Log vordefinieren
	lineLog	=	"Wordinstanz geöffnet."
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

	fctChkFile	wordSourceFile , _
				True , _
				nameMsgBox , _
				sysMessaging , _
				sysDebug

	'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

	If fileExist = True Then

		'Log vordefinieren
		lineLog	=	"Worddokument kopieren und Textmarken füllen:" & vbCrLf & _
					"Overwrite = " & Overwrite & vbCrLf & _
					"Quelle: " & wordSourceFile & vbCrLf & _
					"Ziel: " & wordDestFile & vbCrLf & vbCrLf & _
					"Befehle:" & vbCrLf & wordWriteCommands
					
'+	Von Quelle kopieren und schreiben

		'CALL:
		'filePath				'Pfad und Name der Datei
		'fileExistWarning		'True = Gibt eine Warnung aus, wenn die Datei nicht existiert
		'nameMsgBox				'Name des Systemnachrichtenfensters
		'sysMessaging			'Systemnachrichten: True = EIN
		'sysDebug				'Exzessives Logging

		fctChkFile	wordDestFile , _
					False , _
					nameMsgBox , _
					sysMessaging , _
					sysDebug

		'RETURN:fileExist		'Antwort: True = Datei ex.; False = Datei ex. nicht

		If fileExist = False Then
		
			objFSO.CopyFile wordSourceFile, wordDestFile, Overwrite
			Set objWorddoc   = objWord.Documents.Open(wordDestFile)
			
			'Log vordefinieren
			lineLog	=	"Datei kopieren und Textmarken füllen." & vbCrLf & _
						"Pfad: " & vbCrLf & wordDestFile & vbCrLf & _
						"Befehle:" & vbCrLf & wordWriteCommands
						
			Execute wordWriteCommands
			fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

			objWorddoc.save
			objWorddoc.Close
			
		ElseIf fileExist = True And Overwrite = False Then
			
			'Log vordefinieren
			lineLog	=	"Datei existiert bereits und wurde nicht überschrieben." & vbCrLf & _
						"Pfad: " & wordDestFile
				
		ElseIf fileExist = True And Overwrite = True Then
			
			objFSO.CopyFile wordSourceFile, wordDestFile, Overwrite
			Set objWorddoc   = objWord.Documents.Open(wordDestFile)
			
			Execute wordWriteCommands
			fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

'+	Worddokument speichern und schließen

			objWorddoc.save
			objWorddoc.Close
			
			'Log vordefinieren
			lineLog	=	"Datei wurde überschrieben." & vbCrLf & _
						"Pfad: " & wordDestFile & vbCrLf & vbCrLf & _
						"Code: " & vbCrLf & wordWriteCommands
			
		End If

		'Logdatei schreiben
		fctLogfile		lineLog , _
						nameMsgBoxStacked , _
						thisPath , _
						nameLogfile , _
						sysMessaging
						
	End If

	objWord.Application.Quit
		
	'Log vordefinieren
	lineLog	=	"Wordinstanz geschlossen"
	'Logdatei schreiben
	fctLogfile		lineLog , _
					nameMsgBoxStacked , _
					thisPath , _
					nameLogfile , _
					sysMessaging
						
	nameMsgBoxStacked	= nameMsgBox
		
End Function