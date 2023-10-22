'---------------------------------------------------------------------------

' Bibliothek:

' TEMPLATEMODUL					-> MODULBESCHREIBUNG

'---------------------------------------------------------------------------

'Allgemeine Variablen

'Public VARIABLE					'BESCHREIBUNG

'---------------------------------------------------------------------------

'+	Deklaration TEMPLATEMODUL

'Public VARIABLE					'BESCHREIBUNG

'+	MODULBESCHREIBUNG

	'CALL:
	'VARIABLE						'BESCHREIBUNG
	'nameMsgBox						'Titel der Messagebox
	'sysMessaging					'Systemnachrichten: True = EIN
	'sysDebug						'Exzessives Logging

	'TEMPLATEMODUL		VARIABLE , _
	'					nameMsgBox , _
	'					sysMessaging , _
	'					sysDebug
	'fctErrorHandling	nameMsgBox , sysMessaging , sysDebug

	'RETURN:VARIABLE				'BESCHREIBUNG

Function TEMPLATEMODUL(ByVal VARIABLE,ByVal nameMsgBox,ByVal sysMessaging,ByVal sysDebug)

	On Error Resume Next

	nameMsgBoxStacked	= "TEMPLATEMODUL (" & nameMsgBox & ")"

'+	ABSCHNITTSBESCHREIBUNG

	#! CODE

End Function
