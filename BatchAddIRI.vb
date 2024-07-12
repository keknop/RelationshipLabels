'$Include "RelationshipLables!RelationshipLabelAddIRI"

Sub Main
 
	Dim CS As Object
On Error Resume Next
	Set CS = GetObject(,"Connex.Client")
On Error GoTo 0
	If CS Is Nothing Then
		Set CS = CreateObject("Connex.Client")
	End If

	Dim TagList(), Relationships() as String
	Dim sFieldData as String
	Dim nSkipTitles, nSkipNames, nFileError as Integer 
	Dim DELIM, SF4 as String

'# Set up variables for common strings and non-alphanumeric characters
	DELIM = Chr(223)           'OCLC subfield delimiter
	SF4 = " " & DELIM & "4 "   'Delimiter subfield 4, with spaces

'#	Dialog box placeholder
    Begin Dialog dOpt 197, 134, "Options and Settings", .StartDlg
		ButtonGroup .RunOptions
        OkButton  145, 20, 50, 14
        CancelButton  145, 40, 50, 14
           
        Text  5, 5, 45, 8, "Run for..."
        OptionGroup .MacroMode
			OptionButton  5, 30, 59, 10, "&All headings", .OptionButton1								'#0
            OptionButton  5, 40, 75, 10, "&Name headings only", .OptionButton2						'#1
            OptionButton  5, 50, 125, 10, "&Title and Name/title headings only", .OptionButton3		'#2
            OptionButton  5, 20, 95, 10, "&Selected heading only", .OptionButton4						'#3
		Text  5, 75, 30, 8, "Options"
			CheckBox  5, 90, 165, 11, "Attempt to &control all headings when done", .CheckBox1	
            'CheckBox  5, 100, 170, 10, "Try to generate labels from &MARC relator codes", .CheckBox2
            'CheckBox  5, 110, 170, 10, "Display &warnings for pre-AACR2 uses of $e", .CheckBox3
    End Dialog
	Dim DlgChoices As dOpt
	dlg = Dialog(DlgChoices)

	If dlg = -1 Then
		Goto Done
	Else 
		Select Case DlgChoices.MacroMode
			Case 0	'# All headings
				Redim TagList(8)
				TagList = { "100", "110", "111", "130", "700", "710", "711", "730" }
				nSkipTitles = False	
				nSkipNames = False
			Case 1	'# Names only, no titles
				Redim TagList(6)
				TagList = { "100", "110", "111", "700", "710", "711" }
				nSkipTitles = True
				nSkipNames = False
			Case 2	'# Title and name/title only
				Redim TagList(5)
				TagList = { "130", "700", "710", "711", "730" }
				nSkipTitles = False
				nSkipNames = True
			Case 3	'# Selected field only
				Goto SingleHeading
		End Select
		Goto MultiHeading
	End If

	Call BuildRelationshipIndex( Relationships(), nFileError )
	If nFileError = TRUE Then Exit Sub


'#  TODO: MARC Relator Conversion Thingy 

'#	TODO: Old janky relator warning

SingleHeading:
'# TODO: Process goes here

Goto Done

MultiHeading:

j = 1
For i = 0 To UBound(TagList)
	bool = CS.GetFieldUnicode(TagList(i), j, sFieldData)
	hastitle = False
	Do While bool = True
		If (InStr(sFieldData, DELIM & "t") Or InStr(sFieldData, "30")) Then hastitle = True Else hastitle = False
		If (nSkipNames = False And nSkipTitles = False) Or (nSkipTitles = True And hastitle = False) Or (nSkipNames = True And hastitle = False) Then
'# TODO: process goes here		
		End If	
		j = j + 1
		bool = CS.GetFieldUnicode(TagList(i), j, sFieldData)
	Loop
'#	reset j 
	j = 1
Next i
   
Done:   
End Sub