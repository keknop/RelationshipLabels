'MacroName:BatchAddIRI
'MacroDescription:Adds IRIs for PCC relationship labels in multiple fields.
'$Include "RelationshipLabels!CommonFunctions"

Sub Main
 
   Dim CS As Object
On Error Resume Next
   Set CS = GetObject(,"Connex.Client")
On Error GoTo 0
   If CS Is Nothing Then
      Set CS = CreateObject("Connex.Client")
   End If

   Dim TagList(), Relationships() as String
   Dim nSkipTitles, nSkipNames, nFileError as Integer 
   Dim DELIM, SF4 as String
   
    Dim nRow As Integer
    Dim nChangedFields As Integer
    Dim nAllSFCount As Integer
    Dim nRelSFCount As Integer
    Dim sFieldText As String, sNewFieldText As String
    Dim nFieldType As Integer
    Dim nTag As Integer
    Dim sBreak As String
    Dim sIRI As String, sIRIConcat As String
    Dim bIsControlled As Integer   
   

'#  Set up variables for common strings and non-alphanumeric characters
   DELIM = Chr(223)           'OCLC subfield delimiter
   SF4 = " " & DELIM & "4 "   'Delimiter subfield 4, with spaces
   nChangedFields = 0

   If CS.IsOnline = False Then
      MsgBox("Parts of this macro require that you be logged in to function. Attempting to log in with default profile; if you wish to use a different profile, log in manually and re-run the macro.")
      bool = CS.Logon("","","")
      If CS.IsOnline = False Then 
         MsgBox("Failed to log in with default profile. Please log in manually.")
         Goto Done
      End If
   End If

'# Dialog box
    Begin Dialog dOpt 197, 134, "Options and Settings"
      ButtonGroup .RunOptions
        OkButton  145, 20, 50, 14
        CancelButton  145, 40, 50, 14
           
        Text  5, 5, 45, 8, "Run for..."
        OptionGroup .MacroMode
         OptionButton  5, 30, 59, 10, "&All headings", .OptionButton1                        '#0
            OptionButton  5, 40, 75, 10, "&Name headings only", .OptionButton2                     '#1
            OptionButton  5, 50, 125, 10, "&Title and Name/title headings only", .OptionButton3         '#2
            OptionButton  5, 20, 95, 10, "&Selected heading only", .OptionButton4                  '#3

      Text  5, 75, 30, 8, "Options"
         CheckBox  5, 90, 165, 11, "Attempt to &control all headings when done", .CheckBox1  
            'CheckBox  5, 100, 170, 10, "Try to generate labels from &MARC relator codes", .CheckBox2
            'CheckBox  5, 110, 170, 10, "Display &warnings for pre-AACR2 uses of $e", .CheckBox3
    End Dialog
   Dim DlgChoices As dOpt
   dlg = Dialog(DlgChoices)

   If DlgChoices.RunOptions = -1 Then
      Goto Done
   Else 
      Select Case DlgChoices.MacroMode
         Case 0   '# All headings
            Redim TagList(7)
            TagList(0) = "100"
            TagList(1) = "110"
            TagList(2) = "111"
            TagList(3) = "130"
            TagList(4) = "700"
            TagList(5) = "710"
            TagList(6) = "711"
            TagList(7) = "730"
            nSkipTitles = False  
            nSkipNames = False
         Case 1   '# Names only, no titles
            Redim TagList(5)
            TagList(0) = "100"
            TagList(1) = "110"
            TagList(2) = "111"
            TagList(3) = "700"
            TagList(4) = "710"
            TagList(5) = "711"        
            nSkipTitles = True
            nSkipNames = False
         Case 2   '# Title and name/title only
            Redim TagList(4)
            TagList(0) = "130"
            TagList(1) = "700"
            TagList(2) = "710"
            TagList(3) = "711"
            TagList(4) = "730"            
            nSkipTitles = False
            nSkipNames = True
         Case 3   '# Selected field only
            Goto SingleHeading
      End Select
      Goto MultiHeading
   End If


'#  TODO: MARC Relator Conversion Thingy 

'# TODO: Old janky relator warning

SingleHeading:


   Call BuildRelationshipIndex( Relationships(), nFileError )
   If nFileError = TRUE Then Exit Sub


'# Code replicated from single-field version (RelationshipLabelAddIRI)
    nRow = CS.CursorRow
    CS.CursorColumn = 6
    CS.GetFieldLineUnicode nRow, sFieldText
    bIsControlled = CS.IsHeadingControlled(nRow)
    nTag = GetTag(sFieldText)
    nFieldType = GetFieldType(sFieldText)
    Select Case nFieldType
       Case 0
          'Goto Done
       Case 2
          If nTag MOD 100 = 11 Then
             sBreak="j"
          Else
             sBreak="e"
          End If
       Case 4
          sBreak = "i"
       Case Else
          sBreak = "e"
    End Select
    nRelSFCount = Count(sFieldText, DELIM & sBreak)
    If nRelSFCount = 0 Then
       MsgBox("No " & DELIM & sBreak & " found in this field!")
       Goto Done
    Else
       If bIsControlled = TRUE Then
          If CS.UncontrolHeading() = TRUE Then
          Else
             MsgBox("Headings cannot be modified while controlled. Failed to uncontrol heading; make sure you are logged in.")
             Goto Done
          End If
       End If
       sNewFieldText = ""
       Call RebuildField( sFieldText, sBreak, nFieldType, sNewFieldText, Relationships() )
       CS.SetFieldLine nRow, sFieldText
       nChangedFields = nChangedFields + 1    
   End If
Goto Done

MultiHeading:


   Call BuildRelationshipIndex( Relationships(), nFileError )
   If nFileError = TRUE Then Exit Sub


   Dim nORow, nOCol, n1Row, n7Row, nStartRow as Integer
   nORow = CS.CursorRow
   nOCol = CS.CursorColumn
   nRow = 1
   n1Row = 1
   n7Row = 1
   CS.CursorRow = nRow

'# Find which row (if any) the first 1xx and first 7xx fields are on 
   If CS.GetFieldUnicode("1..", 1, sFieldText) = True Then n1Row = FindMyRow(sFieldText, CS)
   If CS.GetFieldUnicode("7..", 1, sFieldText) = True Then n7Row = FindMyRow(sFieldText, CS)

   j = 1
   For i = 0 To UBound(TagList)
      bool = CS.GetFieldUnicode(TagList(i), j, sFieldText)
      hastitle = False
      Do While bool = True
'#       Tests: Skip to the next field if the field does not match the user-provided name/title settings
         If (InStr(sFieldText, DELIM & "t") Or InStr(Mid(sFieldText, 2,2), "30")) Then hastitle = True Else hastitle = False
'          #All headings                                    #Names only                                #Titles only   
         If (nSkipNames = False And nSkipTitles = False) Or (nSkipTitles = True And hastitle = False) Or (nSkipNames = True And hastitle = True) Then
            If InStr(Left(TagList(i),1),"1") Then nStartRow = n1Row Else nStartRow = n7Row
            nRow = FindMyRow(sFieldText, CS, nStartRow)
            bIsControlled = CS.IsHeadingControlled(nRow)
            nTag = GetTag(sFieldText)
            nFieldType = GetFieldType(sFieldText)
            Select Case nFieldType
               Case 0
                 'Goto Done
               Case 2
                 If nTag MOD 100 = 11 Then
                   sBreak="j"
                 Else
                   sBreak="e"
                 End If
               Case 4
                 sBreak = "i"
               Case Else
                 sBreak = "e"
            End Select
            nRelSFCount = Count(sFieldText, DELIM & sBreak)
            If nRelSFCount > 0 Then
               If bIsControlled = TRUE Then CS.UncontrolHeading
               sNewFieldText = ""
               Call RebuildField( sFieldText, sBreak, nFieldType, sNewFieldText, Relationships() )
               CS.SetFieldLine nRow, sFieldText
               nChangedFields = nChangedFields + 1
            End If
         End If
         j = j + 1
         bool = CS.GetFieldUnicode(TagList(i), j, sFieldText)
      Loop
'#    reset j
   j = 1
   Next i
   
   CS.CursorRow = nORow
   CS.CursorColumn = nOCol   

Done:
   If DlgChoices.RunOptions <> -1 Then  
      If DlgChoices.CheckBox1 = 1 Then bool = CS.ControlHeadingsAll
      bool = CS.GetField("040",1,tmp$)
      If InStr(tmp$, "e rda") = False Then
         warn$ = "Note: This record is not coded as RDA! Be sure to make any other necessary updates." & Chr(13) & Chr(10) & Chr(13) & Chr(10) 
      End If
      MsgBox(warn$ & "Attempted changes to " & nChangedFields & " fields.")
   End If
   
End Sub