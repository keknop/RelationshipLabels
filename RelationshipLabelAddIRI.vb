'MacroName:RelationshipLabelAddIRI
'MacroDescription:Adds RDA relationship element IRIs based on LC-PCC relationship labels in the selected field.
'$Include "RelationshipLabels!CommonFunctions"


Sub Main
 
   Dim CS As Object
On Error Resume Next
   Set CS = GetObject(,"Connex.Client")
On Error GoTo 0
   If CS Is Nothing Then
      Set CS = CreateObject("Connex.Client")
   End If
   
   Dim DELIM As String
   Dim Relationships() As String
   Dim nFileError As Integer

'# Set up variables for common strings and non-alphanumeric characters
   DELIM = Chr(223)           'OCLC subfield delimiter


   Call BuildRelationshipIndex( Relationships(), nFileError )
   If nFileError = TRUE Then Exit Sub
  
   Dim nRow As Integer
   Dim nAllSFCount As Integer
   Dim nRelSFCount As Integer
   Dim sFieldText As String, sNewFieldText As String
   Dim nFieldType As Integer
   Dim nTag As Integer
   Dim sBreak As String
   Dim sIRI As String, sIRIConcat As String
   Dim bIsControlled As Integer   
  
'# Get the content from the current cursor row and check if it's controlled
   nRow = CS.CursorRow
   CS.CursorColumn = 6
   CS.GetFieldLineUnicode nRow, sFieldText
   bIsControlled = CS.IsHeadingControlled(nRow)
   'MsgBox("Debug: bIsControlled = " & bIsControlled)
   
'# Determine what kind of relationship field we're looking at and set the delimiter variable as the appropriate subfield. (If it's 
'# not a 1xx or 7xx field, GetFieldType will kick up an error message.) Fields 111 and 711 need special handling because they use 
'# subfield j instead of e for agent relationships.

   nTag = GetTag(sFieldText)
   nFieldType = GetFieldType(sFieldText)
   Select Case nFieldType
      Case 0
         Goto Done
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

   'MsgBox("Debug: sBreak = " & sBreak)
   
'  # Verify there are actually relationship subfields in the field, as identified by the subfield letter in sBreak.
'  # If there are, loop through each subfield. If the subfield starts with sBreak, trim off the subfield letter, any
'  # starting or ending whitespace, and trailing punctuation. Then loop through each subfield in the field text. 
'  # If it's a relationship subfield, trim any whitespace and ending punctuation from the right and the delimiter 
'  # and subfield from the left. Then search the list for the label and corresponding IRI based on the type of relationship.

   nRelSFCount = Count(sFieldText, DELIM & sBreak)

   If nRelSFCount = 0 Then
      MsgBox("No " & DELIM & sBreak & " found in this field!")
      Goto Done
   Else
      If bIsControlled = TRUE Then
         If CS.UncontrolHeading() = TRUE Then
            'MsgBox("Debug: Heading uncontrolled.")
         Else
            MsgBox("Headings cannot be modified while controlled. Failed to uncontrol heading; make sure you are logged in.")
            Goto Done
         End If
      End If

      sNewFieldText = ""
     
      Call RebuildField( sFieldText, sBreak, nFieldType, sNewFieldText, Relationships() )

      'MsgBox("Debug: Adding IRI(s): " & sIRIConcat)
      'MsgBox("Debug: Final field value: " & sFieldText)
      CS.SetFieldLine nRow, sFieldText    
   End If

Done:
End Sub  