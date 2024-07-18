'MacroName:CommonFunctions
'MacroDescription:Common functions for the relationship label macros.

Declare Sub BuildRelationshipIndex(ByRef rels() As String, ByRef nFileError As Integer)
Declare Sub RebuildField( ByRef sFieldText As String, ByRef sBreak As String, ByRef nFieldType As Integer, ByRef sNewFieldText as String, ByRef Relationships() as String )
Declare Function SelectDomain( Label$, Domains% ) As String
Declare Function Count( InWhat$, Find$ ) As Integer
Declare Function GetTag ( FieldData$ ) As Integer
Declare Function GetFieldType( FieldData$ ) As Integer
Declare Function FindRelationship( FieldType%, ByRef Search$, ByRef Index() As String ) As String
Declare Function IsolateLabel ( CheckText$, ByRef RightText$ )
Declare Function UpdateLabel ( Label$, Optional Replacement )
Declare Function FindMyRow ( sLookForMe as String, ByRef CS as Object, Optional nStartFrom As Variant ) As Integer


Sub RebuildField( ByRef sFieldText As String, ByRef sBreak As String, ByRef nFieldType As Integer, ByRef sNewFieldText as String, ByRef Relationships() As String )
      sIRIConcat = ""
      sRightOfLabel = ""
     sLabel = "" 
     DELIM = Chr(223)           'OCLC subfield delimiter
      SF4 = " " & DELIM & "4 "   'Delimiter subfield 4, with spaces
      nAllSFCount = Count(sFieldText, DELIM)

      For i = 0 to nAllSFCount
         s = GetField(sFieldText, i+1, DELIM)
         If i = 0 Then sNewFieldText = sNewFieldText & s                     'The first "field" will always be the tag up through the first delimiter, 
                                                                             'so initialize it outside the loop to avoid having to do logic tests to see 
                                                                             'if the first subfield is an implied $a or not    
         If Left(s, 1) = sBreak Then
            sLabel = IsolateLabel(s, sRightOfLabel)
            sLabel = UpdateLabel(sLabel)
            sIRI = FindRelationship(nFieldType, sLabel, Relationships() )
            If Len(sIRI) > 0 Then 

              'MsgBox("Debug: Result found for label " & sLabel & ": " & sIRI)
'#             First check to see if we got an ERR: message instead of an IRI
               If InStr(Left(sIRI, 4), "ERR:") Then
                  sIRI = Mid(sIRI, 6)
                  MsgBox("Problem with label " & sLabel & " for heading " & Chr(13) & Chr(10) & sFieldText & ":" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sIRI & " No action will be taken.")
                  sIRI = ""
           
'#             Next check to see if the IRI is already present in the field and skip it if so
               ElseIf InStr(sFieldText, sIRI) Then
                 'MsgBox("Debug: IRI for " & sLabel & " already present in field! Skipping.")
                  sIRI = ""                  

               Else
                  sIRIConcat = sIRIConcat & SF4 & sIRI
               End If
            End If
            sNewFieldText = sNewFieldText & DELIM & sbreak & " " & sLabel & sRightOfLabel
            sRightOfLabel = ""
         Else
            If i > 0 Then sNewFieldText = sNewFieldText & DELIM & s              'If it's not a label field, add it as-is to the reconstructed field
         End If    
      Next i

'#    Finally, append the IRIs to the field text and replace it.
      sFieldText = sNewFieldText & sIRIConcat
End Sub    

Sub BuildRelationshipIndex(ByRef rels() As String, ByRef nFileError As Integer)
   Dim FSO As Object, TxtFile As Object, Data As Object
   Dim sFileName As String

   sFileName = Environ$("APPDATA") & "\OCLC\Connex\Macros\RelationshipTable.txt"
   Set FSO = CreateObject("Scripting.FileSystemObject")
   On Error Goto FileError
   Set TxtFile = FSO.GetFile(sFileName)
   Set Data = TxtFile.OpenAsTextStream(1)
   
'Dump the whole file to a string, count the number of line breaks to get the necessary array size, then feed the string into the array. GetField only takes
'single-character delimiters, so have to test for and strip stray carriage returns. Not sure this is actually any more efficient than the previous approach, which
'required opening the file twice, but  I'm blind guessing which file manipulation and array features are actually supported in SBL, this actually compiles, and the 
'file size is small enough that the inefficiency is unlikely to matter.

   tmp$ = Data.Readall
   'MsgBox("Debug: tmp$: " & tmp$)
   eol$ = Chr(13) & Chr(10)
   Data.Close
   
   ReDim rels( Count(tmp$, eol$) + 1)
      
   For k = 0 To UBound(rels)
      st$ = GetField(tmp$, k+1, Chr(10) )
      If InStr( Right(st$, 1), Chr(13) ) Then st$ = Left(st$, Len(st$) - 1)
      rels(k) = st$
   Next k
   tmp$ = ""
   
   Exit Sub

FileError:
   nFileError = TRUE
   MsgBox("Failed to find/open the file " & VbCrLf & Chr(13) & Chr(10) & sFileName & ". " & Chr(13) & Chr(10) & "Make sure you downloaded this file from https://github.com/keknop/RelationshipLabels. It should be saved in the same folder as your Connexion macro files.")
End Sub

Function GetTag( FieldData$ ) As Integer
   On Error Goto TagError
   tagstring$ = Left(FieldData$, 3)
   tagint% = CInt(tagstring$)
   
   GetTag = tagint%
'  MsgBox("Debug: Tag is " & tagint%)
   Goto Done

TagError:
   MsgBox("Tag " & tagstring$ & " contains non-numeric characters. Skipping field.")
   
Done:   
End Function

Function GetFieldType( FieldData$ ) As Integer
   x = 0                                                                '0 = Invalid data
   tag% = GetTag(FieldData$)
   tagend% = tag% MOD 100
   sInd1$ = Mid(FieldData$, 4, 1)                                       
   sInd2$ = Mid(FieldData$, 5, 1)
   serr$ = ""
   
   
   If InStr(FieldData$, Chr(223) & "t") Then
      If InStr(FieldData$, Chr(223) & "j") AND tagend% = 11 Then
         x = 0
         serr$ = "Error: " & sTag$ & " fields containing title information ($t) cannot contain an agent relationship ($j). Field will be skipped."
         Goto FieldTypeError
      ElseIf InStr(FieldData$, Chr(223) & "e") Then
         x = 0
         serr$ = "Error: " & sTag$ & " fields containing title information ($t) cannot contain an agent relationship ($e). Field will be skipped."
         Goto FieldTypeError
      Else
         x = 4                                                          '4 = Name/title 
      End If
   End If
   
   Select Case tag%
      Case 100, 700
         If InStr(FieldData$, Chr(223) & "e") AND x <> 4 Then 
            If InStr(sInd1$, "0") OR InStr(sInd1$, "1") Then
               x = 1                                                    '1 = Person
            ElseIf InStr(sInd1$, "3") Then   
               x = 3                                                    '3 = Family
            Else 
               x = 0
               serr$ = serr$ & chr(13) & chr(10) & "Error: " & sTag$ & " does not have a valid first indicator."
            End If   
         End If

      Case 110, 710
         If InStr(FieldData$, Chr(223) & "e") AND x <> 4 Then x = 2     '2 = Corporate Body
      
      Case 111, 711
         If InStr(FieldData$, Chr(223) & "j") AND x <> 4 Then x = 2     '2 = Corporate Body   
      
      Case Else
         'MsgBox("This macro only works on 1xx and 7xx fields.")     
   End Select   

FieldTypeError:
   If InStr(serr$, "Error") Then 
      MsgBox(serr$)
   End If
   
   GetFieldType = x

End Function

Function FindRelationship( FieldType%, ByRef Search$, ByRef Index() As String ) As String
   ft% = FieldType% + 2
   entry$ = ""
   specialcase% = 0
   domain$ = ""
   
'# Set up special handling for ambiguous labels
   Select Case Search$
      Case "creator", "interviewee", "interviewer"                'work/expression
         specialcase% = 1
      Case "dedicatee", "honouree"                                'work/item
         specialcase% = 2
      Case "restorationist"                                       'expression/item
         specialcase% = 3
      Case "related person", "related body", "related family"     'work/expression/manifestation/item
         specialcase% = 4
      Case Else
         specialcase% = 0
   End Select
   domain$ = SelectDomain(Search$, specialcase%)
   'MsgBox("Debug: specialcase for search term " & Search$ & " is " & specialcase% & " giving domain result " & domain$)

   For i = LBound(Index()) To UBound(Index())
      check$ = GetField(Index(i), 1, "|")
      domcheck$ = GetField(Index(i), 2, "|")

'#    Check if the strings match AND if they are the same length, to prevent false matches like "director" being 
'#    found in "casting director"
'#    If the label matches multiple WEMI domains, query the user and also match the domain from the IRI

      If InStr(Search$, check$) AND Len(Search$) = Len(check$) Then
         'MsgBox("Debug: potential match found. domcheck is " & domcheck$)
         If specialcase% = 0 Then
            entry$ = GetField(Index(i), ft%, "|")
          Else
            If domain$ = domcheck$ Then 
               'MsgBox("Debug: domain is " & domain$ & ", domcheck is " & domcheck$)
               entry$ = GetField(Index(i), ft%, "|")
            End If
         End If

'#       Check to see if there is a USE: reference to a different label instead of an IRI in the table, e.g. "USE: related body" if "related 
'#       person" has been used in a 710. If there is, hop back three spaces in the array and resume searching.

         If InStr(Left(entry$, 4), "USE:") Then
           'MsgBox("Debug: Whoopsie doodles: searching for " & Search$ & " found " & entry$ & " at array index " & i)
            Search$ = UpdateLabel(Search$, Mid(entry$, 6))
            j = i-3
            
            'MsgBox("Debug: Now searching for " & Search$ & " starting at array index " & j)
            For j = j To UBound(Index())
               check$ = GetField(Index(j), 1, "|")
               If InStr(Search$, check$) AND Len(Search$) = Len(check$) Then
                  If specialcase% = 0 Then
                     entry$ = GetField(Index(j), ft%, "|")
                     FindRelationship = entry$
                     'MsgBox("Searched instead for " & Search$ & " and found " & entry$)
                     Goto Done
                  Else
                     If domain$ = domcheck$ Then 
                       entry$ = GetField(Index(j), ft%, "|")
                       FindRelationship = entry$
                       Goto Done
                     End If
                  End If
               End If  
            Next
         End If 
      End If
      If Len(entry$) > 0 Then
         FindRelationship = entry$
         Goto Done    
      End If
   Next

'# Whoops, we went through the whole array without a match. Show an error message and return an empty value:
   MsgBox("Error: the label " & Search$ & " was not found in the PCC relationship label list. Check the spelling.")
   FindRelationship = ""
   entry$=""
     
Done:
'MsgBox("Debug: Findrelationship is " & entry$)
End Function

Function IsolateLabel ( CheckText$, ByRef RightText$ )

' Takes a relationship subfield (CheckText$), chops off the initial subfield letter + space, then separates the label text from any
' trailing punctuation and/or spaces. We also pass in an empty string (RightText$) to store the trailing characters so they can 
' be recombined later.

   CheckText$ = Mid(CheckText$, 3)                                         'Chop off the subfield delimiter + space
   For i = 0 To Len(CheckText$)
      r$ = Mid(CheckText$, Len(CheckText$) - i, 1)
         Select Case Asc(r$)
            Case 32 To 40, 42 To 47, 58 To 64, 91 To 96, 123 To 126        'Any ASCII space or punctuation except for ")"
               RightText$ = r$ & RightText$
            Case Else
               CheckText$ = Trim(Mid(CheckText$, 1, Len(CheckText$) - Len(RightText$)))
               Goto Done
         End Select
   Next 
Done:
   IsolateLabel = CheckText$
End Function


Function UpdateLabel ( Label$, Optional Replacement )
   update$ = ""
   If IsMissing(Replacement) Then
      'MsgBox("Debug: No replacement text passed")
      
      Select Case Label$
      
'# Spelling sanity checks: RDA/the LC-PCC label list are not consistent in preferring American or Commonwealth spelling
'# Commonwealth spelling used: hono(u)ree, colo(u)rist
         Case "honoree" : update$ = "honouree"
         Case "colorist" : update$ = "colourist"
'# (Predominantly?) American spelling used: draftsman (draughtsman)
         Case "draughtsman" : update$ = "draftsman"
   
'# Some terms have changed slightly or significantly
         Case "contributor"   :  update$ = "creator"
         Case "host institution" : update$ = "hosting institution"
         Case "sponsoring body" : update$ = "sponsor"
         Case "participant in treaty" : update$ = "treaty participant"
         Case "composer (expression)" : update$ = "contributor of music"
         Case "interviewee (expression)" : update$ = "interviewee"
         Case "interviewer (expression)" : update$ = "interviewer"
         Case "make-up artist" : update$ = "makeup artist"
         Case "on-screen participant" : update$ = "onscreen participant"
         Case "restorationist (expression)" : update$ = "restorationist"
         Case Else
            'MsgBox("Debug: No label updates found")
      End Select

'# new terms: academic supervisor; description of; civil defendant, criminal defendant (narrower terms of defendant); contributor to amalgamation; 
'# contributor to performance (new intermediate between contributor to amalgamation and various and sundry performer terms); research supervisor; reviser

   Else
      update$ = Replacement
   End If

   If Len(update$) > 0 Then
      MsgBox("Label " & Label$ & " updated to " & update$)
      UpdateLabel = update$
   Else
      UpdateLabel = Label$
   End If
  
End Function

Function SelectDomain( Label$, Domains% ) As String
   Dim domain$

   Select Case Domains%
      Case 1
         Begin Dialog WEChoice 180, 116, "Which " & Label$ & "?"
            Text 4, 4, 170, 24, "The label " & Label$ & " could relate an agent to multiple things. Select one (or type the first letter):"
            OptionGroup .WEMI
               OptionButton 8, 28, 80, 12, "&Work"
               OptionButton 8, 40, 80, 12, "&Expression"
            OKButton 4, 80, 40, 20
            CancelButton 48, 80, 40, 20
         End Dialog
         Dim LabelDomainWE As WEChoice
         z = Dialog(LabelDomainWE)
         If LabelDomainWE.WEMI = 0 Then
            domain$ = "w"
         Else 
            domain$ = "e"
         End If

      Case 2
         Begin Dialog WIChoice 180, 116, "Which " & Label$ & "?"
            Text 4, 4, 170, 24, "The label " & Label$ & " could relate an agent to multiple things. Select one (or type the first letter):"
            OptionGroup .WEMI
               OptionButton 8, 28, 80, 12, "&Work"
               OptionButton 8, 40, 80, 12, "&Item"
            OKButton 4, 80, 40, 20
            CancelButton 48, 80, 40, 20
         End Dialog
         Dim LabelDomainWI As WIChoice
         z = Dialog(LabelDomainWI)
         If LabelDomainWI.WEMI = 0 Then
            domain$ = "w"
         Else 
            domain$ = "i"
         End If

      Case 3
         Begin Dialog EIChoice 180, 116, "Which " & Label$ & "?"
            Text 4, 4, 170, 24, "The label " & Label$ & " could relate an agent to multiple things. Select one (or type the first letter):"
            OptionGroup .WEMI
               OptionButton 8, 28, 80, 12, "&Expression"
               OptionButton 8, 40, 80, 12, "&Item"
            OKButton 4, 80, 40, 20
            CancelButton 48, 80, 40, 20
         End Dialog
         Dim LabelDomainEI As EIChoice
         z = Dialog(LabelDomainEI)
         If LabelDomainEI.WEMI = 0 Then
            domain$ = "e"
         Else 
            domain$ = "i"
         End If

      Case 4
         Begin Dialog WEMIChoice 180, 116, "Which " & Label$ & "?"
            Text 4, 4, 170, 24, "The label " & Label$ & " could relate an agent to multiple things. Select one (or type the first letter):"
            OptionGroup .WEMI
               OptionButton 8, 28, 80, 12, "&Work"
               OptionButton 8, 40, 80, 12, "&Expression"
               OptionButton 8, 52, 80, 12, "&Manifestation"
               OptionButton 8, 64, 80, 12, "&Item"
            OKButton 4, 80, 40, 20
            CancelButton 48, 80, 40, 20
         End Dialog
         Dim LabelDomainWEMI As WEMIChoice
         z = Dialog(LabelDomainWEMI)
         If LabelDomainWEMI.WEMI = 0 Then
            domain$ = "w"
         ElseIf LabelDomainWEMI.WEMI = 1 Then
            domain$ = "e"
         ElseIf LabelDomainWEMI.WEMI = 2 Then
            domain$ = "m"            
         ElseIf LabelDomainWEMI.WEMI = 3 Then
            domain$ = "i"            
         End If
   End Select
   
   SelectDomain = domain$

End Function

Function FindMyRow ( sLookForMe as String, ByRef CS as Object, Optional nStartFrom As Variant ) As Integer
   testrow% = 1
   rowexists% = True
   found% = False
   testtext$ = ""
   
   If IsMissing(nStartFrom) Then textrow% = 1 Else testrow% = nStartFrom
   
   Do While rowexists% = True And found% = False
      bool = CS.GetFieldLineUnicode(testrow%, testtext$)
'#    Make sure it's a valid numeric tag; if so, check if the tags match; if so, check if the field content is equivalent
      If GetTag(testtext$) Then
         If GetTag(testtext$) = GetTag(sLookForMe) Then
            If InStr(sLookForMe, testtext$) And (Len(sLookForMe) = Len(testtext$)) Then
               FindMyRow = testrow%
               found% = True
               Goto Done
         End If
         End If
      End If
'#    Advance the row count and try to move the cursor to the new row. If     
      testrow% = testrow% + 1
      CS.CursorRow = testrow%
      If CS.CursorRow <> testrow% Then rowexists% = False
   Loop
   FindMyRow = False

Done:
End Function


'################################################################################
'Count function from the NikAdds library provided by Joel Hahn, http://www.hahnlibrary.net/libraries/oml/connex.html

Function Count(InWhat$, Find$) As Integer
  place = 1
  ct = 0
  Do While InStr(place, InWhat$, Find$, 0)
    place = InStr(place, InWhat$, Find$, 0) + 1
    ct = ct + 1
  Loop
  Count = ct
End Function