' USES THE FOLLOWING GLOBAL VARIABLES
Global PMUserForm_Abort_Pressed As Boolean
    
Public Sub Highlight_Passive_Voice()
'
' Highlight_Passive_Voice Macro
' This macro identifies "to be" verbs and highlights them,
'   if preceded by a space and followed by a space or common
'   punctuation mark.
' This macro identifies negated contractions of a few
'   "to be" verbs and highlights them.
' This macro identifies a few expletives and strikes
'   them out, while highlighting them in red.
' When finished, it reports results in a Message Box.
'
' V5.0
'   Adds progress bar
'   Displays long message at end
' V6
'   Saves file with new file name at start
'   and end of run
' V7
'   Adds stats for weak verb and preposition fractions.
'   Makes abort button work more clearly with new code for
'       AbortButton_Click().
'
' Uses:
'	Display_Long_MsgBox
'   Highlight_String
'   SaveWithNewFileName
'   Split
'   PMUserForm
'
' USES GLOBAL VARIABLES (Though not proudly):
'   PMUserForm_Abort_Pressed As Boolean

    Const CHARS = " .!?,;:""'()[]{}-_<>"
    Dim i As Long
    Dim J As Long
    Dim TimesFound As Long
    Dim TotalWordsToCheck As Long
    Dim WordsChecked As Long
    Dim PercentDone As Single
    Dim Search_Word As String
    Dim Word_List() As String
    Dim ProgressMessageString As String
    Dim TempMessageString As String
	Dim MessageString As String
    Dim ToBeVerbCount As Long
    Dim PrepositionCount As Long

    
    ProgessMessageString = "Paramedic Method Analysis "
    ToBeVerbCount = 0
    PrepositionCount = 0
    
    Dim Lone_Words(1 To 5) As String
    
    ' Save current version, if hasn't been saved recently
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    ' From http://msdn.microsoft.com/en-us/library/office/ff835916(v=office.14).aspx
    ' Application.StatusBar Property (Excel)
    ' Application.DisplayStatusBar = True
    
    Application.DisplayStatusBar = True
    Application.StatusBar = ("Performing Grammar Check . . . ")
    Application.StatusBar = "Performing Grammar Check . . . "
    
    Word_List = Split("is, are, were, was, will, be, been, being, shall, am")
    Preposition_List = Split("at, in, on, to, of, for, as, with, by, from, into, onto, than, that, under, over, toward, towards, until, up, upon, within, without")
    Lone_Words(1) = "isn't"
    Lone_Words(2) = "aren't"
    Lone_Words(3) = "won't"
    Lone_Words(4) = "wasn't"
    Lone_Words(5) = "weren't"
    
    TotalWordsToCheck = UBound(Word_List) + UBound(Lone_Words) + UBound(Preposition_List) + 8
    ' 8 Represents number of expletives
    WordsChecked = 0
    
    MessageString = "" & Now & " Starting . . ." & vbCr
    PMUserForm.Show
	PMUserForm_Abort_Pressed = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    TempMessageString = "Saving file with new 'Highlighted' name"
    PMUserForm.IncrementProgressBar 0, "0% Complete: " & TempMessageString
    
    Call SaveWithNewFileName

    
    ' Hunt for each word in the word list, when preceded by
    ' a space and followed by the punctuation characters
    ' Next improvement: deal with word preceded by punctuation
    ' and followed by a space
    
    For i = LBound(Word_List) To UBound(Word_List)
        TimesFound = 0
        TempMessageString = "Looking for 'to be' verbs: '" & Word_List(i) & "'"
        StatusBar = ("Performing Grammar Check . . . " & TempMessageString)
        PercentDone = Round(100 * WordsChecked / TotalWordsToCheck, 0)
        WordsChecked = WordsChecked + 1
        PMUserForm.IncrementProgressBar PercentDone, PercentDone & "% Complete " & TempMessageString
        
        For J = 1 To Len(CHARS)
            Search_Word = " " & Word_List(i) & Mid(CHARS, J, 1)
            'If I = 0 And J = 1 Then
            '    Call MsgBox("'" & Search_Word & "'")
            'End If
            TimesFound = TimesFound + Highlight_String(Search_Word, wdYellow)
        Next J
        If TimesFound > 0 Then
           MessageString = MessageString & Now & " Highlighted '" & Word_List(i) & "' " & TimesFound & " times " & vbCr
           ToBeVerbCount = ToBeVerbCount + TimesFound
        End If
        
        ' Check if Abort button pressed:
        If PMUserForm_Abort_Pressed Then GoTo Conclude
        
    Next i '"to be" verbs
    
    ' Treat contractions separately
    For i = LBound(Lone_Words) To UBound(Lone_Words)
        TempMessageString = "Looking for 'to be' verbs: '" & Lone_Words(i) & "'"
        StatusBar = ("Performing Grammar Check . . . " & TempMessageString)
        PercentDone = Round(100 * WordsChecked / TotalWordsToCheck, 0)
        WordsChecked = WordsChecked + 1
        PMUserForm.IncrementProgressBar PercentDone, PercentDone & "% Complete " & TempMessageString
            
        TimesFound = Highlight_String(Lone_Words(i), wdYellow)
        If TimesFound > 0 Then
           MessageString = MessageString & Now & " Highlighted '" & Lone_Words(i) & _
            "' " & TimesFound & " times " & vbCr
            ToBeVerbCount = ToBeVerbCount + TimesFound
        End If
		' Check if Abort button pressed:
        If PMUserForm_Abort_Pressed Then GoTo Conclude
    Next i
    
    ' Hunt for each word in the preposition list, when preceded by
    ' a space and followed by the punctuation characters
    ' Next improvement: deal with word preceded by punctuation
    ' and followed by a space
    
    For i = LBound(Preposition_List) To UBound(Preposition_List)
        TimesFound = 0
        TempMessageString = "Looking for prepositions: '" & Preposition_List(i) & "'"
        StatusBar = ("Performing Grammar Check . . . " & TempMessageString)
        PercentDone = Round(100 * WordsChecked / TotalWordsToCheck, 0)
        WordsChecked = WordsChecked + 1
        PMUserForm.IncrementProgressBar PercentDone, PercentDone & "% Complete " & TempMessageString
            
        For J = 1 To Len(CHARS)
            Search_Word = " " & Preposition_List(i) & Mid(CHARS, J, 1)
            'If I = 0 And J = 1 Then
            '    Call MsgBox("'" & Search_Word & "'")
            'End If
            TimesFound = TimesFound + Highlight_String(Search_Word, wdBrightGreen)
        Next J
        If TimesFound > 0 Then
           MessageString = MessageString & Now & " Highlighted '" & Preposition_List(i) & "' " & TimesFound & " times " & vbCr
           PrepositionCount = PrepositionCount + TimesFound
        End If
		' Check if Abort button pressed:
        If PMUserForm_Abort_Pressed Then GoTo Conclude
    Next i 'Preposition Search
    
    ' Look for expletives
    
    TempMessageString = "Looking for expletives '"
    StatusBar = ("Performing Grammar Check . . . " & TempMessageString)
    PercentDone = Round(100 * WordsChecked / TotalWordsToCheck, 0)
    WordsChecked = WordsChecked + 4
    PMUserForm.IncrementProgressBar PercentDone, PercentDone & "% Complete " & TempMessageString
    
    Call Highlight_String("it is observed that", wdRed, True)
    Call Highlight_String("it was observed that", wdRed, True)
    Call Highlight_String("I think that", wdRed, True)
    Call Highlight_String("we think that", wdRed, True)
    PercentDone = Round(100 * WordsChecked / TotalWordsToCheck, 0)
    WordsChecked = WordsChecked + 4
    PMUserForm.IncrementProgressBar PercentDone, PercentDone & "% Complete " & TempMessageString

    Call Highlight_String("I believe that", wdRed, True)
    Call Highlight_String("we believe that", wdRed, True)
    Call Highlight_String("respectively", wdRed, True)
    Call Highlight_String("based off", wdRed, True)
    
Conclude:
    
    ' Clear Find formatting
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    MessageString = MessageString & Now & " Completed slow starts " & vbCr
    MessageString = MessageString & vbCr & "Yellow highlights indicate weak verbs." & vbCr
    MessageString = MessageString & "Green highlights indicate prepositions." & vbCr
    MessageString = MessageString & "Red highlights indicate expletives." & vbCr
    MessageString = MessageString & "Edit colorful sentences." & vbCr & vbCr
    TotalWordCount = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)
    
    'MessageString = MessageString & "Word found " & TotalWordCount & " words." & vbCr
    'MessageString = MessageString & "ReadabilityStats found " & ActiveDocument.Range.ReadabilityStatistics(1).Value & " words." & vbCr
    'MessageString = MessageString & "ReadabilityStats found " & ActiveDocument.Range.ReadabilityStatistics(8).Value & " passive sentences." & vbCr
    
    MessageString = MessageString & "Weak verb fraction = " & _
        Round(100 * ToBeVerbCount / TotalWordCount, 2) & "%. Strive for < 0.5%." & vbCr
    MessageString = MessageString & "Preposition fraction = " & _
        Round(100 * PrepositionCount / TotalWordCount, 1) & "%. Strive for < 10%." & vbCr
    If PMUserForm_Abort_Pressed Then
        MessageString = MessageString & "*************************************************************************" & vbCr
        MessageString = MessageString & "* ANALYSIS ABORTED, SO THE ABOVE STATISTICS MAY NOT MEAN MUCH.*" & vbCr
        MessageString = MessageString & "*************************************************************************"
    End If
    
    
    ' From Proofreading Collection Object description on
    ' http://msdn.microsoft.com/en-us/library/office/aa223064%28v=office.11%29.aspx
    Set myErrors = ActiveDocument.Range.GrammaticalErrors
    Application.ScreenUpdating = True
    StatusBar = ("Grammar Check COMPLETED")
    
    ' Insert message at end of file
    ActiveDocument.Content.InsertAfter vbCr & vbCr & "MS Word found " & myErrors.Count & _
            " grammar errors. Please check grammar and edit." _
            & vbCr & MessageString
    ActiveDocument.Save
    
    Unload PMUserForm
    
    Call Display_Long_MsgBox("NOTE: New File Name" & vbCr & vbCr & _
        "MS Word found " & myErrors.Count & " grammar errors. Please check grammar and edit." _
        & vbCr & MessageString)

    'MsgBox "NOTE: New File Name" & vbCr & vbCr & _
    '    "MS Word found " & myErrors.Count & " grammar errors. Please check grammar and edit." _
    '    & vbCr & MessageString
    '    & vbCr & "MS Word found " & ActiveDocument.Range.ReadabilityStatistics(8).Value & _
    '         " passive sentences. Please edit."
    StatusBar = ""
    Application.StatusBar = False
 
    
End Sub 'Highlight_Passive_Voice

Public Function Highlight_String(InputText As String, Optional ColorCode As Integer, _
    Optional StrikeText As Boolean)

'
' This subroutine finds the InputText and Highlights it with an
' optional colorcode. It returns the number of times the InputText appears.
'
' To count the number of highlights or strikethroughs, use the technique
' given by Bart Verbeek and Dave Rado, "How to find out, using VBA, how
' many replacements Word made during a Find & Replace All"
' http://word.mvps.org/faqs/MacrosVBA/GetNoOfReplacements.htm

    Dim CurrentHighlightColor As Integer
    Dim TimesFound As Integer
    TimesFound = 0
    Dim MessageString As String, StrFind As String, strReplace As String
    Dim NumCharsBefore As Long, NumCharsAfter As Long
    
    'Application.ScreenUpdating = False
    
    strReplace = "#" & InputText
    
    CurrentHighlightColor = Options.DefaultHighlightColorIndex
    Options.DefaultHighlightColorIndex = ColorCode
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    'ActiveDocument.Range.HighlightColorIndex = ColorCode
    If Not IsMissing(StrikeText) Then
        With Selection.Find.Replacement.Font
            .StrikeThrough = StrikeText
            .DoubleStrikeThrough = False
        End With
    End If
    
    'Get the number of chars in the doc BEFORE doing Find & Replace
    NumCharsBefore = ActiveDocument.Characters.Count
    With Selection.Find
        .Text = InputText
        .Replacement.Text = strReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
        'If .Found Then TimesFound = TimesFound + 1
    End With

    'Get the number of chars AFTER doing Find & Replace
    NumCharsAfter = ActiveDocument.Characters.Count

    'Calculate of the number of replacements,
    'and put the result into the function name variable
    TimesFound = (NumCharsAfter - NumCharsBefore) / _
            (Len(strReplace) - Len(InputText))

    'If the lengths of the find & replace strings were equal at the start, _
    'do another replace to strip out the #
    If TimesFound > 0 Then

        StrFind = strReplace
        'Strip off the hash
        strReplace = Mid$(strReplace, 2)

        With Selection.Find
            .Text = StrFind
            .Replacement.Text = strReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
        End With

    End If
    Options.DefaultHighlightColorIndex = CurrentHighlightColor
    'MessageString = "Highlighted '" & InputText & "' " & TimesFound & " times"
    'If Not IsMissing(StrikeText) Then
    '    If StrikeText Then
    '        MessageString = "Struckout '" & InputText & "' " & TimesFound & " times"
    '    End If
    'End If
    
    '' Application.ScreenUpdating = True
    'If TimesFound > 0 Then
    '    Call MsgBox(MessageString)
    'End If
    Highlight_String = TimesFound
    
End Function 'Highlight_String

Public Sub Display_Long_MsgBox(ByVal InputText As String)
'
' This subroutine displays a string with a series
' of message boxes, if it's too long to fit in one
' message box.
    Dim CurrentLine As Integer
    Dim LinesInMessage As Integer
    Dim Word_List() As String
    Dim PartialMessage As String
    
    Word_List = Split(InputText, vbCr)
    CurrentLine = LBound(Word_List)
    LinesInMessage = UBound(Word_List)
    Do While (LinesInMessage - CurrentLine) > 18
        PartialMessage = ""
        For i = CurrentLine To (CurrentLine + 18)
            PartialMessage = PartialMessage & Word_List(i) & vbCr
        Next i
        MsgBox PartialMessage
        CurrentLine = CurrentLine + 19
    Loop
    ' display remainder now
    PartialMessage = ""
    For i = CurrentLine To LinesInMessage
        PartialMessage = PartialMessage & Word_List(i) & vbCr
    Next i
    MsgBox PartialMessage
    
End Sub 'Display_Long_MsgBox

Public Function Split(ByVal InputText As String, _
         Optional ByVal Delimiter As String) As Variant

' This function comes from:
' http://msdn.microsoft.com/en-us/library/office/aa155763%28v=office.10%29.aspx

' This function splits the sentence in InputText into
' words and returns a string array of the words. Each
' element of the array contains one word.

    ' This constant contains punctuation and characters
    ' that should be filtered from the input string.
    Const CHARS = ".!?,;:""'()[]{}"
    Dim strReplacedText As String
    Dim intIndex As Integer

    ' Replace tab characters with space characters.
    strReplacedText = Trim(Replace(InputText, _
         vbTab, " "))

    ' Filter all specified characters from the string.
    For intIndex = 1 To Len(CHARS)
        strReplacedText = Trim(Replace(strReplacedText, _
            Mid(CHARS, intIndex, 1), " "))
    Next intIndex

    ' Loop until all consecutive space characters are
    ' replaced by a single space character.
    Do While InStr(strReplacedText, "  ")
        strReplacedText = Replace(strReplacedText, _
            "  ", " ")
    Loop

    ' Split the sentence into an array of words and return
    ' the array. If a delimiter is specified, use it.
    'MsgBox "String:" & strReplacedText
    If Len(Delimiter) = 0 Then
        Split = VBA.Split(strReplacedText)
    Else
        Split = VBA.Split(strReplacedText, Delimiter)
    End If
End Function 'Split

Public Sub SaveWithNewFileName()
'
' This subroutine saves the file with a new file name
' to indicate it has had highlights added
'
' The method comes from:
' http://msdn.microsoft.com/en-us/library/office/aa220734(v=office.11).aspx
'
    Dim strDocName As String
    Dim intPos As Integer

    'Find position of extension in filename
    strDocName = ActiveDocument.Name
    strDocExtension = ".doc" ' Default extension
    intPos = InStrRev(strDocName, ".")

    If intPos = 0 Then

        'If the document has not yet been saved
        'Ask the user to provide a filename
        strDocName = InputBox("Please enter the name " & _
            "of your document.")
    Else

        'Strip off extension and add "Highlighted" to file name
        strDocExtension = Right(strDocName, (Len(strDocName) - intPos + 1))
        strDocName = Left(strDocName, intPos - 1)
        strDocName = strDocName & "Highlighted" & strDocExtension
    End If

    'Save file with new extension
    ActiveDocument.SaveAs FileName:=strDocName

End Sub

