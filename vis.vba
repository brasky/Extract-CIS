Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Public CurrentXLRow As Integer
Public CurrentControl As String
Public ControlText As String
Dim sourceDoc As Object
Dim targetDoc As Object
Dim name As String
Dim xFind As String
Dim xReplace As String
Dim toFind As String
Dim toReplace As String
Dim nameList() As String
Dim findList() As String
Dim replaceList() As String
Dim nameResult() As Boolean
Dim find_and_replace_result As Boolean
Dim controlArray() As String
Dim xlApp As Object
Dim xlBook As Object
Dim book As Object
Dim myPlaceholderText As String
Dim cc As ContentControl

Sub extract_SSP_answers()

    Set sourceDoc = ActiveDocument
    
    Set xlApp = GetObject(, "Excel.Application")
    xlApp.workbooks.Add
    Set targetDoc = xlApp.ActiveWorkbook
    myRow = 5
    
    sourceDoc.Activate
    Selection.Start = 0
    Selection.Collapse
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "is the solution and how is it"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute

    Do Until Selection.Find.Found = False
    
        ControlText = Empty
        
        Selection.Expand unit:=wdCell
        CurrentControl = Trim(Left(Selection.Text, InStr(1, Selection.Text, "What") - 1))
'need to format currentControl here
        Selection.Expand unit:=wdTable
        myRowCount = Selection.Tables(1).Rows.Count
        
        If myRowCount > 2 Then
        
            For i = 2 To myRowCount
            
                If i <> myRowCount Then seperator = Chr(13) Else seperator = ""
                Selection.Tables(1).Rows(i).Select
                ControlText = ControlText & Selection.Text & seperator
            
            Next
            
        Else 'if myRowCount = 2
        
            If myRowCount <> 2 Then Stop 'myRowCount issue
            Selection.Tables(1).Rows(2).Select
            ControlText = Selection.Text
        
        End If
        
        xlApp.Application.Range("A" & myRow).Value = CurrentControl
        xlApp.Application.Range("B" & myRow).Value = ControlText
        myRow = myRow + 1
        Selection.Collapse
        Selection.Find.Execute
       
       
    Loop

End Sub

Sub Export_O365_SSP_to_Excel()

    xlRow = 3
    
    Set sourceDoc = ActiveDocument
    
    Set xlApp = GetObject(, "Excel.Application")
    For Each book In xlApp.workbooks
        If InStr(1, book.name, "target", vbTextCompare) <> 0 Then Set targetDoc = book
    Next

    If targetDoc Is Nothing Then Stop 'couldn't set targetdoc. Is excel open??
    
    Selection.Start = 0
    Selection.Collapse
    sourceDoc.Activate
    Selection.Start = 0
    Selection.Collapse
   
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Control Summary Information"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
   
    Selection.Find.Execute

'this is the beginning of the loop to search the document
Do Until Selection.Find.Found = False

  'get the control summary information

   'find the bounds of the table
    Selection.Expand unit:=wdTable
    tableStart = Selection.Start
    tableEnd = Selection.End
    
   'count the parameters in the source doc
    i = 1
    parameterCount = 0
    Do Until InStr(i, Selection.Text, "Parameter") = 0
        parameterCount = parameterCount + 1
        Let i = InStr(i, Selection.Text, "Parameter") + 2
    Loop
    
   'get the control number
    Selection.Collapse
    Selection.Expand unit:=wdCell
    CurrentControl = CleanTrim(CStr(Selection.Text))
    Selection.Collapse
    
   'Get the Responsible Role
    Selection.Find.Text = "Responsible Role:"
    Selection.Find.Execute
    Selection.Expand unit:=wdCell
    myResponsibleRole = Trim(Replace(CleanTrim(Selection.Text), "Responsible Role:", ""))
    
   'look for parameters and store them
    Selection.Move unit:=wdCell
    Selection.Expand unit:=wdCell
    If parameterCount = 0 Then parameterCount = 1
    ReDim myParameters(1 To parameterCount) As String
    currParam = 1
    
    Do Until InStr(1, Selection.Text, "(check all that apply)") <> 0
    
        If InStr(1, Selection.Text, "parameter", vbTextCompare) = 0 Then MsgBox ("not a parameter, not checkboxes")
        If InStr(1, Selection.Text, "parameter", vbTextCompare) = 0 Then Stop
        myParameters(currParam) = Trim(LessCleanTrim(Selection.Text))
        myParameters(currParam) = Replace(myParameters(currParam), Chr(10), "") 'remove LF from end of string
        myParameters(currParam) = Trim(Replace(myParameters(currParam), Chr(13), " ")) & ";" 'remove CR from end of string
        currParam = currParam + 1
        Selection.Move unit:=wdCell
        Selection.Expand unit:=wdCell
        
    Loop
    
    Let parameters = Empty
    For i = 1 To UBound(myParameters())
        parameters = parameters & myParameters(i) & Chr(10)
    Next
    parameters = Left(parameters, Len(parameters) - 1)
    
    'capture the Implementation Status checkbox data
    ReDim impStatBoxes(1 To 5) As Boolean
    For Each Box In Selection.FormFields
        If Box.Result = 1 Then
            Box.Range.Select
            Selection.EndKey Extend:=False
            Selection.Start = Box.Range.Start
            resultText = CleanTrim(Selection.Text)
            statusCheckCount = statusCheckCount + 1
            If resultText = "Implemented" Or resultText = "implemented" Then impStatBoxes(1) = True
            If InStr(1, resultText, "Partially Implemented", vbTextCompare) <> 0 Then impStatBoxes(2) = True
            If InStr(1, resultText, "Planned", vbTextCompare) <> 0 Then impStatBoxes(3) = True
            If InStr(1, resultText, "Alternative Implementation", vbTextCompare) <> 0 Then impStatBoxes(4) = True
            If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then impStatBoxes(5) = True
        End If
        Selection.Expand unit:=wdCell
    Next
    
    Selection.Expand unit:=wdTable
    Selection.Range.Tables(1).Rows(Selection.Range.Tables(1).Rows.Count).Select
    
    ReDim contOrigBoxes(1 To 8) As Boolean
    For Each Box In Selection.FormFields
        If Box.Result = 1 Then
            Box.Range.Select
            Selection.EndKey Extend:=False
            Selection.Start = Box.Range.Start
            resultText = CleanTrim(Selection.Text)
            originationCheckCount = originationCheckCount + 1
            If InStr(1, resultText, "Provider Corporate", vbTextCompare) <> 0 Then contOrigBoxes(1) = True
            If InStr(1, resultText, "Provider System", vbTextCompare) <> 0 Then contOrigBoxes(2) = True
            If InStr(1, resultText, "Provider Hybrid", vbTextCompare) <> 0 Then contOrigBoxes(3) = True
            If InStr(1, resultText, "Configured by Cust", vbTextCompare) <> 0 Then contOrigBoxes(4) = True
            If InStr(1, resultText, "Provided by Cust", vbTextCompare) <> 0 Then contOrigBoxes(5) = True
            If InStr(1, resultText, "Shared", vbTextCompare) <> 0 Then contOrigBoxes(6) = True
            If InStr(1, resultText, "Inherited", vbTextCompare) <> 0 Then contOrigBoxes(7) = True
            If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then contOrigBoxes(8) = True

        End If
        Selection.Expand unit:=wdCell
    Next
    
   'find the implementation table.
    Selection.Collapse
    Selection.Find.Text = "What is the solution and how is it implemented?"
    Selection.Find.Execute
    If Selection.Find.Found = False Then Stop 'didn't find the implementation table!!!
    Selection.Expand unit:=wdTable
    If InStr(1, Selection.Text, CurrentControl, vbTextCompare) = 0 Then Stop 'control number doesn't match perfectly!!!
    
   'if there's no part a, b, c, etc... grab the answer and move on
    If Selection.Tables(1).Columns.Count = 1 And Selection.Cells.Count = 2 Then
    
        answerRows = 1
        answerCols = 1
        Selection.Cells(2).Select
        Selection.MoveEnd unit:=wdCharacter, Count:=-1
        ReDim impText(2, 1) As String
        impText(1, 1) = ""
        impText(2, 1) = Selection.Text
    
    End If
    
   'if there are parts, grab the answers into an array, then move on
    If Selection.Tables(1).Columns.Count = 2 Then
    
        answerRows = Selection.Tables(1).Rows.Count - 1
        answerCols = 2
        ReDim impText(2, 1 To answerRows) As String
        i = 2
        j = 1
        Do Until i >= Selection.Cells.Count
            impText(1, j) = "(" & CleanTrim(Replace(Selection.Cells(i).Range.Text, "part", "", , , vbTextCompare)) & ")"
            i = i + 1
            impText(2, j) = Selection.Cells(i).Range.Text
            i = i + 1
            j = j + 1
        Loop
    
    End If
    
    'put the answers in the spreadsheet.
    For i = 1 To answerRows
        xlApp.Range("A" & xlRow).Value = CurrentControl & impText(1, i)
        xlApp.Range("B" & xlRow).Value = myResponsibleRole
        xlApp.Range("C" & xlRow).Value = parameters
        xlApp.Range("D" & xlRow).Value = impStatBoxes(1)
        xlApp.Range("E" & xlRow).Value = impStatBoxes(2)
        xlApp.Range("F" & xlRow).Value = impStatBoxes(3)
        xlApp.Range("G" & xlRow).Value = impStatBoxes(4)
        xlApp.Range("H" & xlRow).Value = impStatBoxes(5)
        xlApp.Range("I" & xlRow).Value = contOrigBoxes(1)
        xlApp.Range("J" & xlRow).Value = contOrigBoxes(2)
        xlApp.Range("K" & xlRow).Value = contOrigBoxes(3)
        xlApp.Range("L" & xlRow).Value = contOrigBoxes(4)
        xlApp.Range("M" & xlRow).Value = contOrigBoxes(5)
        xlApp.Range("N" & xlRow).Value = contOrigBoxes(6)
        xlApp.Range("O" & xlRow).Value = contOrigBoxes(7)
        xlApp.Range("P" & xlRow).Value = contOrigBoxes(8)
        xlApp.Range("Q" & xlRow).Value = impText(2, i)
        xlRow = xlRow + 1
    Next
    
    'reset the Find method to look for "Control Summary Information" instead of "What is the Solu...."
    sourceDoc.Activate
    Selection.Collapse
    Selection.Find.Text = "Control Summary Information"
    Selection.Find.Execute
   
Loop

Stop 'end of script

End Sub
Function CleanTrim(ByVal S As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
  Dim X As Long, CodesToClean As Variant
  CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
  If ConvertNonBreakingSpace Then S = Replace(S, Chr(160), " ")
  For X = LBound(CodesToClean) To UBound(CodesToClean)
    If InStr(S, Chr(CodesToClean(X))) Then S = Replace(S, Chr(CodesToClean(X)), "")
  Next
  CleanTrim = Trim(S)
End Function
Function LessCleanTrim(ByVal S As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
  Dim X As Long, CodesToClean As Variant
  CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
  If ConvertNonBreakingSpace Then S = Replace(S, Chr(160), " ")
  For X = LBound(CodesToClean) To UBound(CodesToClean)
    If InStr(S, Chr(CodesToClean(X))) Then S = Replace(S, Chr(CodesToClean(X)), "")
  Next
  LessCleanTrim = Trim(S)
End Function
Function EndOfStringCRLF(ByVal S As String)
  E = 7 'number of characters at the end of the string to clean
  NE = E + 1
  S = Left(S, Len(S) - NE) & Replace(S, Chr(10), "", Len(S) - E)
  S = Left(S, Len(S) - NE) & Replace(S, Chr(13), "", Len(S) - E)
  EndOfStringCRLF = S
End Function

