Dim xlApp As Object
Dim xlBook As Object
Dim book As Object
Dim targetDoc As Object
Dim sourceDoc As Object
Dim paraStart As Long

Sub copy_checkbox_data_from_SSP_to_CIS_workbook()

    bailout = MsgBox("Don't forget to have the CIS document open in Excel before you click OK", vbOKCancel)
    If bailout = vbCancel Then Exit Sub
    
    Set sourceDoc = ActiveDocument
    Set xlApp = GetObject(, "Excel.Application")
    
    For Each book In xlApp.workbooks
        If InStr(1, book.Name, "CIS", vbTextCompare) <> 0 Then Set targetDoc = book
    Next

    If targetDoc Is Nothing Then
        a = MsgBox("Didn't find a CIS workbook. The file should have CIS in the name. Please open the workbook and try again.", , "FATAL ERROR")
        Exit Sub
    End If
    
    targetDoc.Activate
    xlApp.Application.worksheets("CIS").Activate
    
    sourceDoc.Activate
    Selection.Start = 0
    Selection.Collapse
    Selection.Find.ClearFormatting
    Selection.Find.Text = "Control Summary Information"
    Selection.Find.Execute
    
    statusErrorList = Empty
    originationErrorList = Empty
    
    Do Until Selection.Find.Found = False
    
        statusCheckCount = 0
        originationCheckCount = 0
        impl = False
        part = False
        plan = False
        alte = False
        NAI = False
        
        corp = False
        syst = False
        hybr = False
        conf = False
        prov = False
        shar = False
        inhe = False
        NAC = False
                
        Selection.Expand unit:=wdTable
        Selection.Collapse
        Selection.Expand unit:=wdCell
        currentControl = CleanTrim(Selection.Text)
        Selection.Expand unit:=wdTable
        tableRows = Selection.Range.Tables(1).Rows.Count
        Selection.Range.Tables(1).Rows(tableRows - 1).Select
        
        If Selection.FormFields.Count = 0 Then
        
            For Each Box In Selection.ContentControls
                If Box.Checked = True Then
                    Box.Range.Select
                    Selection.EndKey Extend:=False
                    Selection.Start = Box.Range.Start
                    resultText = CleanTrim(Selection.Text)
                    statusCheckCount = statusCheckCount + 1
                    If InStr(1, resultText, "Implemented", vbTextCompare) <> 0 Then impl = True
                    If InStr(1, resultText, "Partially Implemented", vbTextCompare) <> 0 Then part = True
                    If InStr(1, resultText, "Partially Implemented", vbTextCompare) <> 0 Then impl = False
                    If InStr(1, resultText, "Planned", vbTextCompare) <> 0 Then plan = True
                    If InStr(1, resultText, "Alternative Implementation", vbTextCompare) <> 0 Then alte = True
                    If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then NAI = True
                End If
                Selection.Expand unit:=wdCell
            Next
            
            Selection.Expand unit:=wdTable
            Selection.Range.Tables(1).Rows(tableRows).Select
            
            For Each Box In Selection.ContentControls
              If Box.Type = 8 Then
                If Box.Checked = True Then
                    Box.Range.Select
                    Selection.EndKey Extend:=False
                    Selection.Start = Box.Range.Start
                    resultText = CleanTrim(Selection.Text)
                    originationCheckCount = originationCheckCount + 1
                    If InStr(1, resultText, "Provider Corporate", vbTextCompare) <> 0 Then corp = True
                    If InStr(1, resultText, "Provider System", vbTextCompare) <> 0 Then syst = True
                    If InStr(1, resultText, "Provider Hybrid", vbTextCompare) <> 0 Then hybr = True
                    If InStr(1, resultText, "Configured by Cust", vbTextCompare) <> 0 Then conf = True
                    If InStr(1, resultText, "Provided by Cust", vbTextCompare) <> 0 Then prov = True
                    If InStr(1, resultText, "Shared", vbTextCompare) <> 0 Then shar = True
                    If InStr(1, resultText, "Inherited", vbTextCompare) <> 0 Then inhe = True
                    If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then NAC = True
    
                End If
                Selection.Expand unit:=wdCell
              End If
            
            Next
            
        Else
        
            For Each Box In Selection.FormFields
                If Box.Result = 1 Then
                    Box.Range.Select
                    Selection.EndKey Extend:=False
                    Selection.Start = Box.Range.Start
                    resultText = CleanTrim(Selection.Text)
                    statusCheckCount = statusCheckCount + 1
                    If resultText = "Implemented" Or resultText = "implemented" Then impl = True
                    If InStr(1, resultText, "Partially Implemented", vbTextCompare) <> 0 Then part = True
                    If InStr(1, resultText, "Planned", vbTextCompare) <> 0 Then plan = True
                    If InStr(1, resultText, "Alternative Implementation", vbTextCompare) <> 0 Then alte = True
                    If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then NAI = True
                End If
                Selection.Expand unit:=wdCell
            Next
            
            Selection.Expand unit:=wdTable
            Selection.Range.Tables(1).Rows(tableRows).Select
            
            For Each Box In Selection.FormFields
                If Box.Result = 1 Then
                    Box.Range.Select
                    Selection.EndKey Extend:=False
                    Selection.Start = Box.Range.Start
                    resultText = CleanTrim(Selection.Text)
                    originationCheckCount = originationCheckCount + 1
                    If InStr(1, resultText, "Provider Corporate", vbTextCompare) <> 0 Then corp = True
                    If InStr(1, resultText, "Provider System", vbTextCompare) <> 0 Then syst = True
                    If InStr(1, resultText, "Provider Hybrid", vbTextCompare) <> 0 Then hybr = True
                    If InStr(1, resultText, "Configured by Cust", vbTextCompare) <> 0 Then conf = True
                    If InStr(1, resultText, "Provided by Cust", vbTextCompare) <> 0 Then prov = True
                    If InStr(1, resultText, "Shared", vbTextCompare) <> 0 Then shar = True
                    If InStr(1, resultText, "Inherited", vbTextCompare) <> 0 Then inhe = True
                    If InStr(1, resultText, "Not Applicable", vbTextCompare) <> 0 Then NAC = True
    
                End If
                Selection.Expand unit:=wdCell
            Next
        
        End If
        
        targetDoc.Activate

        If InStr(1, currentControl, "(") <> 0 Then 'positive test here means its a control enhancement
            preamble = Left(currentControl, 3)
            enhancement = Trim(Mid(currentControl, InStr(1, currentControl, "(")))
            controlNumber = Trim(Mid(currentControl, 4, InStr(1, currentControl, "(") - 4))
            If Len(controlNumber) = 1 Then controlNumber = "0" & controlNumber
            If Len(enhancement) = 3 Then enhancement = "(0" & Right(enhancement, 2)
            currentControl = preamble & controlNumber & " " & enhancement
        Else 'this would be a base control.
            If Len(currentControl) = 4 Then currentControl = Left(currentControl, 3) & "0" & Right(currentControl, 1)
        End If
        Set c = Nothing
        Set c = xlApp.Application.Range("b:b").Find(currentControl)
        
        If c Is Nothing Then
            'add the control to the bottom of the list
            For Each Cell In xlApp.Application.Range("B350:B1000")
                If Cell.Value = "" Then Exit For
            Next
            Let Cell.Value = currentControl
            'do the find operation again
            Set c = xlApp.Application.Range("b:b").Find(currentControl)
        End If
        
        If impl = True Then c.Offset(0, 1).Value = "x"
        If part = True Then c.Offset(0, 2).Value = "x"
        If plan = True Then c.Offset(0, 3).Value = "x"
        If alte = True Then c.Offset(0, 4).Value = "x"
        If NAI = True Then c.Offset(0, 5).Value = "x"
        
        X = 0
        For Each Cell In xlApp.Application.Range("c" & c.Row & ":g" & c.Row)
            If Cell.Value = "x" Then X = X + 1
        Next
        If X <> statusCheckCount Then statusErrorList = statusErrorList & "; " & currentControl 'status check count doesn't match
        
        If corp = True Then c.Offset(0, 6).Value = "x"
        If syst = True Then c.Offset(0, 7).Value = "x"
        If hybr = True Then c.Offset(0, 8).Value = "x"
        If conf = True Then c.Offset(0, 9).Value = "x"
        If prov = True Then c.Offset(0, 10).Value = "x"
        If shar = True Then c.Offset(0, 11).Value = "x"
        If inhe = True Then c.Offset(0, 12).Value = "x"
        If NAC = True Then c.Offset(0, 13).Value = "x"
        
        X = 0
        For Each Cell In xlApp.Application.Range("h" & c.Row & ":o" & c.Row)
            If Cell.Value = "x" Then X = X + 1
        Next
        If X <> originationCheckCount Then originationErrorList = originationErrorList & "; " & currentControl 'origination check count doesn't match
        
        sourceDoc.Activate
        Selection.Collapse
        Selection.Find.Execute
        
    Loop
    
    xlApp.Application.Range("O3").Value = "N/A"
    xlApp.Application.Range("P4").Value = "=COUNTIF(C4:G4," & Chr(34) & "x" & Chr(34) & ")"
    xlApp.Application.Range("Q4").Value = "=COUNTIF(H4:O4," & Chr(34) & "x" & Chr(34) & ")"
    xlApp.Application.Range("R4").Value = "=IF(OR(P4=0,Q4=0)," & Chr(34) & "ERROR" & Chr(34) & "," & Chr(34) & Chr(34) & ")"
    xlApp.Application.Range("P4:R800").Select
    xlApp.Application.Selection.FillDown
    
    If statusErrorList <> Empty Then statusErrorList = Mid(statusErrorList, 3)
    If originationErrorList <> Empty Then originationErrorList = Mid(originationErrorList, 3)
    If statusErrorList <> Empty Or originationErrorList <> Empty Then MsgBox "Issues were found with the checkboxes on the following controls; they should be reviewed manually:" & Chr(10) & Chr(10) & "Implementation Status: " & statusErrorList & Chr(10) & Chr(10) & "Control Origination: " & originationErrorList
    
    MsgBox "Checkbox operation complete. Don't forget to look for rows without checkboxes (see column R) and run the customer responsibility script."

End Sub
Sub copy_customer_responsibility_from_SSP_to_CIS_workbook()
'
'
'
    
   'set the current document as the source
    Set sourceDoc = ActiveDocument
    
   'this section finds and establishes the spreadsheet for putting the data in to
    Set xlApp = GetObject(, "Excel.Application")
    
    For Each book In xlApp.workbooks
        If InStr(1, book.Name, "CIS", vbTextCompare) <> 0 Then Set targetDoc = book
    Next

    If targetDoc Is Nothing Then
        a = MsgBox("Didn't find a CIS workbook. The file should have CIS in the name. Please open the workbook and try again.", , "FATAL ERROR")
        Exit Sub
    End If
    
    targetDoc.Activate
    xlApp.Application.worksheets("Customer Responsibility Matrix").Activate
    pasteRow = 4 'this sets the first spreadsheet row to paste in to
    
   'this sets the cursor to the start and establishes the find exercise
    sourceDoc.Activate
    Selection.Start = 0
    Selection.Collapse
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    With Selection.Find
        .Text = "Customer Responsibility"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
   'beginning of loop
    Do Until Selection.Find.Found = False Or firstFound = Selection.Start
       
       'this section finds the customer responsibility text and loads it in to 'custRespText
        custResp1 = Selection.Start
        custResp2 = Selection.End + 1
        Selection.Collapse
        Selection.MoveDown
        paraStart = Selection.Start
        
        Selection.Expand unit:=wdCell
        cellStart = Selection.Start
        cellEnd = Selection.End
        Selection.Collapse
        
        Selection.MoveDown
        Do Until Selection.Start > cellEnd
            Selection.MoveDown
            If Selection.Font.Bold = True Then
                Selection.MoveUp
                Selection.Expand unit:=wdLine 'extend the selection to the end of the row and see if its blank or not
                If Selection.End - Selection.Start <= 2 Then Selection.MoveUp
                Selection.Expand unit:=wdLine
                If Selection.End - Selection.Start <= 2 Then Selection.MoveUp
                Selection.Expand unit:=wdLine
                paraEnd = Selection.End
                Selection.Start = paraStart
                GoTo exitLoop1
            End If
        Loop
exitLoop1:
        If Selection.Start > cellEnd Then
            Selection.Start = paraStart
            Selection.End = cellEnd
        End If
        
        custRespText = Selection.Text
        Do Until Right(custRespText, 1) <> Chr(13) And Right(custRespText, 1) <> Chr(10)
            If Right(custRespText, 1) = Chr(13) Or Right(custRespText, 1) = Chr(10) Then custRespText = Left(custRespText, Len(custRespText) - 1)
        Loop
        
       'this section gets the control number, and control part if applicable (a,b,c,etc...)
        Selection.Move unit:=wdCell, Count:=-2
        Selection.Expand unit:=wdCell
        
        If Len(Selection.Text) < 10 Then
            control_part = Trim(Replace(Selection.Text, "Part ", "("))
            control_part = Trim(Replace(control_part, Chr(7), ""))
            control_part = Trim(Replace(control_part, Chr(13), ""))
            control_part = Trim(Replace(control_part, "  ", " "))
            control_part = Trim(Replace(control_part, "  ", " "))
            If Left(control_part, 1) = "(" Then control_part = control_part & ")"
            If Left(control_part, 3) = "Req" Then control_part = " " & control_part
        Else
            control_part = Empty
        End If
        
        Selection.Expand unit:=wdTable
        Selection.Collapse
        Selection.Expand unit:=wdCell
        controlName = Trim(Replace(Selection.Text, "- What is the solution and how is it implemented?", " "))
        controlName = Trim(Replace(controlName, Chr(160), Chr(32)))
        controlName = Trim(Replace(controlName, "What is the solution and how is it implemented?", " "))
        controlName = Trim(Replace(controlName, Chr(7), ""))
        controlName = Trim(Replace(controlName, Chr(10), ""))
        controlName = Trim(Replace(controlName, Chr(13), ""))
        controlName = Trim(Replace(controlName, "  ", " "))
        controlName = Trim(Replace(controlName, "  ", " "))
        controlName = Trim(Replace(controlName, " ", ""))
        If Left(control_part, 5) <> Left(controlName, 5) Then controlName = Trim(controlName) & control_part
        controlName = Trim(Replace(controlName, Chr(7), ""))
        controlName = Trim(Replace(controlName, Chr(10), ""))
        controlName = Trim(Replace(controlName, Chr(13), ""))
        controlName = Trim(Replace(controlName, "  ", " "))
        controlName = Trim(Replace(controlName, "  ", " "))
        firstFound = custResp1
        custResp1 = custResp2
        
       'this section places the results in the spreadsheet
        Let xlApp.Application.Range("A" & pasteRow).Value = pasteRow - 3
        Let xlApp.Application.Range("B" & pasteRow).Value = custRespText
        Let xlApp.Application.Range("C" & pasteRow).Value = controlName
        pasteRow = pasteRow + 1
    
       'this section looks for the next match before restarting the loop
        sourceDoc.Activate
        Selection.Start = custResp2 + 5
        Selection.Collapse
        Selection.Find.Execute
    
    Loop
    
    MsgBox "Customer responsibility operation complete. Don't forget to run the checkbox script too"
    
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

