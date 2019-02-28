Sub create()
Dim urows As Long, urows2 As Long

Range("R1").EntireColumn.Insert
Range("R2").Value = "Cross Check"
ActiveWorkbook.Worksheets(2).Activate
urows = Cells(Rows.Count, 1).End(xlUp).Row
urows = urows / 8
ActiveWorkbook.Worksheets(1).Activate
urows2 = Cells(Rows.Count, 1).End(xlUp).Row


Dim y As Long

For i = 3 To urows2 - 2
    x = Range("A" & i).Value
    If Worksheets(2).Range("A:A").Find(x) Is Nothing Then
        Range("Q" & i).Offset(, 1).Value = "NA:Person not found"
        Range("Q" & i).Offset(, 1).Interior.ColorIndex = 8
        GoTo continue2
    End If
    
    y = Worksheets(2).Range("A:A").Find(x).Row
    cvalarr = Split(Worksheets(1).Range("Q" & i).Value, " ", -1)
    cval = CInt(cvalarr(0))
    If cvalarr(3) = "Year" Then
        cval = ((cval / 3) / 4)
    End If
    zcount = 0
    jj = 0
    ytot = 0
    For j = 0 To 10
        
        If j = 0 Then
            Do While x = Worksheets(2).Range("A" & (y + jj)).Value
                jj = jj + 1
            Loop
        End If
        
        If x <> Worksheets(2).Range("A" & (y + j)).Value Then GoTo continue2
        
        dval = Worksheets(2).Range("K" & (y + j)).Value
        If dval = 0 Then zcount = zcount + 1
        If dval = 0 And zcount = jj And cval > 0 Then
            Range("Q" & i).Offset(, 1).Value = "Errors occurred"
            Range("Q" & i).Offset(, 1).Interior.ColorIndex = 3
            GoTo continue2
        End If
        If dval = 0 And zcount = jj Then
            Range("Q" & i).Offset(, 1).Value = "No errors"
            Range("Q" & i).Offset(, 1).Interior.ColorIndex = 6
            GoTo continue
        End If
        If dval = 0 Then
            GoTo continue
        End If
        
        If dval < cval + 5 And dval > cval - 5 Then
            Range("Q" & i).Offset(, 1).Value = "No errors"
            Range("Q" & i).Offset(, 1).Interior.ColorIndex = 6
        Else
            Range("Q" & i).Offset(, 1).Value = "Errors occurred"
            Range("Q" & i).Offset(, 1).Interior.ColorIndex = 3
            GoTo continue2
        End If
continue:
     Next
continue2:
Next

Worksheets(1).Range("R:R").Select



End Sub
Sub look()
Worksheets(1).Range("S1").EntireColumn.Insert
Worksheets(1).Range("S2").Value = "Reason for Error"

s1rowcount = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
s2rowcount = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To s1rowcount - 2
    If Worksheets(1).Range("R" & i).Value = "No errors" Or Worksheets(1).Range("R" & i).Value = "NA:Person not found" Then GoTo continue4
    y = Worksheets(1).Range("A" & i).Value
    dedvalarr = Split(Worksheets(1).Range("Q" & i).Value, " ", -1)
    dedval = CInt(dedvalarr(0))
    If dedvalarr(3) = "Year" Then
        dedval = ((dedval / 3) / 4)
    End If
    
    Count = 0
    focc = 0
    For ii = 1 To s2rowcount
        If focc <> 0 And (Worksheets(2).Range("A" & ii)) <> y Then GoTo continue3
        If focc = 0 And (Worksheets(2).Range("A" & ii)) = y Then
            focc = ii
        End If
        If (Worksheets(2).Range("A" & ii)) = y Then
            Count = Count + 1
        End If
    Next
continue3:
    Worksheets(2).Activate
    'Dim rng As Range

    'rng = Worksheets(2).Range("a" & focc & ":k" & ((focc + Count) - 1))
    errormsg = ""
    'rng(1, 1).Select
    numzero = 0
    For j = focc To ((focc + Count) - 1)
        If Worksheets(2).Cells(j, 11).Value = 0 Then
            numzero = numzero + 1
        End If
        If Worksheets(2).Cells(j, 11).Value <> dedval And Worksheets(2).Cells(j, 11).Value <> 0 And dedval <> 0 Then
            errormsg = errormsg & Worksheets(2).Cells(j, 10).Value & "," & CStr(Worksheets(2).Cells(j, 11).Value) & " "
        End If
    Next
    If numzero = Count Then errormsg = "No deductions were made!"
    Worksheets(1).Range("S" & i).Value = errormsg
continue4:
Next

End Sub
