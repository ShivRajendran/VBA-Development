Sub colorcheckLp()
    'Debug.Print Worksheets(1).Range("K6").Interior.Color
    'Debug.Print Worksheets(1).Range("C44").Interior.Color
    trow1 = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    trow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    Sheets.Add after:=Worksheets(2)
    Worksheets(3).Name = "Retention Status LP"
    Z = 2
    For i = 2 To trow1
        firtinst = False
        If Worksheets(1).Range("C" & i).Interior.Color = Worksheets(1).Range("K8").Interior.Color Then
            asid = Worksheets(1).Range("A" & i).Value
            For j = 3 To trow2
                If (Replace(Worksheets(2).Range("B" & j).Value, " ", "") = asid) Then
                    firstinst = True
                    Worksheets(3).Range("A" & Z).Value = asid
                    Worksheets(3).Range("B" & Z).Value = Worksheets(1).Range("B" & i).Value
                    Worksheets(3).Range("C" & Z).Value = Worksheets(1).Range("C" & i).Value
                    Worksheets(3).Range("D" & Z).Value = "retained Low Plan"
                    Z = Z + 1
                    GoTo cont
                End If
            Next
                If firstint = False Then
                    Worksheets(3).Range("A" & Z).Value = asid
                    Worksheets(3).Range("B" & Z).Value = Worksheets(1).Range("B" & i).Value
                    Worksheets(3).Range("C" & Z).Value = Worksheets(1).Range("C" & i).Value
                    Worksheets(3).Range("D" & Z).Value = "Did not retain Low Plan"
                    Z = Z + 1
            End If
        End If
cont:
    Next
     Worksheets(3).Range("A1").Value = "Associate ID"
     Worksheets(3).Range("A1").Font.Bold = True
     Worksheets(3).Range("B1").Value = "First Name"
     Worksheets(3).Range("B1").Font.Bold = True
     Worksheets(3).Range("C1").Value = "Last Name"
     Worksheets(3).Range("C1").Font.Bold = True
     Worksheets(3).Range("D1").Value = "Decision to opt out"
     Worksheets(3).Range("D1").Font.Bold = True
     Worksheets(3).Range("A1").Interior.Color = vbGreen
     Worksheets(3).Range("B1").Interior.Color = vbGreen
     Worksheets(3).Range("C1").Interior.Color = vbGreen
     Worksheets(3).Range("D1").Interior.Color = vbGreen
     
     trow3 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
     For i = 2 To trow3
        If Worksheets(3).Range("D" & i).Value = "retained Low Plan" Then
            Worksheets(3).Range("D" & i).Interior.Color = vbYellow
        Else
            Worksheets(3).Range("D" & i).Interior.Color = vbRed
        End If
     Next
     Worksheets(3).Columns.AutoFit
     
     Worksheets(3).Buttons.Add(812.25, 114.75, 267.75, 82.5).Select
    Selection.Characters.Text = "Click me for a count report!"
    Selection.OnAction = "reportL"
    Worksheets(3).Range("A1").Select
End Sub

Sub reportL()
    trow3 = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    c1 = 0
    c2 = 0
    
    For i = 2 To trow3
        If ActiveSheet.Range("D" & i).Value = "retained Low Plan" Then
            c1 = c1 + 1
        Else
            c2 = c2 + 1
        End If
     Next
     
     MsgBox c1 & " People opted in while " & c2 & " people opted out, out of " & c1 + c2 & " total employees", vbOKOnly, "report"
End Sub
