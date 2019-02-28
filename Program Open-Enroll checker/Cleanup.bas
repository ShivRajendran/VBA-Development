Sub RRforCross()

    Sheets("MOO data").Select
    Sheets("MOO data").Move Before:=Sheets(1)
    Sheets("LP data").Select
    Sheets("LP data").Move Before:=Sheets(2)
    Range("AB14").Select
    Sheets("HP data").Select
    Sheets("Hp data").Move Before:=Sheets(3)
    
End Sub

Sub cross()
    
    trow1 = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row - 2
    trow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row - 2
    trow3 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row - 2
    trk = 3
    Worksheets.Add after:=Worksheets(3)
    
    For i = 3 To trow1
        x = True
        y = False
        Z = False
        ma = Worksheets(1).Range("B" & i).Value
        
            For j = 3 To trow2
                
                la = Worksheets(2).Range("B" & j).Value
                If (la = ma) Then
                    y = True
                    GoTo cont1
                End If
            Next
cont1:
            For k = 3 To trow3
                
                ha = Worksheets(3).Range("B" & k).Value
                If (ha = ma) Then
                    Z = True
                    GoTo cont2
                End If
            Next
cont2:
        
        If (x = True And (y = True Or Z = True)) Then
            Worksheets(4).Range("A" & trk & " : " & "C" & trk).Value = Worksheets(1).Range("A" & i & " : " & "C" & i).Value
            If (y = True And Z = True) Then
                Worksheets(4).Range("D" & trk).Value = "in Med opt out, low plan, and high plan"
                Worksheets(4).Range("D" & trk).Interior.Color = vbRed
            ElseIf (Z = False) Then
                Worksheets(4).Range("D" & trk).Value = "in Med opt out and Low Plan"
                Worksheets(4).Range("D" & trk).Interior.Color = vbRed
            ElseIf (y = False) Then
                Worksheets(4).Range("D" & trk).Value = "in Med opt out and High Plan"
                Worksheets(4).Range("D" & trk).Interior.Color = vbRed
            End If
            trk = trk + 1
        End If
             
    Next
    
        
     For i = 3 To trow2
        x = True
        y = False
        la = Worksheets(2).Range("B" & i).Value
        
        For k = 3 To trow3
                
            ha = Worksheets(3).Range("B" & k).Value
            If (ha = ma) Then
                y = True
                GoTo cont3
            End If
        Next
cont3:
        If (x = True And y = True) Then
            Worksheets(4).Range("A" & trk & " : " & "C" & trk).Value = Worksheets(2).Range("A" & i & " : " & "C" & i).Value
            Worksheets(4).Range("D" & trk).Value = "in low plan and high plan"
            Worksheets(4).Range("D" & trk).Interior.Color = vbRed
            trk = trk + 1
        End If
        
    Next
   
    Worksheets(4).Range("A2 : C2").Value = Worksheets(1).Range("A2: C2").Value
    Worksheets(4).Rows(1).Value = Worksheets(1).Rows(1).Value
    Worksheets(4).Range("D2").Value = "Multiple Plans Enrolled"
    Worksheets(4).Columns("A:Z").AutoFit
    Flavor
End Sub

Sub Flavor()
    Range("A1:D2").Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("E8").Select
End Sub
