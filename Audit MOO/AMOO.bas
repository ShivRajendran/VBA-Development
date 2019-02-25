Sub shet1()
    
    Worksheets.Add before:=Worksheets(1)
    Worksheets(1).Name = "correct Addition amounts"
    trow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    idcheck = ""
    j = 1
    For i = 1 To trow2
        aID = CStr(Worksheets(2).Range("A" & i).Value)
        If (idcheck <> aID) Then
            Worksheets(2).Rows(i).Copy
            Worksheets(1).Rows(j).PasteSpecial xlPasteValues
            j = j + 1
        End If
        idcheck = aID
    Next
    Worksheets(1).Columns("A:z").AutoFit
    
    trow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets(1).Range("I2:I" & trow).Value = 76.92
End Sub

Sub ac()
    Sheets.Add after:=Worksheets(2)
    Worksheets(3).Name = "Report"
    
    Worksheets.Add after:=Worksheets(3)
    Worksheets(4).Name = "People with Descrepency"
    
    trow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To trow
        aID = CStr(Worksheets(1).Range("B" & i).Value)
        trow2 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
        amt = (Worksheets(1).Range("I" & i).Value)
        If (trow2 <> 1) Then trow2 = trow2 + 1
        Macro1 aID, trow2, amt
        
    Next
    

End Sub

Function Macro1(x, y, z)
    strt = y
    Worksheets(2).Rows(1).Copy Destination:=Worksheets(3).Rows(y + 1)
    Worksheets(3).Range("J" & y + 1).Value = "Correct Dependent Amount"
    Worksheets(3).Range("J" & y + 1).Font.Bold = True
    Worksheets(3).Range("K" & y + 1).Value = "Descrepency"
    Worksheets(3).Range("K" & y + 1).Font.Bold = True
    Worksheets(3).Columns.AutoFit
    y = y + 1
    Worksheets(2).Range("$A$1:$I$1646").AutoFilter Field:=2, Criteria1:=x
    trow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    Rowz = Worksheets(2).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    Worksheets(2).Rows("2:" & trow).Copy
    Worksheets(3).Range("A" & y + 1).PasteSpecial xlCellTypeVisible
    trows3 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
    y = y + 1
    
    For j = y To trows3
        Worksheets(3).Range("J" & j).Value = (z)
        Worksheets(3).Range("k" & j) = Worksheets(3).Range("J" & j).Value - Worksheets(3).Range("I" & j).Value
        If (Worksheets(3).Range("k" & j).Value <> 0) Then
            flag = True
            Worksheets(3).Range("k" & j).Interior.Color = vbRed
        End If
    Next
    a = y - 2
    b = y
    Worksheets(3).Activate
    Macro3 a, b, trows3
    Worksheets(3).Columns.AutoFit
    'added
    trows4 = Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Row
    If (flag = True) Then
        zed = trows4 + 1
        For g = strt To trows3
            Worksheets(3).Rows(g).Copy
            Worksheets(4).Rows(zed).PasteSpecial xlPasteAll
            zed = zed + 1
        Next
    End If
   ' end of added
    Worksheets(4).Columns.AutoFit
    
End Function
Sub Macro3(a, b, c)
'
' Macro3 Macro
'

'
    Range("A" & b).Select
    Selection.Copy
    Range("A" & a).Select
    ActiveSheet.Paste
    Range("B" & b).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B" & a).Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-6
   
    Rows(a & ":" & a).Select
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Rows(b - 2 & ":" & c).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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
    Range("J" & b & ":J" & c).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("K" & b & ":K" & c).Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("I" & b & ":I" & c).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("J" & b & ":J" & c).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("A" & b & ":A" & c).Select
    ActiveWindow.SmallScroll Down:=-12
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Rows(a & ":" & c).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A" & c + 1).Select
End Sub

Sub shet3()
    
    Worksheets.Add after:=Worksheets(3)
    Worksheets(4).Name = "People with Descrepency"
    trow2 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
    idcheck = ""
    j = 1
    z = Worksheets(3).Range("A2").Value
    For i = 1 To trow2
        aID = CStr(Worksheets(2).Range("A" & i).Value)
        If (idcheck <> aID Or aID = z) Then
            Worksheets(3).Rows(i).Copy
            Worksheets(4).Rows(j).PasteSpecial xlPasteFormats
            j = j + 1
        End If
        idcheck = aID
    Next
    Worksheets(4).Columns("A:z").AutoFit

End Sub
