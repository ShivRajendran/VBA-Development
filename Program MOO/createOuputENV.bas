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
Sub Macro2()
'
' Macro2 Macro
'

'
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("A:J").Select
    Columns("A:J").EntireColumn.AutoFit
    Range("C3").Select
End Sub
