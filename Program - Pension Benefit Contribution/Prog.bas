Sub prog()
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    Workbooks(namez).Worksheets(1).Range("$A$2:$H$554").AutoFilter Field:=6, Operator:= _
        xlFilterValues, Criteria2:=Array(0, "7/2/2018", 0, "12/29/2017", 0, "12/19/2016", 0 _
        , "12/21/2015", 0, "12/29/2014", 0, "12/16/2013", 0, "12/10/2012", 0, "12/19/2011", 0, _
        "12/6/2010", 0, "8/31/2009", 0, "9/29/2008", 0, "12/24/2007", 0, "10/23/2006", 0, _
        "7/8/2005", 0, "12/13/2004", 0, "12/29/2003")
    CopyItOver (namez)
    Workbooks("output").Activate
    Worksheets(1).Rows(2).AutoFilter
    Worksheets(1).Columns("A:I").AutoFit
    Worksheets.Add After:=Worksheets(1)
    trow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    'Worksheets(2).Range("A1").Value = Worksheets(1).Range("F118").Value
    'x = Format(Worksheets(2).Range("A1").Value, "mm/dd")
    'y = Format(DateAdd("d", -9, Date), "mm/dd")
    'Debug.Print x > y
    Worksheets(1).Rows(1).Copy Destination:=Worksheets(2).Rows(1)
    Worksheets(1).Rows(2).Copy Destination:=Worksheets(2).Rows(2)
    j = 3
    For i = 3 To trow
        If Worksheets(1).Range("D" & i).Value = "" Then
            GoTo cont1:
        Else
        x = Format(Worksheets(1).Range("D" & i).Value, "mm/dd")
        y1 = Format(DateAdd("d", 4, Date), "mm/dd")
        y2 = Format(DateAdd("d", -9, Date), "mm/dd")
            If x <= y1 And x >= y2 Then
                Worksheets(1).Rows(i).Copy Destination:=Worksheets(2).Rows(j)
                j = j + 1
            End If
        End If
cont1:
    Next
    
    Worksheets(2).Range("I2").Value = "Anniversary Years"
    Worksheets(2).Range("I2").Font.Bold = True
    trow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For i = 3 To trow2
        If (Worksheets(2).Range("D" & i)) = "" Then
            GoTo cont
        Else
            x = Format(Worksheets(2).Range("D" & i).Value, "mm/dd/2018")
            y = Format(Worksheets(2).Range("D" & i).Value, "yyyy")
            ad = 2018 - CInt(y)
            Worksheets(2).Range("i" & i).Value = ad
        End If
cont:
    Next
    Worksheets(2).Columns("A:K").AutoFit
    
    Dim dict As New Scripting.Dictionary
    dict.Add Key:=1, Item:=6
    dict.Add Key:=3, Item:=10
    dict.Add Key:=4, Item:=12
    dict.Add Key:=5, Item:=14
    dict.Add Key:=10, Item:=16
    dict.Add Key:=15, Item:=18
    
    Worksheets(2).Range("J2").Value = "Due for Contribution Increase?"
    Worksheets(2).Range("j2").Font.Bold = True
    Worksheets(2).Range("K2").Value = "New Contribution Percentage(%)"
    Worksheets(2).Range("k2").Font.Bold = True
    trow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For i = 3 To trow2
        If dict.Exists(Worksheets(2).Range("I" & i).Value) Then
            Worksheets(2).Range("J" & i).Value = "Yes"
            Worksheets(2).Range("K" & i).Value = dict(Worksheets(2).Range("I" & i).Value)
            Loc1 = "A" & i
            Loc2 = "K" & i
            Worksheets(2).Range(Loc1 & ":" & Loc2).Interior.Color = vbYellow
        End If
    Next
    Worksheets(2).Columns("I:W").HorizontalAlignment = xlRight
    Worksheets(2).Columns("A:M").AutoFit
    Worksheets(2).Range("A2:K2").Interior.Color = vbGreen
    Worksheets(2).Range("A1:K1").Interior.Color = vbGreen
    border (trow2)
    Worksheets(1).Name = "Data utilized"
    Worksheets(2).Name = "Output"
    
End Sub
Function CopyItOver(zz)
  Set NewBook = Workbooks.Add
  Workbooks(zz).Worksheets(1).Range("A:I").Copy
  NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteAll)
  NewBook.SaveAs Filename:="output.xlsx"
End Function
Function border(xx)
    Range("A1:K" & xx).Select
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
    Range("F13").Select
End Function

Sub FileOpenDialogBox()

'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        fullpath = .SelectedItems.Item(1)
        Workbooks.Open Filename:=fullpath
    End With
    
    
End Sub
