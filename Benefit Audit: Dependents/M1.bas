Sub master()
    na = CStr(InputBox("What would you like to save this new report as?"))
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:="---------Insert path here----------" & na & ".xlsx"
    
    MsgBox "Please give me the Annual FSA data", vbOKOnly
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    CopyItOver2 namez, na
    Workbooks(namez).Close
    
    Workbooks(na).Worksheets.Add After:=Worksheets(1)
    
    MsgBox "Please give me the payroll deduction FSA data", vbOKOnly
    FileOpenDialogBox
    namez2 = ActiveWorkbook.Name
    CopyItOver2 namez2, na
    Workbooks(namez2).Close
    
    Workbooks(na).Activate
    
    peryear
    ac
    
    Worksheets(1).Name = "Data"
    Worksheets(2).Name = "Payroll Deductions Data"
    
End Sub

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

Function CopyItOver2(zz, yy)
   
  Debug.Print Workbooks(yy).Sheets.Count
  Workbooks(zz).Worksheets(1).Range("A:Z").Copy
  Workbooks(yy).Worksheets(Workbooks(yy).Sheets.Count).Range("A1").PasteSpecial (xlPasteAll)
  Application.CutCopyMode = False
End Function
