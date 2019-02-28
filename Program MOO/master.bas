Sub master()
    MsgBox "Please provide me your Data sheet to be analyzed for MOO addition errors!", vbOKOnly
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    shet1
    Macro2
    ac
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
