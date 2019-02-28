Sub crazy()
    
    MsgBox "Hello!Please Select the MasterData you want me to analyze!", vbOKOnly, "Data"
    FileOpenDialogBox2
    namez = ActiveWorkbook.Name
    CopyItOver2 (namez)
    nname = ActiveWorkbook.Name
    
    MsgBox "Hi! please feed me the Med opt out data you pulled from adp", vbOKOnly, "MOO"
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    CopyItOver namez, nname
    Worksheets(2).Name = "MOO data"
    colorcheckMOO
    Worksheets(1).Activate
    
    MsgBox "Now feed me the Low Plan data you pulled from adp", vbOKOnly, "LP"
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    CopyItOver namez, nname
    Worksheets(2).Name = "LP data"
    colorcheckLp
    Worksheets(1).Activate
    
    MsgBox "Now feed me the High Plan data you pulled from adp", vbOKOnly, "HP"
    FileOpenDialogBox
    namez = ActiveWorkbook.Name
    CopyItOver namez, nname
    Worksheets(2).Name = "HP data"
    colorcheckHp
    Worksheets(1).Activate
    
    RRforCross
    cross
    Worksheets(4).Name = "Error Capture-MultiPlan People"
    
    Sheets(5).Select
    Sheets(5).Move Before:=Sheets(1)
    MsgBox "Congratulation Your Report was successful!!! \(°°)/", vbOKOnly, "HP"
End Sub
Function CopyItOver(zz, yy)
  Set NewBook = Workbooks(yy)
  Workbooks(zz).Worksheets(1).Range("A:Z").Copy
  NewBook.Worksheets.Add after:=NewBook.Worksheets(1)
  NewBook.Worksheets(2).Range("A1").PasteSpecial (xlPasteAll)
  Application.CutCopyMode = False
  Workbooks(zz).Close
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
Function CopyItOver2(zz)
  Set NewBook = Workbooks.Add
  Workbooks(zz).Worksheets(1).Range("A:Z").Copy
  NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteAll)
  nem = InputBox("Please Give me a name to save this workbook as")
  NewBook.SaveAs Filename:="O:\HR Department\HR Staff\Shiv Rajendran\" & nem
  Debug.Print "O:\HR Department\HR Staff\Shiv Rajendran\" & nem & ".xlsm"
End Function

Sub FileOpenDialogBox2()

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
