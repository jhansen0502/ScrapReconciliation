Sub getInvoiceReport()
On Error GoTo ErrorHandler
    'This sub allows the user to browse local machine for Oracle Invoice report
    'file.  Is set up to handle .xlsx, .xls & .csv files.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    invworksheet = "Invoice Report"

    Dim rg As Range
    Dim choiceRange As Range
    Dim xAddress
    Dim invTxtBox As MSForms.Control
    Dim invSheetRange As Range
    Dim invTextBox As MSForms.Control
    Dim invFileName As String
    Dim invFile As Variant
    
    'Open file location
    invFile = Application.GetOpenFilename( _
    "Excel Files (*.csv;*.xls;*.xlsx), *.csv;*.xls;*.xlsx")
    If invFile = False Then Exit Sub
    
    'add new sheet for Oracle report data
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = invworksheet
    Sheets(invworksheet).Activate
    
    ActiveSheet.DisplayPageBreaks = False
        
    'import ebs file data onto new sheet
    Set rg = Application.Range("A1")
'    On Error GoTo 0
    
    xAddress = rg.Address
    
    'for .csv files
    If invFile Like "*.csv" Then
    
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & ebsFile, _
        Destination:=Worksheets(invworksheet).Range(xAddress))
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 936
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        
    'for .xls & .xlsx files:
    ElseIf invFile Like "*.xls*" Then
        
        Sheets(invworksheet).Activate
        
        Dim ebsLR As Long
        Dim ebsLC As Long
        Dim wbCopy As Workbook
        Dim wsCopy As Worksheet
        Dim rngCopy As Range
        Dim tgtbook As Workbook
        Dim tgtsheet As Worksheet
        Dim rngpaste As Range
                
        Set wbCopy = Workbooks.Open(invFile)
        
        invLR = ActiveSheet.UsedRange.Rows _
        (ActiveSheet.UsedRange.Rows.Count).Row
        invLC = ActiveSheet.UsedRange.Columns _
        (ActiveSheet.UsedRange.Columns.Count).Column
            
        Set wsCopy = wbCopy.Worksheets(1)
        Set rngCopy = wsCopy.Range(Cells(1, 1).Address(), Cells(invLR, invLC).Address())
        Set tgtbook = ThisWorkbook
        Set tgtsheet = tgtbook.Worksheets(invworksheet)
        Set rngpaste = tgtsheet.Range("A1")
        
        rngCopy.Copy
        tgtsheet.Paste
        
        wbCopy.Close savechanges:=False
        
    Else
        MsgBox ("You must select a valid Excel file type (*.xls; *.xlsx; *.csv)")
        Sheets(invworksheet).Delete
    End If
    
    invfield = "Receipt Num"
    invStartingRow = Sheets(invworksheet).UsedRange.Find(what:=invfield).Row
    
    For i = invStartingRow - 1 To 1 Step -1
        Sheets(invworksheet).Rows(i).Delete
    Next
    
    'find used range of sheet
    invsheetlr = ActiveSheet.UsedRange.Rows _
    (ActiveSheet.UsedRange.Rows.Count).Row
    invSheetLC = ActiveSheet.UsedRange.Columns _
    (ActiveSheet.UsedRange.Columns.Count).Column
    Set invSheetRange = Sheets(invworksheet).Range(Sheets(invworksheet).Cells(1, 1), _
    Sheets(invworksheet).Cells(invsheetlr, invSheetLC))
    
    'formatting
    With Sheets(invworksheet)
        .Range(Sheets(invworksheet).Cells(1, 1), Sheets(invworksheet).Cells(1, invSheetLC)). _
        Font.Bold = True
    End With
    
    'remove carriage returns
    With invSheetRange
        .Cells.Replace what:=vbCr, Replacement:="", LookAt:=xlPart
        .Cells.Replace what:=vbLf, Replacement:="", LookAt:=xlPart
        .Cells.Replace what:=vbCrLf, Replacement:="", LookAt:=xlPart
        .Columns.AutoFit
        .Rows.AutoFit
        .Borders.LineStyle = xlContinuous
    End With

    For i = 1 To invSheetLC
        Sheets(invworksheet).Columns(i).TextToColumns DataType:=xlDelimited
    Next
    
    'fill userform textbox with filepath
    With UserForm1.Controls.Item("TextBox3")
        .Value = invFile
        .ForeColor = RGB(0, 0, 255)
        .BackColor = RGB(255, 255, 255)
    End With
    
    'enable/disable buttons on userform
    With UserForm1
        .invReportUpload.Enabled = False
        .invReportUpload.BackColor = RGB(214, 214, 214)
        .findDiscrepancies.Enabled = True
        .findDiscrepancies.BackColor = RGB(0, 238, 0)
    End With

    Sheets(invworksheet).Visible = xlSheetHidden
    Sheets(1).Activate
    
    're-enable excel screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
Exit Sub
ErrorHandler:     Call ErrorHandle

End Sub
