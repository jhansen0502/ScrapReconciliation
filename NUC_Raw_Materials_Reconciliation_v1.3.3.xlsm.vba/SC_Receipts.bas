Sub getScrapConnectReport()
On Error GoTo ErrorHandler
    'This sub allows the user to browse local machine for Oracle report
    'file.  Is set up to handle .xlsx, .xls & .csv files.
    scWorksheet = "ScrapConnect Report"
    ebsWorksheet = "Oracle Report"
    reconciledSheet = "Reconciled Receipts"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim rg As Range
    Dim choiceRange As Range
    Dim xAddress
    Dim scTxtBox As MSForms.Control
    Dim scSheetRange As Range
    Dim scTextBox As MSForms.Control
    Dim scFileName As String
    Dim scFile As Variant
    
    'get SC file
    scFile = Application.GetOpenFilename( _
    "Excel Files (*.csv;*.xls;*.xlsx), *.csv;*.xls;*.xlsx")
    If scFile = False Then Exit Sub
    
    'add new sheet
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = scWorksheet
   
    Sheets(scWorksheet).Activate
    
    ActiveSheet.DisplayPageBreaks = False
        
    'import SC file data onto new sheet
    Set rg = Application.Range("A1")
    
    xAddress = rg.Address
    
    'for .csv files
    If scFile Like "*.csv" Then
    
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & scFile, _
        Destination:=Worksheets(scWorksheet).Range(xAddress))
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
    ElseIf scFile Like "*.xls*" Then
        
        Dim scLR As Long
        Dim scLC As Long
        Dim wbCopy As Workbook
        Dim wsCopy As Worksheet
        Dim rngCopy As Range
        Dim tgtbook As Workbook
        Dim tgtsheet As Worksheet
        Dim rngpaste As Range
                
        Set wbCopy = Workbooks.Open(scFile)
        
        scLR = ActiveSheet.UsedRange.Rows _
        (ActiveSheet.UsedRange.Rows.Count).Row
        scLC = ActiveSheet.UsedRange.Columns _
        (ActiveSheet.UsedRange.Columns.Count).Column
            
        Set wsCopy = wbCopy.Worksheets(1)
        Set rngCopy = wsCopy.Range(Cells(1, 1).Address(), Cells(scLR, scLC).Address())
        Set tgtbook = ThisWorkbook
        Set tgtsheet = tgtbook.Worksheets(scWorksheet)
        Set rngpaste = tgtsheet.Range("A1")
        
        rngCopy.Copy
        tgtsheet.Paste
        
        wbCopy.Close savechanges:=False
        
    Else
        MsgBox ("You must select a valid Excel file type (*.xls; *.xlsx; *.csv)")
        Sheets(scWorksheet).Delete
    End If
    
    scfield = "Ticket Number"
    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield).Row
    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
    scColumn = scFieldCell.Column
    scRow = scFieldCell.Row

    For i = scStartingRow - 1 To 1 Step -1
        Sheets(scWorksheet).Rows(i).Delete
    Next
    
    'find used range of sheet
    scSheetLR = ActiveSheet.UsedRange.Rows _
    (ActiveSheet.UsedRange.Rows.Count).Row
    scSheetLC = ActiveSheet.UsedRange.Columns _
    (ActiveSheet.UsedRange.Columns.Count).Column
    Set scSheetRange = Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(1, 1), _
    Sheets(scWorksheet).Cells(scSheetLR, scSheetLC))
    
    'formatting
    With Sheets(scWorksheet)
        .Range(Sheets(scWorksheet).Cells(1, 1), Sheets(scWorksheet).Cells(1, scSheetLC)). _
        Font.Bold = True
    End With
    
    With scSheetRange
        .Cells.Replace what:=vbCr, Replacement:="", LookAt:=xlPart
        .Cells.Replace what:=vbLf, Replacement:="", LookAt:=xlPart
        .Cells.Replace what:=vbCrLf, Replacement:="", LookAt:=xlPart
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
        .Rows.AutoFit
        .NumberFormat = "General"
    End With

    For i = 1 To scSheetLC
        Sheets(scWorksheet).Columns(i).TextToColumns DataType:=xlDelimited
    Next
            
    With UserForm1.Controls.Item("TextBox2")
        .Value = scFile
        .ForeColor = RGB(0, 0, 255)
        .BackColor = RGB(255, 255, 255)
    End With

    With UserForm1
        .scReportUpload.Enabled = False
        .scReportUpload.BackColor = RGB(214, 214, 214)
    End With
    
    If UserForm1.OptionButton1.Value = "True" Then
        UserForm1.invReportUpload.Enabled = True
        UserForm1.invReportUpload.BackColor = RGB(0, 0, 255)
    Else
    UserForm1.OptionButton1.Enabled = False
    UserForm1.OptionButton1.ForeColor = RGB(255, 255, 255)
    UserForm1.findDiscrepancies.Enabled = True
    UserForm1.findDiscrepancies.BackColor = RGB(0, 238, 0)
    End If
        
    Sheets(scWorksheet).Visible = xlSheetHidden
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    
    Sheets(1).Activate

Exit Sub
ErrorHandler: Call ErrorHandle

End Sub
