Sub getOracleReport()
    
    'This sub allows the user to browse local machine for Oracle report
    'file.  Is set up to handle .xlsx, .xls & .csv files.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    ebsWorksheet = "Oracle Report"

    Dim rg As Range
    Dim choiceRange As Range
    Dim xAddress
    Dim ebsTxtBox As MSForms.Control
    Dim ebsSheetRange As Range
    Dim ebsTextBox As MSForms.Control
    Dim ebsFileName As String
    Dim ebsFile As Variant
    
    'Open file location
    ebsFile = Application.GetOpenFilename( _
    "Excel Files (*.csv;*.xls;*.xlsx), *.csv;*.xls;*.xlsx")
    If ebsFile = False Then Exit Sub
    
    'add new sheet for Oracle report data
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = ebsWorksheet
    Sheets(ebsWorksheet).Activate
    
    ActiveSheet.DisplayPageBreaks = False
        
    'import ebs file data onto new sheet
    Set rg = Application.Range("A1")
    On Error GoTo 0
    
    xAddress = rg.Address
    
    'for .csv files
    If ebsFile Like "*.csv" Then
    
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & ebsFile, _
        Destination:=Worksheets(ebsWorksheet).Range(xAddress))
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
    ElseIf ebsFile Like "*.xls*" Then
        
        Sheets(ebsWorksheet).Activate
        
        Dim ebsLR As Long
        Dim ebsLC As Long
        Dim wbCopy As Workbook
        Dim wsCopy As Worksheet
        Dim rngCopy As Range
        Dim tgtbook As Workbook
        Dim tgtsheet As Worksheet
        Dim rngpaste As Range
                
        Set wbCopy = Workbooks.Open(ebsFile)
        
        ebsLR = ActiveSheet.UsedRange.Rows _
        (ActiveSheet.UsedRange.Rows.Count).Row
        ebsLC = ActiveSheet.UsedRange.Columns _
        (ActiveSheet.UsedRange.Columns.Count).Column
            
        Set wsCopy = wbCopy.Worksheets(1)
        Set rngCopy = wsCopy.Range(Cells(1, 1).Address(), Cells(ebsLR, ebsLC).Address())
        Set tgtbook = ThisWorkbook
        Set tgtsheet = tgtbook.Worksheets(ebsWorksheet)
        Set rngpaste = tgtsheet.Range("A1")
        
        rngCopy.Copy
        tgtsheet.Paste
        
        wbCopy.Close savechanges:=False
        
    Else
        MsgBox ("You must select a valid Excel file type (*.xls; *.xlsx; *.csv)")
        Sheets(ebsWorksheet).Delete
    End If
    
<<<<<<< HEAD
    
=======
>>>>>>> c44028e6e08b3d0769644990805848ef419aa8e9
    ebsfield = "S C Tkt"
    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield).Row
    
    For i = ebsStartingRow - 1 To 1 Step -1
        Sheets(ebsWorksheet).Rows(i).Delete
    Next
    
    'find used range of sheet
    ebsSheetLR = ActiveSheet.UsedRange.Rows _
    (ActiveSheet.UsedRange.Rows.Count).Row
    ebsSheetLC = ActiveSheet.UsedRange.Columns _
    (ActiveSheet.UsedRange.Columns.Count).Column
    Set ebsSheetRange = Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(1, 1), _
    Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsSheetLC))
    
    'formatting
    With Sheets(ebsWorksheet)
        .Range(Sheets(ebsWorksheet).Cells(1, 1), Sheets(ebsWorksheet).Cells(1, ebsSheetLC)). _
        Font.Bold = True
    End With
    
    With ebsSheetRange
        .Cells.Replace what:=vbCr, Replacement:="", Lookat:=xlPart
        .Cells.Replace what:=vbLf, Replacement:="", Lookat:=xlPart
        .Cells.Replace what:=vbCrLf, Replacement:="", Lookat:=xlPart
        .Columns.AutoFit
        .Rows.AutoFit
        .Borders.LineStyle = xlContinuous
    End With

    For i = 1 To ebsSheetLC
        Sheets(ebsWorksheet).Columns(i).TextToColumns DataType:=xlDelimited
    Next

    With UserForm1.Controls.Item("TextBox1")
        .Value = ebsFile
        .ForeColor = RGB(0, 0, 255)
        .BackColor = RGB(255, 255, 255)
    End With
    
    With UserForm1
        .ebsReportUpload.Enabled = False
        .ebsReportUpload.BackColor = RGB(214, 214, 214)
        .scReportUpload.Enabled = True
        .scReportUpload.BackColor = RGB(0, 0, 255)
    End With

    If UserForm1.OptionButton1.Value = "False" Then
    UserForm1.OptionButton1.Enabled = False
    UserForm1.OptionButton1.ForeColor = RGB(255, 255, 255)
    End If
    
    Sheets(ebsWorksheet).Visible = xlSheetHidden
    Sheets(1).Activate

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False

End Sub