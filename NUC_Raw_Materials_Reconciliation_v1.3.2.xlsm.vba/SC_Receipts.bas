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
'    Sheets.Add(after:=Sheets(1)).Name = reconciledSheet
'    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Receipts Missing From Oracle"
'    Sheets.Add(after:=Sheets("Receipts Missing From Oracle")).Name = "Receipts Missing From SC"
    
    
    Sheets(scWorksheet).Activate
    
    ActiveSheet.DisplayPageBreaks = False
        
    'import SC file data onto new sheet
    Set rg = Application.Range("A1")
'    On Error GoTo 0
    
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
    
'    ebsfield = "S C Tkt"
'    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield).Row
    scfield = "Ticket Number"
    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield).Row
'    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
'    ebsColumn = ebsFieldCell.Column
'    ebsRow = ebsFieldCell.Row
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
    
'    ebsSheetLR = Sheets(ebsWorksheet).UsedRange.Rows _
'    (Sheets(ebsWorksheet).UsedRange.Rows.Count).Row
'    ebsSheetLC = Sheets(ebsWorksheet).UsedRange.Columns _
'    (Sheets(ebsWorksheet).UsedRange.Columns.Count).Column
'    Set ebsSheetRange = Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(1, 1), _
'    Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsSheetLC))
    
'    scSheetRange.Copy
'    Sheets("Receipts Missing From Oracle").Range("A1").PasteSpecial xlPasteValues
'
'    ebsSheetRange.Copy
'    Sheets("Receipts Missing From SC").Range("A1").PasteSpecial xlPasteValues
         
'    Set receiptTicketCell = Sheets(ebsWorksheet).UsedRange.Find(what:="S C Tkt")
'    receiptTicketColumn = receiptTicketCell.Column
'    ebsStartingRow = receiptTicketCell.Row
'    Set receiptTicketCell_1 = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number")
'    receiptTicket_1Column = receiptTicketCell_1.Column
'    scStartingRow = receiptTicketCell_1.Row
'    Set transactionDateCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Date")
'    transactionDateColumn = transactionDateCell.Column
'    Set poNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Number")
'    poNumberColumn = poNumberCell.Column
'    Set receiptNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Receipt Num")
'    receiptNumberColumn = receiptNumberCell.Column
'    Set brokerCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Supplier", Lookat:=xlWhole)
'    brokerColumn = brokerCell.Column
'    Set supplierCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Supplier", Lookat:=xlWhole)
'    suppliercolumn = supplierCell.Column
'    Set itemNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number")
'    itemNumberColumn = itemNumberCell.Column
'    Set itemDescCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description")
'    itemDescColumn = itemDescCell.Column
'    Set primaryQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity")
'    primaryQtyColumn = primaryQtyCell.Column
'    Set unitPriceCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="PO Unit Price")
'    unitPriceColumn = unitPriceCell.Column
'    Set invoiceNumCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice #")
'    invoiceNumColumn = invoiceNumCell.Column
'    Set invoiceDateCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Date")
'    invoiceDateColumn = invoiceDateCell.Column
'    Set invoiceTotalCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Total")
'    invoiceTotalColumn = invoiceTotalCell.Column
'
'    With Sheets(reconciledSheet)
'        .Range("A1").Value = receiptTicketCell_1.Value
'        .Range("B1").Value = transactionDateCell.Value
'        .Range("C1").Value = poNumberCell.Value
'        .Range("D1").Value = receiptNumberCell.Value
'        .Range("E1").Value = "Broker"
'        .Range("F1").Value = supplierCell.Value
'        .Range("G1").Value = itemNumberCell.Value
'        .Range("H1").Value = itemDescCell.Value
'        .Range("I1").Value = primaryQtyCell.Value
'        .Range("J1").Value = unitPriceCell.Value
'        .Range("K1").Value = invoiceNumCell.Value
'        .Range("L1").Value = invoiceDateCell.Value
'        .Range("M1").Value = invoiceTotalCell.Value
'    End With
    
'    Sheets(reconciledSheet).Range("A2:A" & ebsSheetLR).Value = _
'    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(2, receiptTicketColumn), _
'    Sheets(ebsWorksheet).Cells(ebsSheetLR, receiptTicketColumn)).Value
'
'    Dim tempLastRow As Long
'    tempLastRow = Sheets(reconciledSheet).UsedRange.Rows.Count
'
'    Sheets(reconciledSheet).Range("A" & tempLastRow & ":A" & tempLastRow + scSheetLR).Value = _
'    Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(2, scColumn), Sheets(scWorksheet). _
'    Cells(scSheetLR, scColumn)).Value
'
'    Sheets(reconciledSheet).Columns(1).RemoveDuplicates Columns:=Array(1), Header:=xlYes
'
'    For m = ebsSheetLR To 2 Step -1
'        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(m, ebsColumn), _
'        Sheets(scWorksheet).Columns(scColumn), 0)) Then
'        Sheets(reconciledSheet).Rows(m).EntireRow.Delete
'        End If
'    Next
'
'    'find tickets missing from SC sheet
'    For j = Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count To 2 Step -1
'        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
'        Sheets(ebsWorksheet).Columns(ebsColumn), 0)) Then
'        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
''        Sheets(reconciledSheet).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
''        Else
''        Sheets(ebsWorksheet).Rows(j).Copy
''        Sheets("Receipts Missing from SC").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        End If
'    Next
'
'    For k = Sheets("Receipts Missing From SC").UsedRange.Rows.Count To 2 Step -1
'        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(k, ebsColumn), _
'        Sheets(scWorksheet).Columns(scColumn), 0)) Then
'        Sheets("Receipts Missing From SC").Rows(k).EntireRow.Delete
'        End If
'    Next
    

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
'        .invReportUpload.Enabled = True
'        .invReportUpload.BackColor = RGB(0, 0, 255)
'        .OptionButton1.Enabled = True
'        .OptionButton1.ForeColor = RGB(0, 0, 0)
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
