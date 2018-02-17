Sub matchInvoices()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    Dim invReceiptNumColumn As Long
    Dim ebsReceiptNumColumn As Long
    Dim invPoColumn As Long
    Dim ebsPoColumn As Long
    Dim invPoLineColumn As Long
    Dim ebsPoLineColumn As Long
    Dim reconciledSCTktColumn As Long
    Dim scTicketNumberColumn As Long
        
    invworksheet = "Invoice Report"
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    
    With Sheets(reconciledSheet)
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Invoiced"
    End With
    
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns.Count
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows.Count
        
    With Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(2, 1), Sheets _
    (reconciledSheet).Cells(reconciledLR, 1))
        .HorizontalAlignment = xlCenter
    End With

    reconciledinvoicenumbercolumn = reconciledLC + 1
    reconciledinvoiceamountcolumn = reconciledinvoicenumbercolumn + 1
    
    With Sheets(reconciledSheet)
        .Cells(1, reconciledinvoicenumbercolumn).Value = "Invoice Number"
        .Cells(1, reconciledinvoiceamountcolumn).Value = "Invoice Total"
    End With
    
    invReceiptNumColumn = Sheets(invworksheet).UsedRange.Find(what:="Receipt Num").Column
    ebsReceiptNumColumn = Sheets(ebsWorksheet).UsedRange.Find(what:="Receipt Num").Column
    reconciledNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Receipt Num").Column
    reconciledSCTktColumn = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt").Column
    scTicketNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number").Column
    
'    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Matched Invoices"
'    Sheets(invworksheet).UsedRange.Copy
'    Sheets("Matched Invoices").Range("A1").PasteSpecial xlPasteValues
    
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Unmatched Invoices"
    Sheets(invworksheet).UsedRange.Copy
    Sheets("Unmatched Invoices").Range("A1").PasteSpecial xlPasteValues
    
    invSheetLR = Sheets(invworksheet).UsedRange.Rows(Sheets(invworksheet) _
    .UsedRange.Rows.Count).Row
    
    Dim matchedRow As Long
'    Dim reconciledLC As Long
       
    Dim ebsInvoiceAmountColumn As Long
    Dim scInvoiceAmountColumn As Long
    Dim ebsInvoiceNumberColumn As Long
    Dim scInvoiceNumberColumn As Long
    Dim reconcileInvNumColumn As Long
    Dim reconcileInvAmountColumn As Long
    
    ebsInvoiceAmountColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Amount").Column
    scInvoiceAmountColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice Total").Column
    ebsInvoiceNumberColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Number").Column
    scInvoiceNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice #").Column
    reconcileInvNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Number").Column
    reconcileInvAmountColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Total").Column
       
    For p = invSheetLR To 2 Step -1
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(invworksheet) _
        .Cells(p, invReceiptNumColumn), Sheets(reconciledSheet).Columns(reconciledNumColumn), 0)) Then
        Sheets("Unmatched Invoices").Rows(p).EntireRow.Delete
        
        matchedRow = Application.Match(Sheets(invworksheet). _
        Cells(p, invReceiptNumColumn).Value, Sheets(reconciledSheet).Columns(reconciledNumColumn), 0)
        
        With Sheets(reconciledSheet)
            .Cells(matchedRow, reconcileInvNumColumn).Value = Sheets(invworksheet).Cells(p, ebsInvoiceNumberColumn).Value
'            .Cells(matchedRow, reconcileInvNumColumn).Value = Application.Index(Sheets(invworksheet) _
'            .Cells(p, ebsInvoiceNumberColumn), Application.Match(Sheets(invworksheet) _
'            .Cells(p, invReceiptNumColumn), Sheets(reconciledSheet).Columns(reconciledNumColumn), 0))
            
            .Cells(matchedRow, reconcileInvAmountColumn).Value = Sheets(invworksheet).Cells(p, ebsInvoiceAmountColumn).Value
'            .Cells(matchedRow, reconcileInvAmountColumn).Value = Application.Index(Sheets(invworksheet) _
'            .Cells(p, ebsInvoiceAmountColumn), Application.Match(Sheets(invworksheet) _
'            .Cells(p, invReceiptNumColumn), Sheets(reconciledSheet).Columns(reconciledNumColumn), 0))
        End With
            
        End If
    Next
    
    For q = 2 To reconciledLR
        If Sheets(reconciledSheet).Cells(q, reconcileInvNumColumn).Value = "" Then
                    
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = ChrW(10006)
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
        
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceNumberColumn), _
        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledSCTktColumn).Value, Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Application.Index(Sheets(invworksheet).Columns(ebsInvoiceNumberColumn), _
        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledNumColumn).Value, Sheets(invworksheet).Columns(invReceiptNumColumn), 0)) Then
        
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = "ERROR"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
        
        With Sheets(reconciledSheet).Cells(q, reconcileInvNumColumn)
            .Font.Bold = True
            .Font.Underline = True
            .Interior.Color = RGB(255, 255, 0)
            .Font.Color = RGB(255, 0, 0)
        End With
                
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceAmountColumn), _
        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledSCTktColumn).Value, Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Application.Index(Sheets(invworksheet).Columns(ebsInvoiceAmountColumn), _
        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledNumColumn).Value, Sheets(invworksheet).Columns(invReceiptNumColumn), 0)) Then
        
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = "ERROR"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
        
        With Sheets(reconciledSheet).Cells(q, reconcileInvAmountColumn)
            .Font.Bold = True
            .Font.Underline = True
            .Interior.Color = RGB(255, 255, 0)
            .Font.Color = RGB(255, 0, 0)
        End With
                
        Else
              
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = ChrW(10004)
            .Font.Bold = True
            .Font.Color = RGB(0, 255, 0)
        End With
        
        End If
    Next




'    For i = 2 To reconciledLR
'        If Sheets(reconciledSheet).Cells(i, 1).Value = "" Then Sheets(reconciledSheet).Range(Sheets _
'        (reconciledSheet).Cells(i, 1), Sheets(reconciledSheet).Cells(i, reconciledLC)).Interior.Color = vbYellow
'    Next
    
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns.Count
'    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows.Count
'
'    Dim reconciledRange As Range
'    Set reconciledRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), Sheets(reconciledSheet) _
'    .Cells(reconciledLR, reconciledLC))
'
''    With reconciledRange
''        .Sort key1:=(Sheets(reconciledSheet).Columns(1)), order1:=xlDescending, Header:=xlYes
''        .Borders.LineStyle = xlContinuous
''    End With
'
'
'    Dim sheetRange As Range
'    Dim sheetlr As Long
'    Dim sheetlc As Long
'
'    For i = 2 To Sheets.Count
'        Sheets(i).Activate
'
'        sheetlr = Sheets(i).UsedRange.Rows _
'        (Sheets(i).UsedRange.Rows.Count).Row
'
'        sheetlc = Sheets(i).UsedRange.Rows _
'        (Sheets(i).UsedRange.Rows.Count).Row
'
'        Set sheetRange = Sheets(i).Range(Sheets(i).Cells(1, 1), Sheets(i).Cells(sheetlr, sheetlc))
'
'        For j = 1 To sheetlc
'            If InStr(1, Sheets(i).Cells(1, j).Value, "Date") <> 0 Then
'            Sheets(i).Columns(j).NumberFormat = "mm/dd/yyyy"
'            End If
'        Next j
'
'        With Sheets(i).UsedRange
'            .Rows(1).Font.Bold = True
'            .Columns.AutoFit
'        End With
'
'        ActiveWindow.FreezePanes = False
'        Sheets(i).Rows(2).Select
'        ActiveWindow.FreezePanes = True
'
'        Sheets(i).Rows(1).EntireRow.Insert
'        With Worksheets(i)
'            .Hyperlinks.Add Anchor:=.Range("A1"), Address:="", SubAddress:="'" & Worksheets(1).Name _
'            & "'" & "!A1", TextToDisplay:="Home"
'            With .Range("A1")
'                .Font.Bold = True
'                .Font.Color = RGB(214, 214, 214)
'                .Font.Size = 16
'                .Font.Name = "arial"
'                .RowHeight = 30
'                .ColumnWidth = 15
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                .Interior.Color = RGB(0, 15, 230)
'            End With
'        End With
'    Next i
    
    If UserForm1.OptionButton1.Value = "True" Then
    Call printSummary
    End If
''*******************************************************  RESULTS SUMMARY ****************************************
'
'    'This section calculates and prints summary data on "Home" page
'    Dim lastRowMissingOracleSheet As Long
'    Dim lastRowMissingSCSheet As Long
'
'
'    ebsWorksheet = "Oracle Report"
'    scWorksheet = "ScrapConnect Report"
'    reconciledSheet = "Reconciled Receipts"
'    ebsfield = "S C Tkt"
'    scfield = "Ticket Number"
'    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield).Row
'    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield).Row
'
'    'set ranges for source data sheets
'    '"LR"=last row
'    '"LC"=last column
'    scSheetLR = Sheets(scWorksheet).UsedRange.Rows _
'    (Sheets(scWorksheet).UsedRange.Rows.Count).Row
'    scSheetLC = Sheets(scWorksheet).UsedRange.Columns _
'    (Sheets(scWorksheet).UsedRange.Columns.Count).Column
'    Set scSheetRange = Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(scStartingRow, 1), _
'    Sheets(scWorksheet).Cells(scSheetLR, scSheetLC))
'    ebsSheetLR = Sheets(ebsWorksheet).UsedRange.Rows _
'    (Sheets(ebsWorksheet).UsedRange.Rows.Count).Row
'    ebsSheetLC = Sheets(ebsWorksheet).UsedRange.Columns _
'    (Sheets(ebsWorksheet).UsedRange.Columns.Count).Column
'    Set ebsSheetRange = Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, 1), _
'    Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsSheetLC))
'
'    'These ranges and variables find the primary keys from the ebs and SC reports
'    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
'    ebsColumn = ebsFieldCell.Column
'    ebsRow = ebsFieldCell.Row
'    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
'    scColumn = scFieldCell.Column
'    scRow = scFieldCell.Row
'
'
'    varLR = Sheets("Void and Return to Vendor").UsedRange.Rows _
'    (Sheets("Void and Return to Vendor").UsedRange.Rows.Count).Row
'
'    wdLR = Sheets("Weight Discrepancies").UsedRange.Rows _
'    (Sheets("Weight Discrepancies").UsedRange.Rows.Count).Row
'
'    lastRowMissingOracleSheet = Sheets("Receipts Missing From Oracle").UsedRange.Rows _
'    (Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count).Row
'
'    lastRowMissingSCSheet = Sheets("Receipts Missing From SC").UsedRange.Rows _
'    (Sheets("Receipts Missing From SC").UsedRange.Rows.Count).Row
'
'    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
'    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
'    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
'    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
'    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
'    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
'
'    a = Application.Count(Sheets(scWorksheet).Range(Sheets(scWorksheet) _
'    .Cells(scStartingRow, scColumn), Sheets(scWorksheet).Cells(scSheetLR, scColumn)))
'    b = Application.Count(Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet) _
'    .Cells(ebsStartingRow, ebsColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn)))
'    c = Application.Count(Sheets("Receipts Missing From SC") _
'    .Range(Sheets("Receipts Missing From SC").Cells(ebsStartingRow, ebsColumn), _
'    Sheets("Receipts Missing From SC").Cells(lastRowMissingSCSheet, ebsColumn)))
'    d = Application.Count(Sheets("Receipts Missing From Oracle") _
'    .Range(Sheets("Receipts Missing From Oracle").Cells(scStartingRow, scColumn), _
'    Sheets("Receipts Missing From Oracle").Cells(lastRowMissingOracleSheet, scColumn)))
'    e = Application.Count(Sheets("Void and Return to Vendor").Range("A1:A" & varLR))
'    f = Application.Count(Sheets("Weight Discrepancies").Columns(1))
'    g = Application.Count(Sheets("Pending Receipts").Range("A1:A" & Sheets("Pending Receipts") _
'    .UsedRange.Rows(Sheets("Pending Receipts").UsedRange.Rows.Count).Row))
'    i = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A1:A" & Sheets(reconciledSheet) _
'    .UsedRange.Rows.Count), ChrW(10006))
'    j = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A1:A" & Sheets(reconciledSheet) _
'    .UsedRange.Rows.Count), "ERROR")
'
'
''    g = Application.WorksheetFunction.CountIf(reconcileRange.Columns(Sheets(reconciledSheet) _
'    .UsedRange.Find(what:="Invoice Total").Column), ">0")
'    h = reconciledLR
'
'    'display summary on Reconciliation page
''    With Sheets(1)
''        .Range("k2").Value = b - 1
'    With Worksheets(1)
'        .Hyperlinks.Add Anchor:=.Range("K2"), _
'        Address:="", SubAddress:="'" & Worksheets(ebsWorksheet).Name & "'" & "!A1", _
'        TextToDisplay:="'" & (b)
'        .Range("l2").Value = "Total Oracle Receipts"
'
'        .Hyperlinks.Add Anchor:=.Range("K3"), _
'        Address:="", SubAddress:="'" & Worksheets(scWorksheet).Name & "'" & "!A1", _
'        TextToDisplay:="'" & (a)
''        .Range("k3").Value = a - 1
'        .Range("l3").Value = "Total ScrapConnect Receipts"
'
'        .Hyperlinks.Add Anchor:=.Range("K4"), _
'        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
'        TextToDisplay:="'" & (h)
''        .Range("k4").Value = h - 1
'        .Range("l4").Value = "Reconciled Receipts"
'
'        .Hyperlinks.Add Anchor:=.Range("K5"), _
'        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
'        TextToDisplay:="'" & (i)
'        .Range("l5").Value = "Uninvoiced Receipts"
'
'        .Hyperlinks.Add Anchor:=.Range("K6"), _
'        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
'        TextToDisplay:="'" & (j)
'        .Range("l6").Value = "Invoices with Errors"
'
'        .Hyperlinks.Add Anchor:=.Range("K7"), _
'        Address:="", SubAddress:="'" & Worksheets("Pending Receipts").Name & "'" & "!A1", _
'        TextToDisplay:="'" & (g)
'        .Range("l7").Value = "Pending Receipts"
'
'        .Hyperlinks.Add Anchor:=.Range("K8"), _
'        Address:="", SubAddress:="'" & Worksheets("Receipts Missing From SC").Name & "'" & "!A1", _
'        TextToDisplay:="'" & (c)
''        .Range("k6").Value = c - 1
'        .Range("l8").Value = "Receipts missing from ScrapConnect"
'
'        .Hyperlinks.Add Anchor:=.Range("K9"), _
'        Address:="", SubAddress:="'" & Worksheets("Receipts Missing From Oracle").Name & "'" & "!A1", _
'        TextToDisplay:="'" & (d)
''        .Range("k7").Value = d - 1
'        .Range("l9").Value = "Receipts missing from Oracle"
'
'        .Hyperlinks.Add Anchor:=.Range("K10"), _
'        Address:="", SubAddress:="'" & Worksheets("Void and Return to Vendor").Name & "'" & "!A1", _
'        TextToDisplay:="'" & (e)
''        .Range("k8").Value = e - 2
'        .Range("l10").Value = "Voided and Return to Vendor receipts"
'
'       .Hyperlinks.Add Anchor:=.Range("K11"), _
'        Address:="", SubAddress:="'" & Worksheets("Weight Discrepancies").Name & "'" & "!A1", _
'        TextToDisplay:="'" & (f)
''        .Range("k9").Value = f - 1
'        .Range("l11").Value = "Weight discrepancies"
'     End With
'                With Sheets(1).Range("K1")
'                    .Value = "Summary - " & Format(Now, "mm/dd/yyyy HH:mm")
'                    .Font.Size = 24
'                    .Font.Bold = True
'                    .Font.Name = "arial"
'                    .Rows(1).AutoFit
'                End With
'                With Sheets(1).Range("k2:k11")
'                    .Font.Bold = True
'                    .Font.ColorIndex = 3
'                End With
'                With Sheets(1).Range("k2:l11")
'                    .Font.Size = 15
'                    .Font.Bold = True
'                    .Font.Name = "arial"
'                    .Rows.AutoFit
'                    .BorderAround ColorIndex:=0, Weight:=xlThick
'                    .Columns.AutoFit
'                End With
    
    'Hides all sheets.  Users will export to view results
'    Sheets(reconciledSheet).Visible = xlSheetHidden
'    Sheets("Pending Receipts").Visible = xlSheetHidden
'    Sheets("Weight Discrepancies").Visible = xlSheetHidden
'    Sheets("Void and Return to Vendor").Visible = xlSheetHidden
'    Sheets("Receipts Missing From Oracle").Visible = xlSheetHidden
'    Sheets("Receipts Missing From SC").Visible = xlSheetHidden
'    Sheets("Unmatched Invoices").Visible = xlSheetHidden

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    With UserForm1
        .invoiceMatch.Enabled = False
        .invoiceMatch.BackColor = RGB(214, 214, 214)
        .ExportToNewWB.Enabled = True
        .ExportToNewWB.BackColor = RGB(0, 238, 0)
    End With

    Sheets(1).Activate


End Sub
