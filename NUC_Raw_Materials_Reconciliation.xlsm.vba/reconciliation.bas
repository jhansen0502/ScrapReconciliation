Sub Reconcile()
    
    'Compares the two source files and pulls matching tickets and
    'invoice data (if available) into "Reconciled Receipts" sheet.
    reconciledSheet = "Reconciled Receipts"
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim ebsRowNum As Long
    Dim scRowNum As Long
    Dim receiptTicketCell As Range, receiptTicketColumn As Long
    Dim transactionDateCell As Range, transactionDateColumn As Long
    Dim poNumberCell As Range, poNumberColumn As Long
    Dim receiptNumberCell As Range, receiptNumberColumn As Long
    Dim supplierCell As Range, supplierColumn As Long
    Dim itemNumberCell As Range, itemNumberColumn As Long
    Dim itemDescCell As Range, itemDescColumn As Long
    Dim primaryQtyCell As Range, primaryQtyColumn As Long
    Dim unitPriceCell As Range, unitPriceColumn As Long
    Dim invoiceNumCell As Range, invoiceNumColumn As Long
    Dim invoiceDateCell As Range, invoiceDateColumn As Long
    Dim invoiceTotalCell As Range, invoiceTotalColumn As Long
    Dim receiptTicketCell_1 As Range, receiptTicket_1Column As Long
    
    'Ranges and variables for fields to compare across the two sheets.
    Set receiptTicketCell = Sheets(ebsWorksheet).UsedRange.Find(what:="S C Tkt")
    receiptTicketColumn = receiptTicketCell.Column
    ebsStartingRow = receiptTicketCell.Row
    Set receiptTicketCell_1 = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number")
    receiptTicket_1Column = receiptTicketCell_1.Column
    scStartingRow = receiptTicketCell_1.Row
    Set transactionDateCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Date")
    transactionDateColumn = transactionDateCell.Column
    Set poNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Number")
    poNumberColumn = poNumberCell.Column
    Set receiptNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Receipt Num")
    receiptNumberColumn = receiptNumberCell.Column
    Set supplierCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Supplier", LookAt:=xlWhole)
    supplierColumn = supplierCell.Column
    Set itemNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number")
    itemNumberColumn = itemNumberCell.Column
    Set itemDescCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description")
    itemDescColumn = itemDescCell.Column
    Set primaryQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity")
    primaryQtyColumn = primaryQtyCell.Column
    Set unitPriceCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="PO Unit Price")
    unitPriceColumn = unitPriceCell.Column
    Set invoiceNumCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice #")
    invoiceNumColumn = invoiceNumCell.Column
    Set invoiceDateCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Date")
    invoiceDateColumn = invoiceDateCell.Column
    Set invoiceTotalCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Total")
    invoiceTotalColumn = invoiceTotalCell.Column
        
    Sheets.Add(After:=Sheets(1)).Name = reconciledSheet
    With Sheets(reconciledSheet)
        .Range("A1").Value = receiptTicketCell_1.Value
        .Range("B1").Value = transactionDateCell.Value
        .Range("C1").Value = poNumberCell.Value
        .Range("D1").Value = receiptNumberCell.Value
        .Range("E1").Value = supplierCell.Value
        .Range("F1").Value = itemNumberCell.Value
        .Range("G1").Value = itemDescCell.Value
        .Range("H1").Value = primaryQtyCell.Value
        .Range("I1").Value = unitPriceCell.Value
        .Range("J1").Value = invoiceNumCell.Value
        .Range("K1").Value = invoiceDateCell.Value
        .Range("L1").Value = invoiceTotalCell.Value
    End With
    
    scSheetLR = Sheets(scWorksheet).UsedRange.Rows _
    (Sheets(scWorksheet).UsedRange.Rows.Count).Row
    scSheetLC = Sheets(scWorksheet).UsedRange.Columns _
    (Sheets(scWorksheet).UsedRange.Columns.Count).Column
    Set scSheetRange = Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(1, 1), _
    Sheets(scWorksheet).Cells(scSheetLR, scSheetLC))
    ebsSheetLR = Sheets(ebsWorksheet).UsedRange.Rows _
    (Sheets(ebsWorksheet).UsedRange.Rows.Count).Row
    ebsSheetLC = Sheets(ebsWorksheet).UsedRange.Columns _
    (Sheets(ebsWorksheet).UsedRange.Columns.Count).Column
    Set ebsSheetRange = Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(1, 1), _
    Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsSheetLC))

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(2, receiptTicketColumn), _
    Sheets(ebsWorksheet).Cells(ebsSheetLR, receiptTicketColumn)).Copy
    Sheets(reconciledSheet).Range("A" & Rows.Count) _
    .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues

    Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(2, receiptTicket_1Column), _
    Sheets(scWorksheet).Cells(scSheetLR, receiptTicket_1Column)).Copy
    Sheets(reconciledSheet).Range("A" & Rows.Count) _
    .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    reconcileRange.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    
    For i = reconciledLR To 2 Step -1
        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledSheet).Cells(i, 1).Value, _
            Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(1, receiptTicketColumn), Sheets(ebsWorksheet) _
            .Cells(ebsSheetLR, receiptTicketColumn)), 0)) Then
            Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(i, 1), Sheets _
            (reconciledSheet).Cells(i, reconciledLC)).Delete
        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledSheet).Cells(i, 1).Value, _
            Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(1, receiptTicket_1Column), Sheets(scWorksheet) _
            .Cells(scSheetLR, receiptTicket_1Column)), 0)) Then
            Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(i, 1), Sheets _
            (reconciledSheet).Cells(i, reconciledLC)).Delete
        Else
            ebsRowNum = Application.Match(Sheets(reconciledSheet).Cells(i, 1).Value, _
            Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(1, receiptTicketColumn), Sheets(ebsWorksheet) _
            .Cells(ebsSheetLR, receiptTicketColumn)), 0)
            
            scRowNum = Application.Match(Sheets(reconciledSheet).Cells(i, 1).Value, _
            Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(1, receiptTicket_1Column), Sheets(scWorksheet) _
            .Cells(scSheetLR, receiptTicket_1Column)), 0)

            With Sheets(reconciledSheet)
                .Cells(i, 2) = Sheets(ebsWorksheet).Cells(ebsRowNum, transactionDateColumn).Value
                .Cells(i, 3) = Sheets(ebsWorksheet).Cells(ebsRowNum, poNumberColumn).Value
                .Cells(i, 4) = Sheets(ebsWorksheet).Cells(ebsRowNum, receiptNumberColumn).Value
                .Cells(i, 5) = Sheets(ebsWorksheet).Cells(ebsRowNum, supplierColumn).Value
                .Cells(i, 6) = Sheets(ebsWorksheet).Cells(ebsRowNum, itemNumberColumn).Value
                .Cells(i, 7) = Sheets(ebsWorksheet).Cells(ebsRowNum, itemDescColumn).Value
                .Cells(i, 8) = Sheets(ebsWorksheet).Cells(ebsRowNum, primaryQtyColumn).Value
                .Cells(i, 9) = Sheets(ebsWorksheet).Cells(ebsRowNum, unitPriceColumn).Value
                .Cells(i, 10) = Sheets(scWorksheet).Cells(ebsRowNum, invoiceNumColumn).Value
                .Cells(i, 11) = Sheets(scWorksheet).Cells(ebsRowNum, invoiceDateColumn).Value
                .Cells(i, 12) = Sheets(scWorksheet).Cells(ebsRowNum, invoiceTotalColumn).Value
            End With
        End If
    Next
    
    'Recalculate used range of "Reconciled Receipts" sheet
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))

    With reconcileRange
        .Sort key1:=(Sheets(reconciledSheet).Columns(11)), _
        Header:=xlYes
        .Borders.LineStyle = xlContinuous
    End With

    Sheets(reconciledSheet).Visible = False
    Sheets(1).Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    
    
    With UserForm1
        .InvoiceSheet.Enabled = False
        .InvoiceSheet.BackColor = RGB(214, 214, 214)
        .findDiscrepancies.Enabled = True
        .findDiscrepancies.BackColor = RGB(0, 238, 0)
    End With
    
End Sub