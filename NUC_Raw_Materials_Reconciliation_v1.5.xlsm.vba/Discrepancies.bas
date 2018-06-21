Sub getDiscrepancies()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim varLR As Long
    Dim aCell As Range, aCellColumn As Long
    Dim bCell As Range, bCellColumn As Long
    Dim cCell As Range, cCellColumn As Long
    Dim dCell As Range, dCellColumn As Long
    Dim eCell As Range, eCellColumn As Long
    Dim fCell As Range, fCellColumn As Long
    Dim gCell As Range, gCellColumn As Long
    Dim hCell As Range, hCellColumn As Long
    Dim iCell As Range, iCellColumn As Long
    Dim pendingTicketNumberColumn As Long
    
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    ebsfield = "S C Tkt"
    scfield = "Ticket Number"
    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield, lookat:=xlWhole).Row
    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield, lookat:=xlWhole).Row
    
    'set ranges for source data sheets
    '"LR"=last row
    '"LC"=last column
    'Need to replace with shorter "usedrange" versions
    
    scSheetLR = Sheets(scWorksheet).UsedRange.Rows _
    (Sheets(scWorksheet).UsedRange.Rows.Count).Row
    scSheetLC = Sheets(scWorksheet).UsedRange.Columns _
    (Sheets(scWorksheet).UsedRange.Columns.Count).Column
    Set scSheetRange = Sheets(scWorksheet).Range(Sheets(scWorksheet).Cells(scStartingRow, 1), _
    Sheets(scWorksheet).Cells(scSheetLR, scSheetLC))
    ebsSheetLR = Sheets(ebsWorksheet).UsedRange.Rows _
    (Sheets(ebsWorksheet).UsedRange.Rows.Count).Row
    ebsSheetLC = Sheets(ebsWorksheet).UsedRange.Columns _
    (Sheets(ebsWorksheet).UsedRange.Columns.Count).Column
    Set ebsSheetRange = Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, 1), _
    Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsSheetLC))
    
    'These ranges and variables find the primary keys from the ebs and SC reports
    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield, lookat:=xlWhole)
    ebsColumn = ebsFieldCell.Column
    ebsRow = ebsFieldCell.Row
    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield, lookat:=xlWhole)
    scColumn = scFieldCell.Column
    scRow = scFieldCell.Row

    'Add Sheets for discrepancies and missing tickets
    Sheets.Add(after:=Sheets(1)).Name = reconciledSheet
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Receipts Missing From Oracle"
    Sheets.Add(after:=Sheets("Receipts Missing From Oracle")).Name = "Receipts Missing From SC"
       
    'Copy scrapconnect ticket data to missing sheet
    scSheetRange.Copy
    Sheets("Receipts Missing From Oracle").Range("A1").PasteSpecial xlPasteValues
    
    'Copy EBS ticket data to missing sheet
    ebsSheetRange.Copy
    Sheets("Receipts Missing From SC").Range("A1").PasteSpecial xlPasteValues
             
    Dim ebsRowNum As Long
    Dim scRowNum As Long
    Dim receiptTicketCell As Range, receiptTicketColumn As Long
    Dim transactionDateCell As Range, transactionDateColumn As Long
    Dim poNumberCell As Range, poNumberColumn As Long
    Dim receiptNumberCell As Range, receiptNumberColumn As Long
    Dim brokerCell As Range, brokerColumn As Long
    Dim supplierCell As Range, supplierColumn As Long
    Dim itemNumberCell As Range, itemNumberColumn As Long
    Dim itemDescCell As Range, itemDescColumn As Long
    Dim primaryQtyCell As Range, primaryQtyColumn As Long
    Dim unitPriceCell As Range, unitPriceColumn As Long
    Dim invoiceNumCell As Range, invoiceNumColumn As Long
    Dim invoiceDateCell As Range, invoiceDateColumn As Long
    Dim invoiceTotalCell As Range, invoiceTotalColumn As Long
    Dim receiptTicketCell_1 As Range, receiptTicket_1Column As Long
    Dim thirdPartySupplierCell As Range, thirdPartySupplierColumn As Long
    Dim grossWtCell As Range, grossWtColumn As Long
    Dim tareWtCell As Range, tareWtColumn As Long
    Dim netWtCell As Range, netWtColumn As Long
    Dim cleanTareWtCell As Range, cleanTareWtColumn As Long
    Dim adjustedQtyCell As Range, adjustedQtyColumn As Long
    Dim shipmentNumberColumn As Long
    Dim poLineColumn As Long
    Dim carrier_column As Long
    Dim comments_column As Long
    Dim reconciledreceiptstatuscolumn As Long
    Dim ebsCarrierColumn As Long
    Dim ebsCommentColumn As Long
    
    
    'Ranges and variables for fields to compare across the two sheets.
    Set receiptTicketCell = Sheets(ebsWorksheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole)
    receiptTicketColumn = receiptTicketCell.Column
    ebsStartingRow = receiptTicketCell.Row
    Set receiptTicketCell_1 = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number", lookat:=xlWhole)
    receiptTicket_1Column = receiptTicketCell_1.Column
    scStartingRow = receiptTicketCell_1.Row
    Set transactionDateCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Date", lookat:=xlWhole)
    transactionDateColumn = transactionDateCell.Column
    Set poNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Number", lookat:=xlWhole)
    poNumberColumn = poNumberCell.Column
    Set receiptNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Receipt Num", lookat:=xlWhole)
    receiptNumberColumn = receiptNumberCell.Column
    Set brokerCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Supplier", lookat:=xlWhole)
    brokerColumn = brokerCell.Column
    Set supplierCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Supplier", lookat:=xlWhole)
    supplierColumn = supplierCell.Column
    Set thirdPartySupplierCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Third Party Supplier", lookat:=xlWhole)
    thirdPartySupplierColumn = thirdPartySupplierCell.Column
    Set itemNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number", lookat:=xlWhole)
    itemNumberColumn = itemNumberCell.Column
    Set itemDescCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description", lookat:=xlWhole)
    itemDescColumn = itemDescCell.Column
    Set primaryQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity", lookat:=xlWhole)
    primaryQtyColumn = primaryQtyCell.Column
    Set unitPriceCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="PO Unit Price", lookat:=xlWhole)
    unitPriceColumn = unitPriceCell.Column
    Set grossWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Gross Weight", lookat:=xlWhole)
    grossWtColumn = grossWtCell.Column
    Set tareWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Tare Weight", lookat:=xlWhole)
    tareWtColumn = tareWtCell.Column
    Set netWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Net Weight", lookat:=xlWhole)
    netWtColumn = netWtCell.Column
    Set cleanTareWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Clean Tare Wgt", lookat:=xlWhole)
    cleanTareWtColumn = cleanTareWtCell.Column
    Set adjustedQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Adjusted Quantity", lookat:=xlWhole)
    adjustedQtyColumn = adjustedQtyCell.Column
    ebsCarrierColumn = Sheets(ebsWorksheet).UsedRange.Find(what:="Carrier Name", lookat:=xlWhole).Column
    ebsCommentColumn = Sheets(ebsWorksheet).UsedRange.Find(what:="Comments", lookat:=xlWhole).Column
        
    shipmentNumberColumn = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Shipment Num", lookat:=xlWhole).Column
    poLineColumn = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Line Num", lookat:=xlWhole).Column

    Set invoiceNumCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice #", lookat:=xlWhole)
    invoiceNumColumn = invoiceNumCell.Column
    Set invoiceDateCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Date", lookat:=xlWhole)
    invoiceDateColumn = invoiceDateCell.Column
    Set invoiceTotalCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Total", lookat:=xlWhole)
    invoiceTotalColumn = invoiceTotalCell.Column
    carrier_column = Sheets(scWorksheet).UsedRange.Find(what:="Carrier", lookat:=xlWhole).Column
    comments_column = Sheets(scWorksheet).UsedRange.Find(what:="Comments", lookat:=xlWhole).Column
        
    'Copy select fields from EBS table to reconciled receipts table
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, receiptTicketColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, receiptTicketColumn)).Copy
        Sheets(reconciledSheet).Range("A1:A" & ebsSheetLR).PasteSpecial xlPasteValues
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, transactionDateColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, transactionDateColumn)).Copy
        Sheets(reconciledSheet).Range("B1:B" & ebsSheetLR).PasteSpecial xlPasteValues
    
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, poNumberColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, poNumberColumn)).Copy
        Sheets(reconciledSheet).Range("C1:C" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, poLineColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, poLineColumn)).Copy
        Sheets(reconciledSheet).Range("D1:D" & ebsSheetLR).PasteSpecial xlPasteValues
    
    Sheets(reconciledSheet).Cells(1, (Sheets(reconciledSheet).UsedRange.Columns(Sheets(reconciledSheet). _
    UsedRange.Columns.Count).Column) + 1).Value = poNumberCell.Value
    For i = 2 To ebsSheetLR
        Sheets(reconciledSheet).Cells(i, 5).Value = Sheets(reconciledSheet).Cells(i, 3).Value & "-" & _
        Sheets(reconciledSheet).Cells(i, 4).Value
    Next
    Sheets(reconciledSheet).Columns("C:D").Delete
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, shipmentNumberColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, shipmentNumberColumn)).Copy
        Sheets(reconciledSheet).Range("D1:D" & ebsSheetLR).PasteSpecial xlPasteValues
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, receiptNumberColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, receiptNumberColumn)).Copy
        Sheets(reconciledSheet).Range("E1:E" & ebsSheetLR).PasteSpecial xlPasteValues
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, supplierColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, supplierColumn)).Copy
        Sheets(reconciledSheet).Range("F1:F" & ebsSheetLR).PasteSpecial xlPasteValues
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, thirdPartySupplierColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, thirdPartySupplierColumn)).Copy
        Sheets(reconciledSheet).Range("G1:G" & ebsSheetLR).PasteSpecial xlPasteValues
        
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, itemNumberColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, itemNumberColumn)).Copy
        Sheets(reconciledSheet).Range("H1:H" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, itemDescColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, itemDescColumn)).Copy
        Sheets(reconciledSheet).Range("I1:I" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, grossWtColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, grossWtColumn)).Copy
        Sheets(reconciledSheet).Range("J1:J" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, tareWtColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, tareWtColumn)).Copy
        Sheets(reconciledSheet).Range("K1:K" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, netWtColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, netWtColumn)).Copy
        Sheets(reconciledSheet).Range("L1:L" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, cleanTareWtColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, cleanTareWtColumn)).Copy
        Sheets(reconciledSheet).Range("M1:M" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, adjustedQtyColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, adjustedQtyColumn)).Copy
        Sheets(reconciledSheet).Range("N1:N" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, primaryQtyColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, primaryQtyColumn)).Copy
        Sheets(reconciledSheet).Range("O1:O" & ebsSheetLR).PasteSpecial xlPasteValues

    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, unitPriceColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, unitPriceColumn)).Copy
        Sheets(reconciledSheet).Range("P1:P" & ebsSheetLR).PasteSpecial xlPasteValues
    
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, ebsCarrierColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsCarrierColumn)).Copy
        Sheets(reconciledSheet).Range("Q1:Q" & ebsSheetLR).PasteSpecial xlPasteValues
    
    Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, ebsCommentColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsCommentColumn)).Copy
        Sheets(reconciledSheet).Range("R1:R" & ebsSheetLR).PasteSpecial xlPasteValues
    
    
'   insert "Reconciled" column in reconciled receipts sheet
    With Sheets(reconciledSheet)
        .Columns(1).EntireColumn.Insert
        .Range("A1").Value = "Receipt Status"
    End With
    reconciledreceiptstatuscolumn = Sheets(reconciledSheet).UsedRange.Find(what:="Receipt Status", lookat:=xlWhole).Column
    Dim reconciledTicketNumberColumn As Long
    Dim tempReconciledSheetRow As Long
    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column

        
    'Flag any tickets from reconciled receipts table that isn't in scrapconnect with note in "Receipt Status" column on
    'reconciled sheet
    For m = ebsSheetLR To 2 Step -1
        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(m, ebsColumn), _
        Sheets(scWorksheet).Columns(scColumn), 0)) Then
        With Sheets(reconciledSheet).Cells(m, reconciledreceiptstatuscolumn)
            .Value = "Receipt Not in Scale Connect"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        End If
    Next
    
    'getting page range parameters
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    'format po price column to dollars
    With Sheets(reconciledSheet).Columns(Sheets(reconciledSheet).Rows(1).Find(what:="PO Unit Price", lookat:=xlWhole).Column)
        .Style = "currency"
    End With
    
    Dim oracleMissingStatusColumn As Long
    oracleMissingStatusColumn = Sheets("Receipts Missing From Oracle").UsedRange.Find(what:="Status", lookat:=xlWhole).Column
    
    'find tickets missing from Oracle
    For j = Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count To 2 Step -1
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
        Sheets(ebsWorksheet).Columns(ebsColumn), 0)) Then
        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
        ElseIf Sheets("Receipts Missing From Oracle").Cells(j, oracleMissingStatusColumn).Value <> "Processed" And _
        Sheets("Receipts Missing From Oracle").Cells(j, oracleMissingStatusColumn).Value <> "Processed Rejected" Then
        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
        ElseIf Sheets("Receipts Missing From Oracle").Cells(j, Sheets("Receipts Missing From Oracle").UsedRange.Find( _
        what:="Order Number", lookat:=xlWhole).Column).Value = "OUTBOUND" Then
        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
        ElseIf Sheets("Receipts Missing From Oracle").Cells(j, Sheets("Receipts Missing From Oracle").UsedRange.Find( _
        what:="Grade", lookat:=xlWhole).Column).Value Like "*OUTBOUND*" Then
        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
        End If
    Next
    
    'add Carrier information and Comments to Reconciled Receipts table (NSUT REQUEST)
'    Dim q As Long
'    Dim rec_ticket_num_column As Long
'    Dim comment_carrier(1) As String
'    Dim rec_last_column As Long
'    rec_last_column = Sheets(reconciledSheet).UsedRange.Columns.Count
'    rec_ticket_num_column = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column
'    Sheets(reconciledSheet).Cells(1, rec_last_column + 1).Value = "Carrier"
'    Sheets(reconciledSheet).Cells(1, rec_last_column + 2).Value = "Comments"
'
'    For q = 2 To Sheets(reconciledSheet).UsedRange.Rows.Count
'
'        If Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value > 0 Then
'        comment_carrier(0) = Application.WorksheetFunction.Index( _
'            Sheets(scWorksheet).Columns(carrier_column), _
'            Application.Match(Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value, _
'            Sheets(scWorksheet).Columns(receiptTicket_1Column), 0))
'        comment_carrier(1) = Application.WorksheetFunction.Index( _
'            Sheets(scWorksheet).Columns(comments_column), _
'            Application.Match(Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value, _
'            Sheets(scWorksheet).Columns(receiptTicket_1Column), 0))
'
'        Sheets(reconciledSheet).Cells(q, rec_last_column + 1).Value = comment_carrier(0)
'        Sheets(reconciledSheet).Cells(q, rec_last_column + 2).Value = comment_carrier(1)
'        End If
'
'    Next q
    
    
'   Copy remaining ScaleConnect Ticket info to reconciled sheet.
    Dim numMissingFromOracle As Long
    Dim x As Long
    Dim y As String
    y = "Receipts Missing From Oracle"

    Dim remainingScLastRow As Long, remTickNum(), remBOL(), remCompDate(), remOrderNumber(), remShipId(), remGrade(), _
    remCarrier(), remGrossWeight(), remTareWeight(), remNetWeight(), remCleanTare(), remComments() As Variant
    Dim counter As Long
    
    x = Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count - 1
    
    ReDim remTickNum(x), remBOL(x), remCompDate(x), remOrderNumber(x), remShipId(x), remGrade(x), _
    remCarrier(x), remGrossWeight(x), remTareWeight(x), remNetWeight(x), remCleanTare(x), remComments(x)
    
    counter = 0
    
    For i = 2 To x + 1
        If Sheets(y).Cells(i, Sheets(y).UsedRange.Find(what:="Status", lookat:=xlWhole).Column).Value <> "Awaiting" Then
            remTickNum(counter) = Sheets(y).Cells(i, 1)
            remBOL(counter) = Sheets(y).Cells(i, 2)
            remCompDate(counter) = Sheets(y).Cells(i, 3)
            remOrderNumber(counter) = Sheets(y).Cells(i, 6)
            remShipId(counter) = Sheets(y).Cells(i, 25)
            remGrade(counter) = Sheets(y).Cells(i, 8)
            remCarrier(counter) = Sheets(y).Cells(i, 11)
            remGrossWeight(counter) = Sheets(y).Cells(i, 13)
            remTareWeight(counter) = Sheets(y).Cells(i, 14)
            remNetWeight(counter) = Sheets(y).Cells(i, 15)
            remCleanTare(counter) = Sheets(y).Cells(i, 16)
            remComments(counter) = Sheets(y).Cells(i, 23)
            counter = counter + 1
        End If
    Next i
    
    Dim reconciledReceiptNumColumn As Long
    Dim reconciledPONumberColumn As Long
'    Dim reconciledTicketNumberColumn As Long
    Dim reconciledTransDateColumn As Long
    Dim reconciledShipmentIdColumn As Long
    Dim reconciledGrossWeightColumn As Long
    Dim reconciledTareWeightColumn As Long
    Dim reconciledNetWeightColumn As Long
    Dim reconciledCleanTareWeightColumn As Long
    Dim reconciledCommentsColumn As Long
    Dim reconciledItemDescColumn As Long
    Dim reconciledCarrierColumn As Long
    
'    Dim q As Long
    
    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column
    reconciledTransDateColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Transaction Date", lookat:=xlWhole).Column
    reconciledPONumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Po Number", lookat:=xlWhole).Column
    reconciledShipmentIdColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Shipment Num", lookat:=xlWhole).Column
    reconciledGrossWeightColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Gross Weight", lookat:=xlWhole).Column
    reconciledTareWeightColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Tare Weight", lookat:=xlWhole).Column
    reconciledNetWeightColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Net Weight", lookat:=xlWhole).Column
    reconciledCleanTareWeightColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Clean Tare Wgt", lookat:=xlWhole).Column
    reconciledCommentsColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Comments", lookat:=xlWhole).Column
    reconciledItemDescColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Item Description", lookat:=xlWhole).Column
    reconciledCarrierColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Carrier Name", lookat:=xlWhole).Column

    
    
    q = 0
    Do While q < UBound(remComments)
        this_row = Sheets(reconciledSheet).UsedRange.Rows.Count + 1
        With Sheets(reconciledSheet)
            .Cells(this_row, reconciledreceiptstatuscolumn).Value = "Receipt Not in Oracle"
            .Cells(this_row, reconciledreceiptstatuscolumn).Font.Color = RGB(255, 0, 0)
            .Cells(this_row, reconciledreceiptstatuscolumn).Font.Bold = RGB(255, 0, 0)
            .Cells(this_row, reconciledTicketNumberColumn).Value = remTickNum(q)
            .Cells(this_row, reconciledTransDateColumn).Value = remCompDate(q)
            .Cells(this_row, reconciledPONumberColumn).Value = remOrderNumber(q)
            .Cells(this_row, reconciledShipmentIdColumn).Value = remShipId(q)
            .Cells(this_row, reconciledGrossWeightColumn).Value = remGrossWeight(q)
            .Cells(this_row, reconciledTareWeightColumn).Value = remTareWeight(q)
            .Cells(this_row, reconciledNetWeightColumn).Value = remNetWeight(q)
            .Cells(this_row, reconciledCleanTareWeightColumn).Value = remCleanTare(q)
            .Cells(this_row, reconciledCommentsColumn).Value = remComments(q)
            .Cells(this_row, reconciledItemDescColumn).Value = remGrade(q)
            .Cells(this_row, reconciledCarrierColumn).Value = remCarrier(q)
        End With
        q = q + 1
    Loop

''   Flag as not in oracle on reconciled receipts page
'    Dim ij As Long
'    For ij = 2 To Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count
'        tempReconciledSheetRow = Application.Match(Sheets("Receipts Missing From Oracle").Cells(ij, scColumn).Value, _
'        Sheets(reconciledSheet).Column(reconciledTicketNumberColumn), 0)
'        Sheets(reconciledSheet).Cells(tempReconciledSheetRow, reconciledTicketNumberColumn).Value = "Receipt Not in Oracle"
'
'    Next ij
    'find tickets missing from Scaleconnect
    For k = Sheets("Receipts Missing From SC").UsedRange.Rows.Count To 2 Step -1
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(k, ebsColumn), _
        Sheets(scWorksheet).Columns(scColumn), 0)) Then
        Sheets("Receipts Missing From SC").Rows(k).EntireRow.Delete
        End If
    Next
  
    'Find "Status" columns from both source file sheets
    Dim ebsStatusColumn As Long
    Dim ebsStatusCell As Range
    Dim scStatusColumn As Long
    Dim scStatusCell As Range
    
    'create Void/RTV table and copy scrapconnect ticket data over
    Dim vartv As String
    vartv = "Void and Return To Vendor"
    Sheets.Add(after:=Sheets("Receipts Missing From SC")).Name = vartv
    Sheets(scWorksheet).UsedRange.Copy
    Sheets(vartv).Range("A2").PasteSpecial xlPasteValues
    Sheets(vartv).Rows(2).Font.Bold = True
    With Sheets(vartv).Range("A1")
        .Value = "Voided ScaleConnect Receipts"
        .Font.Bold = True
        .Font.Name = "arial"
        .Font.Size = 14
    End With

    Set ebsStatusCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Type", lookat:=xlWhole)
    ebsStatusColumn = ebsStatusCell.Column
    Set scStatusCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Status", lookat:=xlWhole)
    scStatusColumn = scStatusCell.Column
    
    'purge scrapconnect tickets that are NOT "void" from Void/RTV table
    For i = Sheets(vartv).UsedRange.Rows.Count To 3 Step -1
        If Sheets(vartv).Cells(i, scStatusColumn) <> "Void" And _
        Sheets(vartv).Cells(i, scStatusColumn) <> "Void Rejected" Then
        Sheets(vartv).Rows(i).EntireRow.Delete
        ElseIf Sheets(vartv).Cells(i, Sheets(vartv).UsedRange.Find( _
        what:="Order Number", lookat:=xlWhole).Column).Value = "OUTBOUND" Then
        Sheets(vartv).Rows(i).EntireRow.Delete
        ElseIf Sheets(vartv).Cells(i, Sheets(vartv).UsedRange.Find( _
        what:="Grade", lookat:=xlWhole).Column).Value Like "*OUTBOUND*" Then
        Sheets(vartv).Rows(i).EntireRow.Delete
        End If
    Next
    
    Dim voidLastRow As Long
    Dim voidLastColumn As Long
    voidLastRow = Sheets(vartv).UsedRange.Rows.Count
    voidLastColumn = Sheets(vartv).UsedRange.Columns.Count
    Dim tempCell As Range
    Dim tempColumn As Long
    Dim firstAddress As String

'   rearrange columns Void/RTV sheet
    Dim voidHeaderRow As Long
    Dim rtvHeaderRow As Long
    Dim rtvLastRow As Long
    
    voidHeaderRow = Sheets(vartv).UsedRange.Find(what:="Ticket Number", lookat:=xlWhole).Row
    Dim tempCol As Long
    
'   set up array for looping through column headers
    Dim HdrArray() As String
    ReDim HdrArray(4)
    HdrArray(0) = "Order Number"
    HdrArray(1) = "Shipment ID"
    HdrArray(2) = "Grade"
    HdrArray(3) = "Completed Date"
    HdrArray(4) = "Status"
    
    Dim aa, bb As Long
    aa = 0
    For bb = 3 To 7
        With Sheets(vartv).Range(Sheets(vartv).Cells(voidHeaderRow, bb), Sheets(vartv).Cells(voidLastRow, bb))
            .Insert shift:=xlToRight
        End With
        tempCol = Sheets(vartv).UsedRange.Find(what:=HdrArray(aa), lookat:=xlWhole).Column
    
        Sheets(vartv).Range(Sheets(vartv).Cells(voidHeaderRow, tempCol), Sheets(vartv).Cells(voidLastRow, tempCol)) _
        .Copy
    
        Sheets(vartv).Range(Sheets(vartv).Cells(voidHeaderRow, bb), Sheets(vartv).Cells(voidLastRow, bb)).PasteSpecial _
        xlPasteValues

        Sheets(vartv).Range(Sheets(vartv).Cells(voidHeaderRow, tempCol), Sheets(vartv).Cells(voidLastRow, tempCol)).Delete _
        shift:=xlToLeft
        
        aa = aa + 1
    Next bb


    'Search scrapconnect ticket data for date fields and format
    With Sheets(vartv).UsedRange
        Set tempCell = .Find("Date", lookat:=xlPart, MatchCase:=False)
        If Not tempCell Is Nothing Then
            firstAddress = tempCell.Address
            Do
                tempColumn = tempCell.Column
                tempRow = tempCell.Row

                Sheets(vartv).Range(Sheets(vartv) _
                .Cells(tempRow, tempColumn), Sheets(vartv).Cells _
                (Sheets(vartv).UsedRange.Rows.Count, tempColumn)).NumberFormat = "mm/dd/yyyy"
                
                Set tempCell = .FindNext(tempCell)
            Loop While Not tempCell Is Nothing And tempCell.Address <> firstAddress
        End If
    End With
   
    tempLastRow = Sheets(vartv).UsedRange.Rows.Count
    
    'copy EBS ticket data to Void/RTV table
    Sheets(ebsWorksheet).UsedRange.Copy
    With Sheets(vartv)
        .Range("A" & (tempLastRow + 3)).PasteSpecial xlPasteValues
        .Range("A" & (tempLastRow + 2)).Value = "EBS Return to Vendor Receipts"
    End With

    With Sheets(vartv).Range("A" & (tempLastRow + 2))
        .Font.Bold = True
        .Font.Name = "arial"
        .Font.Size = 14
    End With

    Sheets(vartv).Rows(tempLastRow + 3).Font.Bold = True
    
    'purge any EBS receipts that are NOT RTV from Void/RTV table
    For j = Sheets(vartv).UsedRange.Rows.Count To (tempLastRow + 4) Step -1
        If Sheets(vartv).Cells(j, ebsStatusColumn) <> "RETURN TO VENDOR" Then
        Sheets(vartv).Rows(j).EntireRow.Delete
        End If
    Next
    
    
    '****FLAG VOID/RTV TICKETS ON RECONCILED RECEIPTS TABLE*****
    Dim varTicketNoColumn As Long
    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:=ebsfield, lookat:=xlWhole).Column
    varTicketNoColumn = Sheets(vartv).UsedRange.Find(what:="Ticket Number", lookat:=xlWhole).Column
    Dim primaryWeightColumn
    
    For Z = 3 To Sheets(vartv).UsedRange.Rows.Count
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(vartv) _
        .Cells(Z, varTicketNoColumn).Value, Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) And _
        Not Sheets(vartv).Cells(Z, varTicketNoColumn).Value = ebsfield Then
        
        tempReconciledSheetRow = (Application.Match(Sheets(vartv).Cells(Z, varTicketNoColumn) _
        .Value, Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0))
        
        With Sheets(reconciledSheet).Cells(tempReconciledSheetRow, reconciledreceiptstatuscolumn)
            .Value = "Void/RTV"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        End If
    Next
    
    rtvHeaderRow = Sheets(vartv).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Row
    rtvLastRow = Sheets(vartv).UsedRange.Rows.Count
'   voidLastRow and voidLastColumn variables were set above (before adding RTV receipts to sheet)
    
'   array loop for RTV headers to rearrange columns
    Dim cc, dd As Long
    ReDim HdrArray(8)
    
    HdrArray(0) = "Bill Of Lading"
    HdrArray(1) = "Po Number"
    HdrArray(2) = "Shipment Num"
    HdrArray(3) = "Item Description"
    HdrArray(4) = "Transaction Date"
    HdrArray(5) = "Transaction Type"
    HdrArray(6) = "Shipped Date"
    HdrArray(7) = "Po Line Num"
    HdrArray(8) = "Receipt Num"
    
    cc = 0
    For dd = 2 To 10
        With Sheets(vartv).Range(Sheets(vartv).Cells(rtvHeaderRow, dd), Sheets(vartv).Cells(rtvLastRow, dd))
            .Insert shift:=xlToRight
        End With
        
        tempCol = Sheets(vartv).Rows(rtvHeaderRow).Find(what:=HdrArray(cc), lookat:=xlWhole).Column
    
        Sheets(vartv).Range(Sheets(vartv).Cells(rtvHeaderRow, tempCol), Sheets(vartv).Cells(rtvLastRow, tempCol)) _
        .Copy
    
        Sheets(vartv).Range(Sheets(vartv).Cells(rtvHeaderRow, dd), Sheets(vartv).Cells(rtvLastRow, dd)).PasteSpecial _
        xlPasteValues

        Sheets(vartv).Range(Sheets(vartv).Cells(rtvHeaderRow, tempCol), Sheets(vartv).Cells(rtvLastRow, tempCol)).Delete _
        shift:=xlToLeft
                
        cc = cc + 1
    Next dd
    
    'search EBS data for date fields and format
    With Sheets(vartv).Range(Sheets(vartv).Cells(tempLastRow + 3, 1), _
    Sheets(vartv).Cells(Sheets(vartv).UsedRange.Rows.Count, _
    Sheets(vartv).UsedRange.Columns.Count))
        Set tempCell = .Find("Date", lookat:=xlPart, MatchCase:=False)
        If Not tempCell Is Nothing Then
            firstAddress = tempCell.Address
            Do
                tempColumn = tempCell.Column
                tempRow = tempCell.Row

                Sheets(vartv).Range(Sheets(vartv) _
                .Cells(tempRow, tempColumn), Sheets(vartv).Cells _
                (Sheets(vartv).UsedRange.Rows.Count, tempColumn)).NumberFormat = "mm/dd/yyyy"
                
                Set tempCell = .FindNext(tempCell)
            Loop While Not tempCell Is Nothing And tempCell.Address <> firstAddress
        End If
    End With
    

    '*****PENDING RECEIPTS TABLE*****
    'create Pending Receipts table and copy all scrapconnect ticket data
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Pending Receipts"
    Sheets(scWorksheet).UsedRange.Copy
    Sheets("Pending Receipts").Range("A1").PasteSpecial xlPasteValues
    pendingTicketNumberColumn = Sheets("Pending Receipts").UsedRange.Find(what:=scfield, lookat:=xlWhole).Column
    
    'purge receipts from Pending Receipts table that are not Pending or are outbound tickets
    For h = scSheetLR To 2 Step -1
        If Sheets("Pending Receipts").Cells(h, scStatusColumn) <> "Awaiting" Then
        Sheets("Pending Receipts").Rows(h).EntireRow.Delete
        ElseIf Sheets("Pending Receipts").Cells(h, Sheets("Pending Receipts").UsedRange.Find(what:="Order Number", _
        lookat:=xlWhole).Column).Value Like "*OUTBOUND*" Then
        Sheets("Pending Receipts").Rows(h).EntireRow.Delete
        ElseIf Sheets("Pending Receipts").Cells(h, Sheets("Pending Receipts").UsedRange.Find(what:="Grade", _
        lookat:=xlWhole).Column).Value Like "*OUTBOUND*" Then
        Sheets("Pending Receipts").Rows(h).EntireRow.Delete
        End If
    Next
    
'   insert pending tickets from scaleconnect
    Dim xx As Long
    xx = Sheets("Pending Receipts").UsedRange.Rows.Count - 1
    Dim yy As String
    yy = "Pending Receipts"
    counter = 0
    
    ReDim remTickNum(xx), remBOL(xx), remCompDate(xx), remOrderNumber(xx), remShipId(xx), remGrade(xx), _
    remCarrier(xx), remGrossWeight(xx), remTareWeight(xx), remNetWeight(xx), remCleanTare(xx), remComments(xx)
    
    For i = 2 To xx + 1
            remTickNum(counter) = Sheets(yy).Cells(i, 1)
            remBOL(counter) = Sheets(yy).Cells(i, 2)
            remCompDate(counter) = Sheets(yy).Cells(i, 3)
            remOrderNumber(counter) = Sheets(yy).Cells(i, 6)
            remShipId(counter) = Sheets(yy).Cells(i, 25)
            remGrade(counter) = Sheets(yy).Cells(i, 8)
            remCarrier(counter) = Sheets(yy).Cells(i, 11)
            remGrossWeight(counter) = Sheets(yy).Cells(i, 13)
            remTareWeight(counter) = Sheets(yy).Cells(i, 14)
            remNetWeight(counter) = Sheets(yy).Cells(i, 15)
            remCleanTare(counter) = Sheets(yy).Cells(i, 16)
            remComments(counter) = Sheets(yy).Cells(i, 23)
            counter = counter + 1
    Next i
    
        q = 0
    Do While q < UBound(remComments)
        this_row = Sheets(reconciledSheet).UsedRange.Rows.Count + 1
        With Sheets(reconciledSheet)
            .Cells(this_row, reconciledreceiptstatuscolumn).Value = "Pending"
            .Cells(this_row, reconciledreceiptstatuscolumn).Font.Color = RGB(255, 0, 0)
            .Cells(this_row, reconciledreceiptstatuscolumn).Font.Bold = RGB(255, 0, 0)
            .Cells(this_row, reconciledTicketNumberColumn).Value = remTickNum(q)
            .Cells(this_row, reconciledTransDateColumn).Value = remCompDate(q)
            .Cells(this_row, reconciledPONumberColumn).Value = remOrderNumber(q)
            .Cells(this_row, reconciledShipmentIdColumn).Value = remShipId(q)
            .Cells(this_row, reconciledGrossWeightColumn).Value = remGrossWeight(q)
            .Cells(this_row, reconciledTareWeightColumn).Value = remTareWeight(q)
            .Cells(this_row, reconciledNetWeightColumn).Value = remNetWeight(q)
            .Cells(this_row, reconciledCleanTareWeightColumn).Value = remCleanTare(q)
            .Cells(this_row, reconciledCommentsColumn).Value = remComments(q)
            .Cells(this_row, reconciledItemDescColumn).Value = remGrade(q)
            .Cells(this_row, reconciledCarrierColumn).Value = remCarrier(q)
        End With
        q = q + 1
    Loop
    
    'search Reconciled Receipts table for Pending tickets and flag as pending
    For g = 2 To Sheets("Pending Receipts").UsedRange.Rows.Count
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets("Pending Receipts").Cells(g, pendingTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) Then
        
        tempReconciledSheetRow = (Application.Match(Sheets("Pending Receipts").Cells(g, pendingTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0))
        
        With Sheets(reconciledSheet).Cells(tempReconciledSheetRow, reconciledreceiptstatuscolumn)
            .Value = "Pending"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        End If
    Next
    
    '*****WEIGHT DISCREPANCIES TABLE*****
    'create table and copy ticket data from Reconciled Receipts table
    Sheets.Add(after:=Sheets(vartv)).Name = "Weight Discrepancies"
    Sheets(reconciledSheet).UsedRange.Copy
    Sheets("Weight Discrepancies").Range("A1").PasteSpecial xlPasteValues
    
    netWtColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Net Weight", lookat:=xlWhole).Column
    primaryQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Primary Quantity", lookat:=xlWhole).Column
    adjustedQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Adjusted Quantity", lookat:=xlWhole).Column
    Dim weightSheetTicketColumn As Long
    weightSheetTicketColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column
    'check EBS weight against scrapconnect weight.  If weights equal, delete ticket from Weight Discrepancies table
    For p = Sheets("Weight Discrepancies").UsedRange.Rows.Count To 2 Step -1
        If Sheets("Weight Discrepancies").Cells(p, netWtColumn).Value = Sheets("Weight Discrepancies"). _
        Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells(p, adjustedQtyColumn).Value _
        = Sheets("Weight Discrepancies").Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells _
        (p, netWtColumn).Value = 0 Then
        Sheets("Weight Discrepancies").Rows(p).EntireRow.Delete
        ElseIf Sheets("Weight Discrepancies").Cells(p, reconciledreceiptstatuscolumn).Value Like "Receipt Not*" Then
        Sheets("Weight Discrepancies").Rows(p).EntireRow.Delete
        Else
        tempReconciledSheetRow = Application.Match(Sheets("Weight Discrepancies").Cells(p, weightSheetTicketColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)
        With Sheets(reconciledSheet).Cells(tempReconciledSheetRow, reconciledreceiptstatuscolumn)
            .Value = "Weight Discrepancy"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        
        End If
    Next
    
    Dim ii As Long
        For ii = 2 To Sheets(reconciledSheet).UsedRange.Rows.Count
        If Sheets(reconciledSheet).Cells(ii, reconciledreceiptstatuscolumn).Value = "" Then
            With Sheets(reconciledSheet).Cells(ii, reconciledreceiptstatuscolumn)
                .Value = "Complete"
                .Font.Color = RGB(0, 255, 0)
                .Font.Bold = True
            End With
        End If
    Next ii
        
        
    With Sheets(reconciledSheet).UsedRange
        .Sort key1:=Sheets(reconciledSheet).Columns(reconciledreceiptstatuscolumn), order1:=xlDescending, Header:=xlYes, _
        key2:=(Sheets(reconciledSheet).Columns(reconciledTransDateColumn)), order2:=xlDescending, Header:=xlYes
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
        .HorizontalAlignment = xlLeft
    End With
    
    
    If UserForm1.OptionButton1.Value = "False" Then
    Call printSummary
    End If
    
    Sheets(reconciledSheet).Columns(1).Columns.AutoFit
        
    're-activate excel updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    'enable/disable buttons on userform
    With UserForm1
        .findDiscrepancies.Enabled = False
        .findDiscrepancies.BackColor = RGB(214, 214, 214)
    End With
    
    'checks if invoice matching radio button enabled on userform.  if enabled,
    'go to invoice matching.  if not enabled, go to export.
    If UserForm1.OptionButton1.Value = "True" Then
    UserForm1.invoiceMatch.Enabled = True
    UserForm1.invoiceMatch.BackColor = RGB(0, 238, 0)
    Else
    UserForm1.ExportToNewWB.Enabled = True
    UserForm1.ExportToNewWB.BackColor = RGB(0, 238, 0)
    End If

    Sheets(1).Activate
    Exit Sub
ErrorHandler:     Call ErrorHandle
End Sub