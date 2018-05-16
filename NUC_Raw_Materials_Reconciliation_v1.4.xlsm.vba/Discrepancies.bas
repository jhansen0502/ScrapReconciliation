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
    
    'Remove any tickets from reconciled receipts table that isn't in scrapconnect
    For m = ebsSheetLR To 2 Step -1
        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(m, ebsColumn), _
        Sheets(scWorksheet).Columns(scColumn), 0)) Then
        Sheets(reconciledSheet).Rows(m).EntireRow.Delete
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
        ElseIf Sheets("Receipts Missing From Oracle").Cells(j, oracleMissingStatusColumn).Value <> "Processed" Then
        Sheets("Receipts Missing From Oracle").Rows(j).EntireRow.Delete
      
        End If
    Next
    
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
    Sheets.Add(after:=Sheets("Receipts Missing From SC")).Name = "Void and Return To Vendor"
    Sheets(scWorksheet).UsedRange.Copy
    Sheets("Void and Return to Vendor").Range("A2").PasteSpecial xlPasteValues
    Sheets("Void and Return to Vendor").Rows(2).Font.Bold = True
    With Sheets("Void and Return to Vendor").Range("A1")
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
    For i = Sheets("Void and Return to Vendor").UsedRange.Rows.Count To 3 Step -1
        If Sheets("Void and Return to Vendor").Cells(i, scStatusColumn) <> "Void" Then
        Sheets("Void and Return to Vendor").Rows(i).EntireRow.Delete
        End If
    Next
    
    Dim tempCell As Range
    Dim tempColumn As Long
    Dim firstAddress As String
    
    'Search scrapconnect ticket data for date fields and format
    With Sheets("Void and Return to Vendor").UsedRange
        Set tempCell = .Find("Date", lookat:=xlPart, MatchCase:=False)
        If Not tempCell Is Nothing Then
            firstAddress = tempCell.Address
            Do
                tempColumn = tempCell.Column
                tempRow = tempCell.Row

                Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor") _
                .Cells(tempRow, tempColumn), Sheets("Void and Return to Vendor").Cells _
                (Sheets("Void and Return to Vendor").UsedRange.Rows.Count, tempColumn)).NumberFormat = "mm/dd/yyyy"
                
                Set tempCell = .FindNext(tempCell)
            Loop While Not tempCell Is Nothing And tempCell.Address <> firstAddress
        End If
    End With
   
    tempLastRow = Sheets("Void and Return to Vendor").UsedRange.Rows.Count
    
    'copy EBS ticket data to Void/RTV table
    Sheets(ebsWorksheet).UsedRange.Copy
    With Sheets("Void and Return to Vendor")
        .Range("A" & (tempLastRow + 3)).PasteSpecial xlPasteValues
        .Range("A" & (tempLastRow + 2)).Value = "EBS Return to Vendor Receipts"
    End With

    With Sheets("Void and Return to Vendor").Range("A" & (tempLastRow + 2))
        .Font.Bold = True
        .Font.Name = "arial"
        .Font.Size = 14
    End With

    Sheets("Void and Return to Vendor").Rows(tempLastRow + 3).Font.Bold = True
    
    'purge any EBS receipts that are NOT RTV from Void/RTV table
    For j = Sheets("Void and Return to Vendor").UsedRange.Rows.Count To (tempLastRow + 4) Step -1
        If Sheets("Void and Return to Vendor").Cells(j, ebsStatusColumn) <> "RETURN TO VENDOR" Then
        Sheets("Void and Return to Vendor").Rows(j).EntireRow.Delete
        End If
    Next
    
    'search EBS data for date fields and format
    With Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor").Cells(tempLastRow + 3, 1), _
    Sheets("Void and Return to Vendor").Cells(Sheets("Void and Return to Vendor").UsedRange.Rows.Count, _
    Sheets("Void and Return to Vendor").UsedRange.Columns.Count))
        Set tempCell = .Find("Date", lookat:=xlPart, MatchCase:=False)
        If Not tempCell Is Nothing Then
            firstAddress = tempCell.Address
            Do
                tempColumn = tempCell.Column
                tempRow = tempCell.Row

                Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor") _
                .Cells(tempRow, tempColumn), Sheets("Void and Return to Vendor").Cells _
                (Sheets("Void and Return to Vendor").UsedRange.Rows.Count, tempColumn)).NumberFormat = "mm/dd/yyyy"
                
                Set tempCell = .FindNext(tempCell)
            Loop While Not tempCell Is Nothing And tempCell.Address <> firstAddress
        End If
    End With
    
    '****PURGE VOID/RTV TICKETS FROM RECONCILED RECEIPTS TABLE*****
    Dim reconciledTicketNumberColumn As Long
    Dim tempReconciledSheetRow As Long
    Dim varTicketNoColumn As Long
    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:=ebsfield, lookat:=xlWhole).Column
    varTicketNoColumn = Sheets("Void and Return to Vendor").UsedRange.Find(what:="Ticket Number", lookat:=xlWhole).Column
    Dim primaryWeightColumn
    
    For Z = 3 To Sheets("Void and Return to Vendor").UsedRange.Rows.Count
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets("Void and Return to Vendor") _
        .Cells(Z, varTicketNoColumn).Value, Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) And _
        Not Sheets("Void and Return to Vendor").Cells(Z, varTicketNoColumn).Value = ebsfield Then
        
        tempReconciledSheetRow = (Application.Match(Sheets("Void and Return to Vendor").Cells(Z, varTicketNoColumn) _
        .Value, Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0))
        
        Sheets(reconciledSheet).Rows(tempReconciledSheetRow).EntireRow.Delete
        End If
    Next
    
    
    '*****PENDING RECEIPTS TABLE*****
    'create Pending Receipts table and copy all scrapconnect ticket data
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Pending Receipts"
    Sheets(scWorksheet).UsedRange.Copy
    Sheets("Pending Receipts").Range("A1").PasteSpecial xlPasteValues
    pendingTicketNumberColumn = Sheets("Pending Receipts").UsedRange.Find(what:=scfield, lookat:=xlWhole).Column
    
    'purge receipts from Pending Receipts table that are not Pending
    For h = scSheetLR To 2 Step -1
        If Sheets("Pending Receipts").Cells(h, scStatusColumn) <> "Awaiting" Then
        Sheets("Pending Receipts").Rows(h).EntireRow.Delete
        End If
    Next
    
    'search Reconciled Receipts table for Pending tickets and delete
    For g = 2 To Sheets("Pending Receipts").UsedRange.Rows.Count
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets("Pending Receipts").Cells(g, pendingTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) Then
        
        tempReconciledSheetRow = (Application.Match(Sheets("Pending Receipts").Cells(g, pendingTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0))
        
        Sheets(reconciledSheet).Rows(tempReconciledSheetRow).EntireRow.Delete
        End If
    Next
    
    '*****WEIGHT DISCREPANCIES TABLE*****
    'create table and copy ticket data from Reconciled Receipts table
    Sheets.Add(after:=Sheets("Void and Return to Vendor")).Name = "Weight Discrepancies"
    Sheets(reconciledSheet).UsedRange.Copy
    Sheets("Weight Discrepancies").Range("A1").PasteSpecial xlPasteValues
    
    netWtColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Net Weight", lookat:=xlWhole).Column
    primaryQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Primary Quantity", lookat:=xlWhole).Column
    adjustedQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Adjusted Quantity", lookat:=xlWhole).Column
    
    'check EBS weight against scrapconnect weight.  If weights equal, delete ticket from Weight Discrepancies table
    For p = reconciledLR To 2 Step -1
        If Sheets("Weight Discrepancies").Cells(p, netWtColumn).Value = Sheets("Weight Discrepancies"). _
        Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells(p, adjustedQtyColumn).Value _
        = Sheets("Weight Discrepancies").Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells _
        (p, netWtColumn).Value = 0 Then
        Sheets("Weight Discrepancies").Rows(p).EntireRow.Delete
        End If
    Next
    
    'add Carrier information and Comments to Reconciled Receipts table (NSUT REQUEST)
    Dim q As Long
    Dim rec_ticket_num_column As Long
    Dim comment_carrier(1) As String
    Dim rec_last_column As Long
    rec_last_column = Sheets(reconciledSheet).UsedRange.Columns.Count
    rec_ticket_num_column = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column
    Sheets(reconciledSheet).Cells(1, rec_last_column + 1).Value = "Carrier"
    Sheets(reconciledSheet).Cells(1, rec_last_column + 2).Value = "Comments"
    
    For q = 2 To Sheets(reconciledSheet).UsedRange.Rows.Count
    
        If Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value > 0 Then
        comment_carrier(0) = Application.WorksheetFunction.Index( _
            Sheets(scWorksheet).Columns(carrier_column), _
            Application.Match(Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value, _
            Sheets(scWorksheet).Columns(receiptTicket_1Column)))
        comment_carrier(1) = Application.WorksheetFunction.Index( _
            Sheets(scWorksheet).Columns(comments_column), _
            Application.Match(Sheets(reconciledSheet).Cells(q, rec_ticket_num_column).Value, _
            Sheets(scWorksheet).Columns(receiptTicket_1Column)))
        
        Sheets(reconciledSheet).Cells(q, rec_last_column + 1).Value = comment_carrier(0)
        Sheets(reconciledSheet).Cells(q, rec_last_column + 2).Value = comment_carrier(1)
        End If
        
    Next q
    
    If UserForm1.OptionButton1.Value = "False" Then
    Call printSummary
    End If
        
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