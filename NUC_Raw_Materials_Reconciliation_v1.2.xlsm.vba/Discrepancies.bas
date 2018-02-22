Sub getDiscrepancies()
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
    
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    ebsfield = "S C Tkt"
    scfield = "Ticket Number"
    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield).Row
    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield).Row
    
    'set ranges for source data sheets
    '"LR"=last row
    '"LC"=last column
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
    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
    ebsColumn = ebsFieldCell.Column
    ebsRow = ebsFieldCell.Row
    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
    scColumn = scFieldCell.Column
    scRow = scFieldCell.Row

    'Add Sheets for discrepancies and missing tickets
'    Sheets.Add(after:=Sheets(scWorksheet)).Name = reconciledSheet
'    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Receipts Missing From Oracle"
'    Sheets(scWorksheet).Rows(scStartingRow).EntireRow.Copy
'    Sheets("Receipts Missing From Oracle").Rows(1).PasteSpecial xlPasteValues
'
'    Sheets.Add(after:=Sheets("Receipts Missing From Oracle")).Name = "Receipts Missing From SC"
'    Sheets(ebsWorksheet).Rows(ebsStartingRow).EntireRow.Copy
'    Sheets("Receipts Missing From SC").Rows(1).PasteSpecial xlPasteValues
'
'    Sheets.Add(after:=Sheets("Receipts Missing From SC")).Name = "Void and Return to Vendor"
'    Sheets(ebsWorksheet).Rows(ebsStartingRow).EntireRow.Copy Destination:=Sheets("Void and Return to Vendor").Range("A1")
'    Sheets.Add(after:=Sheets(Sheets.Count)).Name = scWorksheet
    Sheets.Add(after:=Sheets(1)).Name = reconciledSheet
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Receipts Missing From Oracle"
    Sheets.Add(after:=Sheets("Receipts Missing From Oracle")).Name = "Receipts Missing From SC"
   
    scSheetRange.Copy
    Sheets("Receipts Missing From Oracle").Range("A1").PasteSpecial xlPasteValues
    
    ebsSheetRange.Copy
    Sheets("Receipts Missing From SC").Range("A1").PasteSpecial xlPasteValues
         
'    Sheets.Add(after:=Sheets("Void and Return to Vendor")).Name = "Weight Discrepancies"
'    Set aCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Bill of Lading")
'    aCellColumn = aCell.Column
'    Set bCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Date")
'    bCellColumn = bCell.Column
'    Set cCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Number")
'    cCellColumn = cCell.Column
'    Set dCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number")
'    dCellColumn = dCell.Column
'    Set eCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description")
'    eCellColumn = eCell.Column
'    Set fCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Ship Mode")
'    fCellColumn = fCell.Column
'    Set gCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Gross Weight")
'    gCellColumn = gCell.Column
'    Set hCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity")
'    hCellColumn = fCell.Column
'    Set iCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Receipt Num")
'    iCellColumn = iCell.Column
'
'    With Sheets("Weight Discrepancies")
'        Range("A" & 1).Value = ebsfield
'        Range("B" & 1).Value = bCell.Value
'        Range("C" & 1).Value = cCell.Value
'        Range("D" & 1).Value = dCell.Value
'        Range("E" & 1).Value = eCell.Value
'        Range("F" & 1).Value = fCell.Value
'        Range("G" & 1).Value = gCell.Value
'        Range("H" & 1).Value = hCell.Value
'        Range("I" & 1).Value = iCell.Value
'        Range("J" & 1).Value = aCell.Value
'    End With
        
'    Call indexMatch("S C Tkt", "Ticket Number", "S C Tkt", 1, "sc")
'    Call indexMatch("S C Tkt", "Ticket Number", "Ticket Number", 1, "ebs")
    
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
    Set brokerCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Supplier", Lookat:=xlWhole)
    brokerColumn = brokerCell.Column
    Set supplierCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Supplier", Lookat:=xlWhole)
    supplierColumn = supplierCell.Column
    Set thirdPartySupplierCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Third Party Supplier")
    thirdPartySupplierColumn = thirdPartySupplierCell.Column
    Set itemNumberCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number")
    itemNumberColumn = itemNumberCell.Column
    Set itemDescCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description")
    itemDescColumn = itemDescCell.Column
    Set primaryQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity")
    primaryQtyColumn = primaryQtyCell.Column
    Set unitPriceCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="PO Unit Price")
    unitPriceColumn = unitPriceCell.Column
    Set grossWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Gross Weight")
    grossWtColumn = grossWtCell.Column
    Set tareWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Tare Weight")
    tareWtColumn = tareWtCell.Column
    Set netWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Net Weight")
    netWtColumn = netWtCell.Column
    Set cleanTareWtCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Clean Tare Wgt")
    cleanTareWtColumn = cleanTareWtCell.Column
    Set adjustedQtyCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Adjusted Quantity")
    adjustedQtyColumn = adjustedQtyCell.Column
    
    shipmentNumberColumn = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Shipment Num").Column
    poLineColumn = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Line Num").Column

    Set invoiceNumCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice #")
    invoiceNumColumn = invoiceNumCell.Column
    Set invoiceDateCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Date")
    invoiceDateColumn = invoiceDateCell.Column
    Set invoiceTotalCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Invoice Total")
    invoiceTotalColumn = invoiceTotalCell.Column
        
    'column headers for Reconciled Receipts sheet
'    Sheets.Add(after:=Sheets(1)).Name = reconciledSheet
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
    
'    sheets(reconciledsheet).range(sheets(reconciledsheet).cells(
'    Sheets(reconciledSheet).Range("E2:E" & ebsSheetLR).Value = Evaluate("=C2:C" & ebsSheetLR _
'    & "&""_""&" & "D2:D" & ebsSheetLR)
'    With Sheets(reconciledSheet).Range("C2:C" & ebsSheetLR)
'        .Offset(, 2).Value = Evaluate(.Address & "& ""_"" & " & .Offset(, 1).Address)
'        .Copy
'        .PasteSpecial xlPasteValues
'    End With
    
    
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

'     Sheets(reconciledSheet).Range("A2:A" & ebsSheetLR).Value = _
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
    
    For m = ebsSheetLR To 2 Step -1
        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(m, ebsColumn), _
        Sheets(scWorksheet).Columns(scColumn), 0)) Then
        Sheets(reconciledSheet).Rows(m).EntireRow.Delete
        End If
    Next
    
'    Sheets(reconciledSheet).Cells(1, (Sheets(reconciledSheet).UsedRange.Columns(Sheets(reconciledSheet). _
'    UsedRange.Columns.Count).Column + 1)).Value = "Invoice Total"
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
'    For n = 2 To reconciledLR
'        Sheets(reconciledSheet).Cells(n, reconciledLC).Value = (Sheets(reconciledSheet).Cells(n, _
'        Sheets(reconciledSheet).Rows(1).Find(what:="Primary Quantity").Column).Value * _
'        Sheets(reconciledSheet).Cells(n, Sheets(reconciledSheet).Rows(1).Find(what:="PO Unit Price").Column).Value)
'    Next
    
    With Sheets(reconciledSheet).Columns(Sheets(reconciledSheet).Rows(1).Find(what:="PO Unit Price").Column)
        .Style = "currency"
    End With
    
    'find tickets missing from Oracle
    For j = Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count To 2 Step -1
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
        Sheets(ebsWorksheet).Columns(ebsColumn), 0)) Then
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
    
    Sheets.Add(after:=Sheets("Receipts Missing From SC")).Name = "Void and Return To Vendor"
    Sheets(scWorksheet).UsedRange.Copy
    Sheets("Void and Return to Vendor").Range("A2").PasteSpecial xlPasteValues
    With Sheets("Void and Return to Vendor").Range("A1")
        .Value = "Voided ScaleConnect Receipts"
        .Font.Bold = True
        .Font.Name = "arial"
        .Font.Size = 14
    End With

    Set ebsStatusCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Type")
    ebsStatusColumn = ebsStatusCell.Column
    Set scStatusCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Status")
    scStatusColumn = scStatusCell.Column

    For i = Sheets("Void and Return to Vendor").UsedRange.Rows.Count To 3 Step -1
        If Sheets("Void and Return to Vendor").Cells(i, scStatusColumn) <> "Void" Then
        Sheets("Void and Return to Vendor").Rows(i).EntireRow.Delete
        End If
    Next

    tempLastRow = Sheets("Void and Return to Vendor").UsedRange.Rows.Count
    Sheets("Void and Return to Vendor").Activate
    With ActiveWindow
        .SplitRow = (tempLastRow + 1)
    End With

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

'    Sheets("Void and Return to Vendor").Range("A" & (tempLastRow + 3)).Activate
'    ActiveWindow.FreezePanes = True
    
    For j = Sheets("Void and Return to Vendor").UsedRange.Rows.Count To (tempLastRow + 4) Step -1
        If Sheets("Void and Return to Vendor").Cells(j, ebsStatusColumn) <> "Return to Vendor" Then
        Sheets("Void and Return to Vendor").Rows(j).EntireRow.Delete
        End If
    Next
    
    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Pending Receipts"
    Sheets(scWorksheet).UsedRange.Copy
    Sheets("Pending Receipts").Range("A1").PasteSpecial xlPasteValues
    For h = scSheetLR To 2 Step -1
        If Sheets("Pending Receipts").Cells(h, scStatusColumn) <> "Awaiting" Then
        Sheets("Pending Receipts").Rows(h).EntireRow.Delete
        End If
    Next
    
    Sheets.Add(after:=Sheets("Void and Return to Vendor")).Name = "Weight Discrepancies"
    Sheets(reconciledSheet).UsedRange.Copy
    Sheets("Weight Discrepancies").Range("A1").PasteSpecial xlPasteValues
    
    netWtColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Net Weight").Column
    primaryQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Primary Quantity").Column
    adjustedQtyColumn = Sheets("Weight Discrepancies").UsedRange.Find(what:="Adjusted Quantity").Column
    
    For p = reconciledLR To 2 Step -1
        If Sheets("Weight Discrepancies").Cells(p, netWtColumn).Value = Sheets("Weight Discrepancies"). _
        Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells(p, adjustedQtyColumn).Value _
        = Sheets("Weight Discrepancies").Cells(p, primaryQtyColumn).Value Or Sheets("Weight Discrepancies").Cells _
        (p, netWtColumn).Value = 0 Then
        Sheets("Weight Discrepancies").Rows(p).EntireRow.Delete
        End If
    Next
'
'    nextRow = Sheets("Weight Discrepancies").Range("A" & Rows.Count).End(xlUp).Row

    
    'Find used range of "Reconciled Receipts" sheet
'    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
'    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
'    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
'    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
'    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
'    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    'Find discrepancies for Oracle Report
'    For i = 2 To ebsSheetLR
'        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(i, ebsColumn), _
'            Sheets(scWorksheet).UsedRange, 0)) And Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = _
'            "RETURN TO VENDOR" Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'            Sheets("Receipts Missing From SC").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(i, ebsColumn), _
'            Sheets(scWorksheet).UsedRange, 0)) Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Receipts Missing From SC").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = "RETURN TO VENDOR" Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        Else: Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets(reconciledSheet).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        End If
'    Next
'    'This loop pulls all "Return to Vendor" records from the ebs source file
'    'and copies to "Void and Return to Vendor" sheet.
'    For i = 2 To ebsSheetLR
'        If Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = "RETURN TO VENDOR" Then
'
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count) _
'            .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'
'            'Removes any Return to Vendor receipts from the "Reconciled Receipts" page
'            Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(ebsWorksheet).Cells(i, ebsColumn).Value))
'            recordToDelete = recordCellToDelete.Row
'            Sheets(reconciledSheet).Rows(recordToDelete).Delete
'
'        End If
'    Next
    
'    Dim LastRow As Long
'    If ebsSheetLR >= scSheetLR Then
'    LastRow = ebsSheetLR
'    Else: LastRow = scSheetLR
'    End If
'
'    Sheets(scWorksheet).Rows(scStartingRow).EntireRow.Copy
'    Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(2, 0).PasteSpecial xlPasteValues
'    Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Rows.EntireRow.Font.Bold = True
'    For i = 2 To LastRow
'        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(scWorksheet).Cells(i, scColumn), _
'            Sheets(ebsWorksheet).UsedRange, 0)) And Sheets(scWorksheet).Cells(i, scStatusColumn).Value = _
'            "VOID" Then
'            Sheets(scWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'            Sheets("Receipts Missing From Oracle").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(scWorksheet).Cells(i, scColumn), _
'            Sheets(ebsWorksheet).UsedRange, 0)) Then
'            Sheets(scWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Receipts Missing From Oracle").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Sheets(scWorksheet).Cells(i, scStatusColumn).Value = "VOID" Then
'            Sheets(scWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(i, ebsColumn), _
'            Sheets(scWorksheet).UsedRange, 0)) And Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = _
'            "RETURN TO VENDOR" Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'            Sheets("Receipts Missing From SC").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(ebsWorksheet).Cells(i, ebsColumn), _
'            Sheets(scWorksheet).UsedRange, 0)) Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Receipts Missing From SC").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = "RETURN TO VENDOR" Then
'            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        ElseIf Application.WorksheetFunction.Index(returnRange, _
'            Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
'            lookupRange, 0)) <> _
'            Application.WorksheetFunction.Index(comparisonRange, _
'            Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
'            thisRange, 0)) Then
'
'            ebsweightrow = Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
'            lookupRange, 0)
'
'            nextRow = nextRow + 1
'
'            With Sheets(ebsWorksheet)
'                .Cells(ebsweightrow, ebsColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("A" & nextRow)
'                .Cells(ebsweightrow, bCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("B" & nextRow)
'                .Cells(ebsweightrow, cCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("C" & nextRow)
'                .Cells(ebsweightrow, dCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("D" & nextRow)
'                .Cells(ebsweightrow, eCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("E" & nextRow)
'                .Cells(ebsweightrow, fCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("F" & nextRow)
'                .Cells(ebsweightrow, gCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("G" & nextRow)
'                .Cells(ebsweightrow, hCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("H" & nextRow)
'                .Cells(ebsweightrow, iCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("I" & nextRow)
'                .Cells(ebsweightrow, aCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("J" & nextRow)
'            End With
'        Else: Sheets(scWorksheet).Rows(i).EntireRow.Copy
'            Sheets(reconciledSheet).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'        End If
'    Next
    
'    'This loop pulls all "Void" records from the SC source file
'    'and copies to "Void and Return to Vendor" sheet.
'    Sheets(scWorksheet).Rows(scStartingRow).EntireRow.Copy
'    Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Offset(2, 0).PasteSpecial xlPasteValues
'    Sheets("Void and Return to Vendor").Range("A" & Rows.Count).End(xlUp).Rows.EntireRow.Font.Bold = True
'    For i = 2 To scSheetLR
'        If Sheets(scWorksheet).Cells(i, scStatusColumn).Value = "Void" Then
'
'            Sheets(scWorksheet).Rows(i).EntireRow.Copy
'            Sheets("Void and Return to Vendor").Range("A" & Rows.Count) _
'            .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
'
'            'Removes any Return to Vendor receipts fromt the "Reconciled Receipts" page
'            Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(scWorksheet).Cells(i, scColumn).Value))
'            recordToDelete = recordCellToDelete.Row
'            Sheets(reconciledSheet).Rows(recordToDelete).Delete
'
'        End If
'    Next
    
    
    'Formatting for "Void and Return to Vendor" sheet
    Dim voidLR As Long
    Dim voidLC As Long
    Dim voidRange As Range
    Dim completedDateCell As Range
    Dim completedDateColumn As Long
    Dim completedDateRow As Long
    
    voidLR = Sheets("Void and Return to Vendor").UsedRange.Rows _
        (Sheets("Void and Return to Vendor").UsedRange.Rows.Count).Row
    voidLC = Sheets("void and Return to Vendor").UsedRange.Columns _
        (Sheets(scWorksheet).UsedRange.Columns.Count).Column
    Set voidRange = Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor").Cells _
    (1, 1), Sheets("Void and Return to Vendor").Cells(voidLR, voidLC))
    
    
    Set completedDateCell = Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor") _
        .Cells(1, 1), Sheets("Void and Return to Vendor").Cells(voidLR, voidLC)).Find(what:="Completed Date")
    completedDateRow = completedDateCell.Row
    completedDateColumn = completedDateCell.Column
    
    With Sheets("Void and Return to Vendor").Range(Sheets("Void and Return to Vendor").Cells(completedDateRow, completedDateColumn), _
        Sheets("Void and Return to Vendor").Cells(voidLR, completedDateColumn))
        .NumberFormat = "mm/dd/yyyy"
    End With
        
    With voidRange
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
        
'    Dim wdLR As Long
'    Dim wdLC As Long
'
'    'Call to function that finds matched tickets with weight discrepancies
''    Call indexMatchComparison("S C Tkt", "Ticket Number", "Primary Quantity", 1, "sc", "Net Weight")
'
'    'Weight Lookups for Weight Discrepancies sheet.  The function above finds the
'    'ticket number.  This block parses the respective weights from each source
'    'file and finds the difference.
'    Worksheets("Weight Discrepancies").Activate
'    Sheets("Weight Discrepancies").Columns("B:D").Insert Shift:=xlToRight, _
'        CopyOrigin:=xlFormatFromLeftOrAbove
'    Sheets("Weight Discrepancies").Range("B1").Value = "ScrapConnect Weight"
'    Sheets("Weight Discrepancies").Range("C1").Value = "Oracle Weight"
'    Sheets("Weight Discrepancies").Range("D1").Value = "Weight Differential"
'
'    wdLR = Sheets("Weight Discrepancies").UsedRange.Rows _
'    (Sheets("Weight Discrepancies").UsedRange.Rows.Count).Row
'
''    Dim lookupRange As Range
'    Dim returnCell As Range
'    Dim returnColumn As Integer
'    Dim returnRow As Integer
''    Dim returnRange As Range
''    Dim j As Long
'    Dim errorWorksheet As String
'
'    ebsfield = "S C Tkt"
'    scfield = "Ticket Number"
'
'    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
'    ebsColumn = ebsFieldCell.Column
'    ebsRow = ebsFieldCell.Row
'    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
'    scColumn = scFieldCell.Column
'    scRow = scFieldCell.Row
'
'    For j = 2 To 3
'        If Sheets("Weight Discrepancies").Cells(1, j) = "ScrapConnect Weight" Then
'            returnfield = "Net Weight"
'            Set returnCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=returnfield)
'            returnColumn = returnCell.Column
'            returnRow = returnCell.Row
'            Set returnRange = Range(Sheets(scWorksheet).Cells(returnRow, _
'            returnColumn), Sheets(scWorksheet).Cells(scSheetLR, returnColumn))
'            Set lookupRange = Range(Sheets(scWorksheet).Cells(scRow, scColumn), _
'            Sheets(scWorksheet).Cells(scSheetLR, scColumn))
'
'                For i = 2 To wdLR
'
'                    Sheets("Weight Discrepancies").Range("B" & i).Value = _
'                    Application.Index(returnRange, _
'                    Application.Match(Sheets("Weight Discrepancies").Range("A" & i).Value, _
'                    lookupRange, 0))
'                Next
'        Else
'            returnfield = "Primary Quantity"
'            Set returnCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=returnfield)
'            returnColumn = returnCell.Column
'            returnRow = returnCell.Row
'            Set returnRange = Range(Sheets(ebsWorksheet).Cells(returnRow, _
'            returnColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, returnColumn))
'            Set lookupRange = Range(Sheets(ebsWorksheet).Cells(ebsRow, ebsColumn), _
'            Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn))
'
'                For i = 2 To wdLR
'
'                    Sheets("Weight Discrepancies").Range("C" & i).Value = _
'                    Application.Index(returnRange, _
'                    Application.Match(Sheets("Weight Discrepancies").Range("A" & i).Value, _
'                    lookupRange, 0))
'                Next
'        End If
'    Next
'
'    For i = 2 To wdLR
'        If Sheets("Weight Discrepancies").Range("B" & i).Value > Sheets("Weight Discrepancies") _
'        .Range("C" & i).Value Then
'        Sheets("Weight Discrepancies").Range("D" & i).Value = (Sheets("Weight Discrepancies") _
'            .Range("B" & i).Value - Sheets("Weight Discrepancies").Range("C" & i).Value)
'        Else
'        Sheets("Weight Discrepancies").Range("D" & i).Value = (Sheets("Weight Discrepancies") _
'        .Range("C" & i).Value - Sheets("Weight Discrepancies").Range("D" & i).Value)
'        End If
'    Next
'
'    With Range(Sheets("Weight Discrepancies").Cells(1, 2), Sheets("Weight Discrepancies").Cells(wdLR, 4))
'        .Font.Bold = True
'        .Interior.Color = RGB(255, 255, 0)
'    End With
'    wdLC = Sheets("Weight Discrepancies").UsedRange.Columns _
'    (Sheets("Weight Discrepancies").UsedRange.Columns.Count).Column
'
'    Set wdrange = Sheets("Weight Discrepancies").Range(Sheets("Weight Discrepancies").Cells _
'    (1, 1), Sheets("Weight Discrepancies").Cells(wdLR, wdLC))
'
'    With wdrange
'        .Borders.LineStyle = xlContinuous
'        .Columns.AutoFit
'    End With
    
    If UserForm1.OptionButton1.Value = "False" Then
    Call printSummary
    End If

'    End With
    
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
'
'    'Hides all sheets.  Users will export to view results
'    Sheets(reconciledSheet).Visible = xlSheetHidden
'    Sheets("Pending Receipts").Visible = xlSheetHidden
'    Sheets("Weight Discrepancies").Visible = xlSheetHidden
'    Sheets("Void and Return to Vendor").Visible = xlSheetHidden
'    Sheets("Receipts Missing From Oracle").Visible = xlSheetHidden
'    Sheets("Receipts Missing From SC").Visible = xlSheetHidden
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    With UserForm1
        .findDiscrepancies.Enabled = False
        .findDiscrepancies.BackColor = RGB(214, 214, 214)
'        .invoiceMatch.Enabled = True
'        .invoiceMatch.BackColor = RGB(0, 238, 0)
    End With

    If UserForm1.OptionButton1.Value = "True" Then
    UserForm1.invoiceMatch.Enabled = True
    UserForm1.invoiceMatch.BackColor = RGB(0, 238, 0)
    Else
    UserForm1.ExportToNewWB.Enabled = True
    UserForm1.ExportToNewWB.BackColor = RGB(0, 238, 0)
    End If

    Sheets(1).Activate

End Sub