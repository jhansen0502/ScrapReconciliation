Sub matchInvoices()
On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    Dim invReceiptNumColumn As Long
    Dim ebsReceiptNumColumn As Long
    Dim reconciledReceiptNumColumn As Long

    Dim invPoColumn As Long
    Dim ebsPoColumn As Long
    
    Dim invPoLineColumn As Long
    Dim ebsPoLineColumn As Long
    Dim reconciledPONumberColumn As Long
    
    Dim reconciledTicketNumberColumn As Long
    Dim scTicketNumberColumn As Long
    Dim invInvoiceAmountColumn As Long
    Dim invInvoiceTypeColumn As Long
    Dim invInvoiceDateColumn As Long
    Dim scInvoiceAmountColumn As Long
    Dim invInvoiceNumberColumn As Long
    Dim scInvoiceNumberColumn As Long
    Dim reconciledInvNumColumn As Long
    Dim reconciledInvAmountColumn As Long
    Dim reconciledInvQtyColumn As Long
    Dim invInvoiceQuantityColumn As Long
    Dim ebsReceiptQuantityColumn As Long
    Dim matchedRow As Long
    Dim reconciledPrimaryQuantityColumn As Long
    Dim reconciledInvoices As String
    Dim invStartingRow As Long
    Dim recInvoicesLC As Long
    Dim recInvTicketNumberColumn As Long
    Dim receiptVerifiedColumn As Long
    Dim invoiceVerifiedColumn As Long
    Dim recInvInvoiceNumberColumn As Long
    Dim recInvInvoiceAmountColumn As Long
    Dim recInvInvoiceQtyColumn As Long
    Dim recInvInvoicePOColumn As Long
    Dim recInvReceiptNumberColumn As Long
    Dim tempRow As Long
    Dim reconciledInvColumn As Long
    Dim reconciledTransDateColumn As Long
    
    invworksheet = "Invoice Report"
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    reconciledInvoices = "Reconciled Invoices"
    
    Sheets.Add(after:=Sheets(1)).Name = reconciledInvoices
    
    invInvoiceTypeColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Type").Column
    invInvoiceDateColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Date").Column
    invReceiptNumColumn = Sheets(invworksheet).UsedRange.Find(what:="Receipt Num").Column
    invInvoiceAmountColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Amount").Column
    invInvoiceQuantityColumn = Sheets(invworksheet).UsedRange.Find(what:="Qty Received").Column
    invInvoiceNumberColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Number").Column
    invPoColumn = Sheets(invworksheet).UsedRange.Find(what:="PO Number").Column
    invPoLineColumn = Sheets(invworksheet).UsedRange.Find(what:="PO Line Num").Column
    invStartingRow = Sheets(invworksheet).UsedRange.Find(what:="PO Line Num").Row
    invsheetlr = Sheets(invworksheet).UsedRange.Rows(Sheets(invworksheet) _
    .UsedRange.Rows.Count).Row
    
    With Sheets(invworksheet)
        .Range(Sheets(invworksheet).Cells(invStartingRow, invReceiptNumColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invReceiptNumColumn)).Copy
        Sheets(reconciledInvoices).Range("A1").PasteSpecial xlPasteValues
        
        .Range(Sheets(invworksheet).Cells(invStartingRow, invInvoiceTypeColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invInvoiceTypeColumn)).Copy
        Sheets(reconciledInvoices).Range("B1").PasteSpecial xlPasteValues
        
        .Range(Sheets(invworksheet).Cells(invStartingRow, invInvoiceNumberColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invInvoiceNumberColumn)).Copy
        Sheets(reconciledInvoices).Range("C1").PasteSpecial xlPasteValues
        
        .Range(Sheets(invworksheet).Cells(invStartingRow, invInvoiceDateColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invInvoiceDateColumn)).Copy
        Sheets(reconciledInvoices).Range("D1").PasteSpecial xlPasteValues

        .Range(Sheets(invworksheet).Cells(invStartingRow, invInvoiceQuantityColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invInvoiceQuantityColumn)).Copy
        Sheets(reconciledInvoices).Range("E1").PasteSpecial xlPasteValues

        .Range(Sheets(invworksheet).Cells(invStartingRow, invInvoiceAmountColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invInvoiceAmountColumn)).Copy
        Sheets(reconciledInvoices).Range("F1").PasteSpecial xlPasteValues
        
        .Range(Sheets(invworksheet).Cells(invStartingRow, invPoColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invPoColumn)).Copy
        Sheets(reconciledInvoices).Range("G1").PasteSpecial xlPasteValues

        .Range(Sheets(invworksheet).Cells(invStartingRow, invPoLineColumn), Sheets(invworksheet) _
        .Cells(invsheetlr, invPoLineColumn)).Copy
        Sheets(reconciledInvoices).Range("H1").PasteSpecial xlPasteValues
    End With
        
        recInvoicesLC = Sheets(reconciledInvoices).UsedRange.Columns.Count
        Sheets(reconciledInvoices).Cells(1, (recInvoicesLC + 1)).Value = "PO Number & PO Line"
        
    For i = 2 To invsheetlr
        Sheets(reconciledInvoices).Cells(i, (recInvoicesLC + 1)).Value = _
        Sheets(reconciledInvoices).Cells(i, (Sheets(reconciledInvoices).UsedRange.Find(what:= _
        "PO Number").Column)).Value & "-" & Sheets(reconciledInvoices).Cells(i, (Sheets(reconciledInvoices) _
        .UsedRange.Find(what:="PO Line Num").Column)).Value
    Next
            
    Sheets(reconciledInvoices).Columns(Sheets(reconciledInvoices).UsedRange.Find(what:= _
        "PO Number").Column).EntireColumn.Delete
    Sheets(reconciledInvoices).Columns(Sheets(reconciledInvoices).UsedRange.Find(what:= _
        "PO Line Num").Column).EntireColumn.Delete
    
    With Sheets(reconciledInvoices)
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Ticket Number"
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Receipt Verified?"
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Invoice Verified?"
        .Range("A2:C" & Sheets(reconciledInvoices).UsedRange.Rows.Count).HorizontalAlignment = xlCenter
    End With
    
    receiptVerifiedColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Receipt Verified?").Column
    invoiceVerifiedColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Verified?").Column
    recInvReceiptNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Receipt Num").Column
    recInvInvoiceNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Number").Column
    recInvInvoiceAmountColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Amount").Column
    recInvInvoiceQtyColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Qty Received").Column
    recInvInvoicePOColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="PO Number & PO Line").Column
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns.Count
    
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

    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt").Column
    recInvTicketNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Ticket Number").Column

    With Sheets(reconciledSheet)
        .Columns((reconciledTicketNumberColumn + 1)).EntireColumn.Insert
        .Cells(1, (reconciledTicketNumberColumn + 1)).Value = "Invoice Number"
    End With
    
    reconciledTransDateColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Transaction Date").Column
    reconciledInvColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoiced").Column
    reconciledInvNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Number").Column
'
'    With Sheets(reconciledSheet)
'        .Columns((reconciledInvNumColumn + 1)).EntireColumn.Insert
'        .Cells(1, (reconciledInvNumColumn + 1)).Value = "Invoice Quantity"
'    End With
'
'    reconciledInvQtyColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Quantity").Column
'
'    With Sheets(reconciledSheet)
'        .Columns((reconciledInvQtyColumn + 1)).EntireColumn.Insert
'        .Cells(1, (reconciledInvQtyColumn + 1)).Value = "Invoice Total"
'    End With
'
'    reconciledInvAmountColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Total").Column
'
'    ebsReceiptNumColumn = Sheets(ebsWorksheet).UsedRange.Find(what:="Receipt Num").Column
    reconciledReceiptNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Receipt Num").Column
    scTicketNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number").Column
'
'    Sheets.Add(after:=Sheets(reconciledSheet)).Name = "Unmatched Invoices"
'    Sheets(invworksheet).UsedRange.Copy
'    Sheets("Unmatched Invoices").Range("A1").PasteSpecial xlPasteValues
'
'
'    ebsReceiptQuantityColumn = Sheets(ebsWorksheet).UsedRange.Find(what:="Primary Quantity").Column
    scInvoiceAmountColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice Total").Column
    scInvoiceNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice #").Column
'    reconciledInvNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Number").Column
'    reconciledInvAmountColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Total").Column
    reconciledPrimaryQuantityColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Primary Quantity").Column
    reconciledPONumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Po Number").Column
    
    Dim pq As Long
    For pq = Sheets(reconciledInvoices).UsedRange.Rows.Count To 2 Step -1
        If Sheets(reconciledInvoices).Cells(pq, recInvInvoiceAmountColumn).Value < 0 Then
        Sheets(reconciledInvoices).Rows(pq).EntireRow.Delete
        End If
    Next pq
    
    For p = invsheetlr To 2 Step -1
        
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(invworksheet) _
        .Cells(p, invReceiptNumColumn), Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) And _
        Sheets(invworksheet).Cells(p, invInvoiceAmountColumn).Value >= 0 Then
'
        matchedRow = Application.Match(Sheets(invworksheet). _
        Cells(p, invReceiptNumColumn).Value, Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)
'
        With Sheets(reconciledSheet)
            .Cells(matchedRow, reconciledInvNumColumn).Value = Sheets(invworksheet).Cells(p, invInvoiceNumberColumn).Value
'            .Cells(matchedRow, reconciledInvAmountColumn).Value = Sheets(invworksheet).Cells(p, invInvoiceAmountColumn).Value
'            .Cells(matchedRow, reconciledInvQtyColumn).Value = Sheets(invworksheet).Cells(p, invInvoiceQuantityColumn).Value
        End With
'
        End If
    Next
    
    For q = 2 To reconciledLR
        If Sheets(reconciledSheet).Cells(q, reconciledInvNumColumn).Value = "" Then
                    
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = ChrW(10006)
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
        
'        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceNumberColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledTicketNumberColumn).Value, _
'        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
'        Application.Index(Sheets(invworksheet).Columns(invInvoiceNumberColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledReceiptNumColumn).Value, _
'        Sheets(invworksheet).Columns(invReceiptNumColumn), 0)) Then
'
'        With Sheets(reconciledSheet).Cells(q, 1)
'            .Value = "ERROR"
'            .Font.Bold = True
'            .Font.Color = RGB(255, 0, 0)
'        End With
'
'        With Sheets(reconciledSheet).Cells(q, reconciledInvNumColumn)
'            .Font.Bold = True
'            .Font.Underline = True
'            .Interior.Color = RGB(255, 255, 0)
'            .Font.Color = RGB(255, 0, 0)
'        End With
'
'        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceAmountColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledTicketNumberColumn).Value, _
'        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
'        Application.Index(Sheets(invworksheet).Columns(invInvoiceAmountColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledReceiptNumColumn).Value, _
'        Sheets(invworksheet).Columns(invReceiptNumColumn), 0)) Then
'
'        With Sheets(reconciledSheet).Cells(q, 1)
'            .Value = "ERROR"
'            .Font.Bold = True
'            .Font.Color = RGB(255, 0, 0)
'        End With
'
'        With Sheets(reconciledSheet).Cells(q, reconciledInvAmountColumn)
'            .Font.Bold = True
'            .Font.Underline = True
'            .Interior.Color = RGB(255, 255, 0)
'            .Font.Color = RGB(255, 0, 0)
'        End With
                
'        ElseIf _
'        Application.Index(Sheets(invworksheet).Columns(invInvoiceQuantityColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledTicketNumberColumn).Value, _
'        Sheets(invworksheet).Columns(invReceiptNumColumn), 0)) <> _
'        Application.Index(Sheets(ebsWorksheet).Columns(ebsReceiptQuantityColumn), _
'        Application.Match(Sheets(reconciledSheet).Cells(q, reconciledTicketNumberColumn).Value, _
'        Sheets(ebsWorksheet).Columns(ebsReceiptNumColumn), 0)) Then
        
'        ElseIf Sheets(reconciledSheet).Cells(q, reconciledInvQtyColumn).Value <> "" And _
'        Sheets(reconciledSheet).Cells(q, reconciledInvQtyColumn).Value <> _
'        Sheets(reconciledSheet).Cells(q, reconciledPrimaryQuantityColumn).Value Then
'
'        With Sheets(reconciledSheet).Cells(q, 1)
'            .Value = "ERROR"
'            .Font.Bold = True
'            .Font.Color = RGB(255, 0, 0)
'        End With
'
'        With Sheets(reconciledSheet).Cells(q, reconciledInvQtyColumn)
'            .Font.Bold = True
'            .Font.Underline = True
'            .Interior.Color = RGB(255, 255, 0)
'            .Font.Color = RGB(255, 0, 0)
'        End With
        
        Else
              
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = ChrW(10004)
            .Font.Bold = True
            .Font.Color = RGB(0, 255, 0)
        End With
        
        End If
    Next
    
    For R = 2 To Sheets(reconciledInvoices).UsedRange.Rows.Count
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledInvoices).Cells(R, recInvReceiptNumberColumn), _
        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) Then
        
        tempRow = Application.Match(Sheets(reconciledInvoices).Cells(R, recInvReceiptNumberColumn), _
        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)
        
        Sheets(reconciledInvoices).Cells(R, recInvTicketNumberColumn).Value = _
        Sheets(reconciledSheet).Cells(tempRow, reconciledTicketNumberColumn).Value
        End If
    Next
    
    
    'Check invoice details against DJJ & against Oracle
    For q = invsheetlr To 2 Step -1
        If Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledInvoices).Cells(q, recInvReceiptNumberColumn), _
        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) Then
        With Sheets(reconciledInvoices)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Value = "Receipt Not Reconciled"
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
        End With
        
        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) Then
        With Sheets(reconciledInvoices)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Value = "Ticket Not in ScrapConnect"
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
        End With
        
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceNumberColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceNumberColumn).Value Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceNumberColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceNumberColumn).Font.Bold = True
            .Cells(q, recInvInvoiceNumberColumn).Interior.Color = RGB(255, 255, 0)
        End With
        
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceAmountColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceAmountColumn).Value Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceAmountColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceAmountColumn).Font.Bold = True
            .Cells(q, recInvInvoiceAmountColumn).Interior.Color = RGB(255, 255, 0)
        End With
        
        ElseIf Application.Index(Sheets(reconciledSheet).Columns(reconciledPrimaryQuantityColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceQtyColumn).Value Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceQtyColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceQtyColumn).Font.Bold = True
            .Cells(q, recInvInvoiceQtyColumn).Interior.Color = RGB(255, 255, 0)
        End With
    
        ElseIf Sheets(reconciledInvoices).Cells(q, recInvInvoicePOColumn).Value <> _
        Application.Index(Sheets(reconciledSheet).Columns(reconciledPONumberColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvReceiptNumberColumn), _
        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) Then
        
        
'        ElseIf Not Application.Index(Sheets(reconciledSheet).Columns(reconciledPONumberColumn), _
'        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvReceiptNumberColumn).Value, _
'        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)).Value = _
'        Sheets(reconciledInvoices).Cells(q, recInvInvoicePOColumn).Value Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoicePOColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoicePOColumn).Font.Bold = True
            .Cells(q, recInvInvoicePOColumn).Interior.Color = RGB(255, 255, 0)
        End With
        
        Else
        
        With Sheets(reconciledInvoices)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10004)
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
        End With
        
        End If
    Next
    
    With Sheets(reconciledInvoices).UsedRange
        .Sort key1:=(Sheets(reconciledInvoices).Columns(1)), order1:=xlDescending, Header:=xlYes, _
        key2:=(Sheets(reconciledInvoices).Columns(2)), order2:=xlDescending, Header:=xlYes
        .Borders.LineStyle = xlContinuous
    End With

    With Sheets(reconciledSheet).UsedRange
        .Sort key1:=(Sheets(reconciledSheet).Columns(reconciledInvColumn)), order1:=xlDescending, Header:=xlYes, _
        key2:=(Sheets(reconciledSheet).Columns(reconciledTransDateColumn)), order2:=xlDescending, Header:=xlYes
        .Borders.LineStyle = xlContinuous
    End With




'    For i = 2 To reconciledLR
'        If Sheets(reconciledSheet).Cells(i, 1).Value = "" Then Sheets(reconciledSheet).Range(Sheets _
'        (reconciledSheet).Cells(i, 1), Sheets(reconciledSheet).Cells(i, reconciledLC)).Interior.Color = vbYellow
'    Next
    
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
    Sheets(reconciledInvoices).Visible = xlSheetHidden
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
Exit Sub
ErrorHandler:     Call ErrorHandle


End Sub
