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
    
    'add table for Reconciled Invoices and define invoice report columns for lookups
    Sheets.Add(after:=Sheets(1)).Name = reconciledInvoices
    
    invInvoiceTypeColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Type", lookat:=xlWhole).Column
    invInvoiceDateColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Date", lookat:=xlWhole).Column
    invReceiptNumColumn = Sheets(invworksheet).UsedRange.Find(what:="Receipt Num", lookat:=xlWhole).Column
    invInvoiceAmountColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Amount", lookat:=xlWhole).Column
    invInvoiceQuantityColumn = Sheets(invworksheet).UsedRange.Find(what:="Qty Received", lookat:=xlWhole).Column
    invInvoiceNumberColumn = Sheets(invworksheet).UsedRange.Find(what:="Invoice Number", lookat:=xlWhole).Column
    invPoColumn = Sheets(invworksheet).UsedRange.Find(what:="PO Number", lookat:=xlWhole).Column
    invPoLineColumn = Sheets(invworksheet).UsedRange.Find(what:="PO Line Num", lookat:=xlWhole).Column
    invStartingRow = Sheets(invworksheet).UsedRange.Find(what:="PO Line Num", lookat:=xlWhole).Row
    invsheetlr = Sheets(invworksheet).UsedRange.Rows(Sheets(invworksheet) _
    .UsedRange.Rows.Count).Row
    
    'copy select fields from invoice report for ALL invoices
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
        "PO Number", lookat:=xlWhole).Column)).Value & "-" & Sheets(reconciledInvoices).Cells(i, (Sheets(reconciledInvoices) _
        .UsedRange.Find(what:="PO Line Num", lookat:=xlWhole).Column)).Value
    Next
            
    Sheets(reconciledInvoices).Columns(Sheets(reconciledInvoices).UsedRange.Find(what:= _
        "PO Number", lookat:=xlWhole).Column).EntireColumn.Delete
    Sheets(reconciledInvoices).Columns(Sheets(reconciledInvoices).UsedRange.Find(what:= _
        "PO Line Num", lookat:=xlWhole).Column).EntireColumn.Delete
    
    'insert summary data/feedback columns
    With Sheets(reconciledInvoices)
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Ticket Number"
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Receipt Verified?"
        .Columns(1).EntireColumn.Insert
        .Cells(1, 1).Value = "Invoice Verified?"
        .Range("A2:C" & Sheets(reconciledInvoices).UsedRange.Rows.Count).HorizontalAlignment = xlCenter
    End With
    
    receiptVerifiedColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Receipt Verified?", lookat:=xlWhole).Column
    invoiceVerifiedColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Verified?", lookat:=xlWhole).Column
    recInvReceiptNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Receipt Num", lookat:=xlWhole).Column
    recInvInvoiceNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Number", lookat:=xlWhole).Column
    recInvInvoiceAmountColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Amount", lookat:=xlWhole).Column
    recInvInvoiceQtyColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Qty Received", lookat:=xlWhole).Column
    recInvInvoicePOColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="PO Number & PO Line", lookat:=xlWhole).Column
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

    reconciledTicketNumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="S C Tkt", lookat:=xlWhole).Column
    recInvTicketNumberColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Ticket Number", lookat:=xlWhole).Column

    With Sheets(reconciledSheet)
        .Columns((reconciledTicketNumberColumn + 1)).EntireColumn.Insert
        .Cells(1, (reconciledTicketNumberColumn + 1)).Value = "Invoice Number"
    End With
    
    reconciledTransDateColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Transaction Date", lookat:=xlWhole).Column
    reconciledInvColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoiced", lookat:=xlWhole).Column
    reconciledInvNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Invoice Number", lookat:=xlWhole).Column
    reconciledReceiptNumColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Receipt Num", lookat:=xlWhole).Column
    scTicketNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Ticket Number", lookat:=xlWhole).Column
    scInvoiceAmountColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice Total", lookat:=xlWhole).Column
    scInvoiceNumberColumn = Sheets(scWorksheet).UsedRange.Find(what:="Invoice #", lookat:=xlWhole).Column
    reconciledPrimaryQuantityColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Primary Quantity", lookat:=xlWhole).Column
    reconciledPONumberColumn = Sheets(reconciledSheet).UsedRange.Find(what:="Po Number", lookat:=xlWhole).Column
    
    For p = invsheetlr To 2 Step -1
        'check invoice info to verify the receipt has been reconciled.  also checks that invoice is not
        'actually a credit memo (value>=0)
        If Not Application.WorksheetFunction.IsNA(Application.Match(Sheets(invworksheet) _
        .Cells(p, invReceiptNumColumn), Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) And _
        Sheets(invworksheet).Cells(p, invInvoiceAmountColumn).Value >= 0 Then

        matchedRow = Application.Match(Sheets(invworksheet). _
        Cells(p, invReceiptNumColumn).Value, Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)

        With Sheets(reconciledSheet)
            .Cells(matchedRow, reconciledInvNumColumn).Value = Sheets(invworksheet).Cells(p, invInvoiceNumberColumn).Value
        End With
        End If
    Next
    
    'if no invoice for receipt, insert red X in invoiced? column on reconciled receipts sheet, else insert green check
    For q = 2 To reconciledLR
        If Sheets(reconciledSheet).Cells(q, reconciledInvNumColumn).Value = "" Then
                    
        With Sheets(reconciledSheet).Cells(q, 1)
            .Value = ChrW(10006)
            .Font.Bold = True
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
    
    Dim recInvInvoiceTypeColumn As Long
    recInvInvoiceTypeColumn = Sheets(reconciledInvoices).UsedRange.Find(what:="Invoice Type", lookat:=xlWhole).Column
    
    'lookup receipt ticket number for reconciled invoices table
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
        
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
            With Sheets(reconciledInvoices)
                .Cells(q, invoiceVerifiedColumn).Value = "CM"
                .Cells(q, invoiceVerifiedColumn).Font.Bold = True
                .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
                .Cells(q, receiptVerifiedColumn).Value = "Receipt Not Reconciled"
                .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
                .Cells(q, receiptVerifiedColumn).Font.Bold = True
            End With
        Else

        With Sheets(reconciledInvoices)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006) '"X"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Value = "Receipt Not Reconciled"
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
        End With
        End If
        
        ElseIf Application.WorksheetFunction.IsNA(Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) Then
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
            With Sheets(reconciledInvoices)
                .Cells(q, invoiceVerifiedColumn).Value = "CM"
                .Cells(q, invoiceVerifiedColumn).Font.Bold = True
                .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
                .Cells(q, receiptVerifiedColumn).Value = "Ticket Not in ScrapConnect"
                .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
                .Cells(q, receiptVerifiedColumn).Font.Bold = True
            End With
        Else
            With Sheets(reconciledInvoices)
                .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006) '"X"
                .Cells(q, invoiceVerifiedColumn).Font.Bold = True
                .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
                .Cells(q, receiptVerifiedColumn).Value = "Ticket Not in ScrapConnect"
                .Cells(q, receiptVerifiedColumn).Font.Color = RGB(255, 0, 0)
                .Cells(q, receiptVerifiedColumn).Font.Bold = True
            End With
        End If
        
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceNumberColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceNumberColumn).Value Then
        
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004) '(CHECK MARK)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = "CM"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
            .Cells(q, recInvInvoiceNumberColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceNumberColumn).Font.Bold = True
            .Cells(q, recInvInvoiceNumberColumn).Interior.Color = RGB(255, 255, 0)
        End With
        Else
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004) '(CHECK MARK)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = ChrW(10006) '"X"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceNumberColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceNumberColumn).Font.Bold = True
            .Cells(q, recInvInvoiceNumberColumn).Interior.Color = RGB(255, 255, 0)
        End With
        End If
        
        ElseIf Application.Index(Sheets(scWorksheet).Columns(scInvoiceAmountColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(scWorksheet).Columns(scTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceAmountColumn).Value Then
        
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = "CM"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
            .Cells(q, recInvInvoiceAmountColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceAmountColumn).Font.Bold = True
            .Cells(q, recInvInvoiceAmountColumn).Interior.Color = RGB(255, 255, 0)
        End With
        Else
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
        End If
        
        ElseIf Application.Index(Sheets(reconciledSheet).Columns(reconciledPrimaryQuantityColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvTicketNumberColumn).Value, _
        Sheets(reconciledSheet).Columns(reconciledTicketNumberColumn), 0)) <> _
        Sheets(reconciledInvoices).Cells(q, recInvInvoiceQtyColumn).Value Then
        
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
        With Sheets(reconciledInvoices)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = "CM"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
            .Cells(q, recInvInvoiceQtyColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoiceQtyColumn).Font.Bold = True
            .Cells(q, recInvInvoiceQtyColumn).Interior.Color = RGB(255, 255, 0)
        End With
        Else
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
        End If
        
        ElseIf Sheets(reconciledInvoices).Cells(q, recInvInvoicePOColumn).Value <> _
        Application.Index(Sheets(reconciledSheet).Columns(reconciledPONumberColumn), _
        Application.Match(Sheets(reconciledInvoices).Cells(q, recInvReceiptNumberColumn), _
        Sheets(reconciledSheet).Columns(reconciledReceiptNumColumn), 0)) Then
        
        If Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value <> "Credit Memo" Then
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
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, invoiceVerifiedColumn).Value = "CM"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
            .Cells(q, recInvInvoicePOColumn).Font.Color = RGB(255, 0, 0)
            .Cells(q, recInvInvoicePOColumn).Font.Bold = True
            .Cells(q, recInvInvoicePOColumn).Interior.Color = RGB(255, 255, 0)
        End With
        End If
        
        ElseIf Sheets(reconciledInvoices).Cells(q, recInvInvoiceTypeColumn).Value = "Credit Memo" Then
        
        With Sheets(reconciledInvoices)
            .Cells(q, invoiceVerifiedColumn).Value = "CM"
            .Cells(q, invoiceVerifiedColumn).Font.Bold = True
            .Cells(q, invoiceVerifiedColumn).Font.Color = RGB(0, 0, 255)
            .Cells(q, receiptVerifiedColumn).Value = ChrW(10004)
            .Cells(q, receiptVerifiedColumn).Font.Color = RGB(0, 255, 0)
            .Cells(q, receiptVerifiedColumn).Font.Bold = True
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
    
    If UserForm1.OptionButton1.Value = "True" Then
    Call printSummary
    End If
       
    're-enable excel screen updating
    Sheets(reconciledInvoices).Visible = xlSheetHidden
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    'enable/disable userform buttons
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
