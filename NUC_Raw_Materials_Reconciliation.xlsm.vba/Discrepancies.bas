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
    Sheets.Add(After:=Sheets(reconciledSheet)).Name = "Receipts Missing From Oracle"
    Sheets(scWorksheet).Rows(scStartingRow).EntireRow.Copy
    Sheets("Receipts Missing From Oracle").Rows(1).PasteSpecial xlPasteValues
    
    Sheets.Add(After:=Sheets("Receipts Missing From Oracle")).Name = "Receipts Missing From SC"
    Sheets(ebsWorksheet).Rows(ebsStartingRow).EntireRow.Copy
    Sheets("Receipts Missing From SC").Rows(1).PasteSpecial xlPasteValues
      
    Sheets.Add(After:=Sheets("Receipts Missing From SC")).Name = "Void and RTV"
    Sheets(ebsWorksheet).Rows(ebsStartingRow).EntireRow.Copy Destination:=Sheets("Void and RTV").Range("A1")
    
    Sheets.Add(After:=Sheets("Void and RTV")).Name = "Weight Discrepancies"
    Set aCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Bill of Lading")
    aCellColumn = aCell.Column
    Set bCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Date")
    bCellColumn = bCell.Column
    Set cCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Po Number")
    cCellColumn = cCell.Column
    Set dCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Number")
    dCellColumn = dCell.Column
    Set eCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Item Description")
    eCellColumn = eCell.Column
    Set fCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Ship Mode")
    fCellColumn = fCell.Column
    Set gCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Gross Weight")
    gCellColumn = gCell.Column
    Set hCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Primary Quantity")
    hCellColumn = fCell.Column
    Set iCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Receipt Num")
    iCellColumn = iCell.Column
    
    With Sheets("Weight Discrepancies")
        Range("A" & 1).Value = ebsfield
        Range("B" & 1).Value = bCell.Value
        Range("C" & 1).Value = cCell.Value
        Range("D" & 1).Value = dCell.Value
        Range("E" & 1).Value = eCell.Value
        Range("F" & 1).Value = fCell.Value
        Range("G" & 1).Value = gCell.Value
        Range("H" & 1).Value = hCell.Value
        Range("I" & 1).Value = iCell.Value
        Range("J" & 1).Value = aCell.Value
    End With
        
    Call indexMatch("S C Tkt", "Ticket Number", "S C Tkt", 1, "sc")
    Call indexMatch("S C Tkt", "Ticket Number", "Ticket Number", 1, "ebs")
    
    'Find "Status" columns from both source file sheets
    Dim ebsStatusColumn As Long
    Dim ebsStatusCell As Range
    Dim scStatusColumn As Long
    Dim scStatusCell As Range
    
    Set ebsStatusCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:="Transaction Type")
    ebsStatusColumn = ebsStatusCell.Column
    Set scStatusCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:="Status")
    scStatusColumn = scStatusCell.Column
    
    'Find used range of "Reconciled Receipts" sheet
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    'This loop pulls all "RTV" records from the ebs source file
    'and copies to "Void and RTV" sheet.
    For i = 2 To ebsSheetLR
        If Sheets(ebsWorksheet).Cells(i, ebsStatusColumn).Value = "RETURN TO VENDOR" Then

            Sheets(ebsWorksheet).Rows(i).EntireRow.Copy
            Sheets("Void and RTV").Range("A" & Rows.Count) _
            .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            
            'Removes any RTV receipts from the "Reconciled Receipts" page
            Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(ebsWorksheet).Cells(i, ebsColumn).Value))
            recordToDelete = recordCellToDelete.Row
            Sheets(reconciledSheet).Rows(recordToDelete).Delete

        End If
    Next
    
    'This loop pulls all "Void" records from the SC source file
    'and copies to "Void and RTV" sheet.
    Sheets(scWorksheet).Rows(scStartingRow).EntireRow.Copy
    Sheets("Void and RTV").Range("A" & Rows.Count).End(xlUp).Offset(2, 0).PasteSpecial xlPasteValues
    Sheets("Void and RTV").Range("A" & Rows.Count).End(xlUp).Rows.EntireRow.Font.Bold = True
    For i = 2 To scSheetLR
        If Sheets(scWorksheet).Cells(i, scStatusColumn).Value = "Void" Then
            
            Sheets(scWorksheet).Rows(i).EntireRow.Copy
            Sheets("Void and RTV").Range("A" & Rows.Count) _
            .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            
            'Removes any RTV receipts fromt the "Reconciled Receipts" page
            Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(scWorksheet).Cells(i, scColumn).Value))
            recordToDelete = recordCellToDelete.Row
            Sheets(reconciledSheet).Rows(recordToDelete).Delete

        End If
    Next
    
    
    'Formatting for "Void and RTV" sheet
    Dim voidLR As Long
    Dim voidLC As Long
    Dim voidRange As Range
    Dim completedDateCell As Range
    Dim completedDateColumn As Long
    Dim completedDateRow As Long
    
    voidLR = Sheets("Void and RTV").UsedRange.Rows _
        (Sheets("Void and RTV").UsedRange.Rows.Count).Row
    voidLC = Sheets("void and RTV").UsedRange.Columns _
        (Sheets(scWorksheet).UsedRange.Columns.Count).Column
    Set voidRange = Sheets("Void and RTV").Range(Sheets("Void and RTV").Cells _
    (1, 1), Sheets("Void and RTV").Cells(voidLR, voidLC))
    
    
    Set completedDateCell = Sheets("Void and RTV").Range(Sheets("Void and RTV") _
        .Cells(1, 1), Sheets("Void and RTV").Cells(voidLR, voidLC)).Find(what:="Completed Date")
    completedDateRow = completedDateCell.Row
    completedDateColumn = completedDateCell.Column
    
    With Sheets("Void and RTV").Range(Sheets("Void and RTV").Cells(completedDateRow, completedDateColumn), _
        Sheets("Void and RTV").Cells(voidLR, completedDateColumn))
        .NumberFormat = "mm/dd/yyyy"
    End With
        
    With voidRange
        .Borders.LineStyle = xlContinuous
    End With
        
    Dim wdLR As Long
    Dim wdLC As Long
    
    'Call to function that finds matched tickets with weight discrepancies
    Call indexMatchComparison("S C Tkt", "Ticket Number", "Primary Quantity", 1, "sc", "Net Weight")
    
    'Weight Lookups for Weight Discrepancies sheet.  The function above finds the
    'ticket number.  This block parses the respective weights from each source
    'file and finds the difference.
    Worksheets("Weight Discrepancies").Activate
    Sheets("Weight Discrepancies").Columns("B:D").Insert Shift:=xlToRight, _
        CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Weight Discrepancies").Range("B1").Value = "ScrapConnect Weight"
    Sheets("Weight Discrepancies").Range("C1").Value = "Oracle Weight"
    Sheets("Weight Discrepancies").Range("D1").Value = "Weight Differential"
    
    wdLR = Sheets("Weight Discrepancies").UsedRange.Rows _
    (Sheets("Weight Discrepancies").UsedRange.Rows.Count).Row

    Dim lookupRange As Range
    Dim returnCell As Range
    Dim returnColumn As Integer
    Dim returnRow As Integer
    Dim returnRange As Range
    Dim j As Long
    Dim errorWorksheet As String
    
    ebsfield = "S C Tkt"
    scfield = "Ticket Number"
    
    Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
    ebsColumn = ebsFieldCell.Column
    ebsRow = ebsFieldCell.Row
    Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
    scColumn = scFieldCell.Column
    scRow = scFieldCell.Row
    
    For j = 2 To 3
        If Sheets("Weight Discrepancies").Cells(1, j) = "ScrapConnect Weight" Then
            returnfield = "Net Weight"
            Set returnCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=returnfield)
            returnColumn = returnCell.Column
            returnRow = returnCell.Row
            Set returnRange = Range(Sheets(scWorksheet).Cells(returnRow, _
            returnColumn), Sheets(scWorksheet).Cells(scSheetLR, returnColumn))
            Set lookupRange = Range(Sheets(scWorksheet).Cells(scRow, scColumn), _
            Sheets(scWorksheet).Cells(scSheetLR, scColumn))
            
                For i = 2 To wdLR
            
                    Sheets("Weight Discrepancies").Range("B" & i).Value = _
                    Application.Index(returnRange, _
                    Application.Match(Sheets("Weight Discrepancies").Range("A" & i).Value, _
                    lookupRange, 0))
                Next
        Else
            returnfield = "Primary Quantity"
            Set returnCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=returnfield)
            returnColumn = returnCell.Column
            returnRow = returnCell.Row
            Set returnRange = Range(Sheets(ebsWorksheet).Cells(returnRow, _
            returnColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, returnColumn))
            Set lookupRange = Range(Sheets(ebsWorksheet).Cells(ebsRow, ebsColumn), _
            Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn))
            
                For i = 2 To wdLR
            
                    Sheets("Weight Discrepancies").Range("C" & i).Value = _
                    Application.Index(returnRange, _
                    Application.Match(Sheets("Weight Discrepancies").Range("A" & i).Value, _
                    lookupRange, 0))
                Next
        End If
    Next
    
    For i = 2 To wdLR
        If Sheets("Weight Discrepancies").Range("B" & i).Value > Sheets("Weight Discrepancies") _
        .Range("C" & i).Value Then
        Sheets("Weight Discrepancies").Range("D" & i).Value = (Sheets("Weight Discrepancies") _
            .Range("B" & i).Value - Sheets("Weight Discrepancies").Range("C" & i).Value)
        Else
        Sheets("Weight Discrepancies").Range("D" & i).Value = (Sheets("Weight Discrepancies") _
        .Range("C" & i).Value - Sheets("Weight Discrepancies").Range("D" & i).Value)
        End If
    Next
    
    With Range(Sheets("Weight Discrepancies").Cells(1, 2), Sheets("Weight Discrepancies").Cells(wdLR, 4))
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
    wdLC = Sheets("Weight Discrepancies").UsedRange.Columns _
    (Sheets("Weight Discrepancies").UsedRange.Columns.Count).Column
    
    Set wdrange = Sheets("Weight Discrepancies").Range(Sheets("Weight Discrepancies").Cells _
    (1, 1), Sheets("Weight Discrepancies").Cells(wdLR, wdLC))
    
    With wdrange
        .Borders.LineStyle = xlContinuous
    End With
    
    '*******************************************************  RESULTS SUMMARY ****************************************
    
    'This section calculates and prints summary data on "Home" page
    Dim lastRowMissingOracleSheet As Long
    Dim lastRowMissingSCSheet As Long
    
    varLR = Sheets("Void and RTV").UsedRange.Rows _
    (Sheets("Void and RTV").UsedRange.Rows.Count).Row
        
    wdLR = Sheets("Weight Discrepancies").UsedRange.Rows _
    (Sheets("Weight Discrepancies").UsedRange.Rows.Count).Row
        
    lastRowMissingOracleSheet = Sheets("Receipts Missing From Oracle").UsedRange.Rows _
    (Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count).Row
        
    lastRowMissingSCSheet = Sheets("Receipts Missing From SC").UsedRange.Rows _
    (Sheets("Receipts Missing From SC").UsedRange.Rows.Count).Row
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    a = Application.WorksheetFunction.CountA(Sheets(scWorksheet).Range(Sheets(scWorksheet) _
    .Cells(scStartingRow, scColumn), Sheets(scWorksheet).Cells(scSheetLR, scColumn)))
    b = Application.WorksheetFunction.CountA(Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet) _
    .Cells(ebsStartingRow, ebsColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn)))
    c = Application.WorksheetFunction.CountA(Sheets("Receipts Missing From SC") _
    .Range(Sheets("Receipts Missing From SC").Cells(ebsStartingRow, ebsColumn), _
    Sheets("Receipts Missing From SC").Cells(lastRowMissingSCSheet, ebsColumn)))
    d = Application.WorksheetFunction.CountA(Sheets("Receipts Missing From Oracle") _
    .Range(Sheets("Receipts Missing From Oracle").Cells(scStartingRow, scColumn), _
    Sheets("Receipts Missing From Oracle").Cells(lastRowMissingOracleSheet, scColumn)))
    e = Application.WorksheetFunction.CountA(Sheets("Void and RTV").Range("A1:A" & varLR))
    f = Application.WorksheetFunction.CountA(Sheets("Weight Discrepancies").Range("A1:A" & wdLR))
    g = Application.WorksheetFunction.CountA(reconcileRange.Columns(10))
    h = reconciledLR

    'display summary on Reconciliation page
    With Sheets(1)
        .Range("k2").Value = b - 1
        .Range("l2").Value = "Total Oracle Receipts"
        .Range("k3").Value = a - 1
        .Range("l3").Value = "Total ScrapConnect Receipts"
        .Range("k4").Value = h - 1
        .Range("l4").Value = "Reconciled Receipts"
        .Range("k5").Value = g - 1
        .Range("l5").Value = "Invoiced Receipts"
        .Range("k6").Value = c - 1
        .Range("l6").Value = "Receipts missing from ScrapConnect"
        .Range("k7").Value = d - 1
        .Range("l7").Value = "Receipts missing from Oracle"
        .Range("k8").Value = e - 2
        .Range("l8").Value = "Voided and RTV receipts"
        .Range("k9").Value = f - 1
        .Range("l9").Value = "Weight discrepancies"
        
                With Sheets(1).Range("K1")
                    .Value = "Summary - " & Format(Now, "mm/dd/yyyy HH:mm")
                    .Font.Size = 24
                    .Font.Bold = True
                    .Font.Name = "arial"
                    .Rows(1).AutoFit
                End With
                With Sheets(1).Range("k2:k9")
                    .Font.Bold = True
                    .Font.ColorIndex = 3
                End With
                With Sheets(1).Range("k2:l9")
                    .Font.Size = 18
                    .Font.Name = "arial"
                    .Rows.AutoFit
                    .BorderAround ColorIndex:=0, Weight:=xlThick
                    .Columns.AutoFit
                End With

    End With
    
    Dim sheetRange As Range
    Dim sheetlr As Long
    Dim sheetlc As Long
    
    For i = 2 To Sheets.Count
        Sheets(i).Activate
        
        sheetlr = Sheets(i).UsedRange.Rows _
        (Sheets(i).UsedRange.Rows.Count).Row
        
        sheetlc = Sheets(i).UsedRange.Rows _
        (Sheets(i).UsedRange.Rows.Count).Row

        Set sheetRange = Sheets(i).Range(Sheets(i).Cells(1, 1), Sheets(i).Cells(sheetlr, sheetlc))
                            
        For j = 1 To sheetlc
            If InStr(1, Sheets(i).Cells(1, j).Value, "Date") <> 0 Then
            Sheets(i).Columns(j).NumberFormat = "mm/dd/yyyy"
            End If
        Next j
        
        With sheetRange
            .Rows(1).Font.Bold = True
            .Columns.AutoFit
        End With
        
        ActiveWindow.FreezePanes = False
        Sheets(i).Rows(2).Select
        ActiveWindow.FreezePanes = True
             
    Next i
    
    'Hides all sheets.  Users will export to view results
    Sheets("Weight Discrepancies").Visible = xlSheetHidden
    Sheets("Void and RTV").Visible = xlSheetHidden
    Sheets("Receipts Missing From Oracle").Visible = xlSheetHidden
    Sheets("Receipts Missing From SC").Visible = xlSheetHidden
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    With UserForm1
        .findDiscrepancies.Enabled = False
        .findDiscrepancies.BackColor = RGB(214, 214, 214)
        .ExportToNewWB.Enabled = True
        .ExportToNewWB.BackColor = RGB(0, 238, 0)
    End With

    Sheets(1).Activate

End Sub