Public Function indexMatchComparison(ByVal ebsfield As String, _
ByVal scfield As String, ByVal returnfield As String, _
ByVal printColumn As Long, ByVal sheetName As String, _
ByVal comparisonField As String) As Variant

    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    ebsfield = "S C Tkt"
    scfield = "Ticket Number"
    ebsStartingRow = Sheets(ebsWorksheet).UsedRange.Find(what:=ebsfield).Row
    scStartingRow = Sheets(scWorksheet).UsedRange.Find(what:=scfield).Row
    
    Dim ebsFieldCell As Range
    Dim ebsColumn As Integer
    Dim ebsRow As Integer
    Dim scFieldCell As Range
    Dim scColumn As Integer
    Dim scRow As Integer
    Dim lookupRange As Range
    Dim returnCell As Range
    Dim returnColumn As Integer
    Dim returnRow As Integer
    Dim returnRange As Range
    Dim comparisonFieldCell As Range
    Dim comparisonRange As Range
    Dim comparisonColumn As Integer
    Dim thisRange As Range
    Dim j As Long
    Dim ebsWeightCell As Range
    Dim ebsWeightColumn As Long
    Dim scWeightCell As Range
    Dim scWeightColumn As Long
    Dim nextRow As Long
    
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
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows _
    (Sheets(reconciledSheet).UsedRange.Rows.Count).Row
    reconciledLC = Sheets(reconciledSheet).UsedRange.Columns _
    (Sheets(reconciledSheet).UsedRange.Columns.Count).Column
    Set reconcileRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), _
    Sheets(reconciledSheet).Cells(reconciledLR, reconciledLC))
    
    If sheetName = "sc" Then
        Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
        ebsColumn = ebsFieldCell.Column
        ebsRow = ebsFieldCell.Row
        Set scFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=scfield)
        scColumn = scFieldCell.Column
        scRow = scFieldCell.Row
        Set returnCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=returnfield)
        returnColumn = returnCell.Column
        returnRow = returnCell.Row
        Set returnRange = Range(Sheets(ebsWorksheet).Cells(returnRow, _
        returnColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, returnColumn))
        Set lookupRange = Range(Sheets(ebsWorksheet).Cells(ebsRow, ebsColumn), _
        Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn))
        Set comparisonFieldCell = Sheets(scWorksheet).Rows(scStartingRow).Find(what:=comparisonField)
        comparisonColumn = comparisonFieldCell.Column
        Set comparisonRange = Range(Sheets(scWorksheet).Cells(scStartingRow, comparisonColumn), Sheets _
        (scWorksheet).Cells(scSheetLR, comparisonColumn))
        Set thisRange = Range(Sheets(scWorksheet).Cells(scStartingRow, scColumn), Sheets(scWorksheet) _
        .Cells(scSheetLR, scColumn))
        
        nextRow = Sheets("Weight Discrepancies").Range("A" & Rows.Count).End(xlUp).Row
        
        For j = (scRow + 1) To scSheetLR
            If Not (Application.WorksheetFunction.IsNA(Application.Match _
            (Sheets(scWorksheet).Cells(j, scColumn), lookupRange, 0))) Then
        
                If Application.WorksheetFunction.Index(returnRange, _
                    Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
                    lookupRange, 0)) <> _
                    Application.WorksheetFunction.Index(comparisonRange, _
                    Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
                    thisRange, 0)) Then
                
                    ebsweightrow = Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
                    lookupRange, 0)
                                        
                    With Sheets(ebsWorksheet)
                        .Cells(ebsweightrow, ebsColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("A" & nextRow)
                        .Cells(ebsweightrow, bCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("B" & nextRow)
                        .Cells(ebsweightrow, cCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("C" & nextRow)
                        .Cells(ebsweightrow, dCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("D" & nextRow)
                        .Cells(ebsweightrow, eCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("E" & nextRow)
                        .Cells(ebsweightrow, fCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("F" & nextRow)
                        .Cells(ebsweightrow, gCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("G" & nextRow)
                        .Cells(ebsweightrow, hCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("H" & nextRow)
                        .Cells(ebsweightrow, iCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("I" & nextRow)
                        .Cells(ebsweightrow, aCellColumn).Copy Destination:=Sheets("Weight Discrepancies").Range("J" & nextRow)
                    End With
                    
                    nextRow = nextRow + 1
                    
                    Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(scWorksheet).Cells(j, scColumn).Value))
                    recordToDelete = recordCellToDelete.Row
                    Sheets(reconciledSheet).Rows(recordToDelete).Delete
                End If
            End If
        Next j
    Else
        Set ebsFieldCell = Sheets(ebsWorksheet).Range(Cells(ebsStartingRow, 1), Cells(ebsSheetLR, ebsSheetLC)) _
        .Find(what:=ebsfield)
        ebsColumn = ebsFieldCell.Column
        ebsRow = ebsFieldCell.Row
        Set scFieldCell = Range(Sheets(scWorksheet).Cells(scStartingRow, 1), Sheets(scWorksheet).Cells(scSheetLR, scSheetLC)) _
        .Find(what:=scfield)
        scColumn = scFieldCell.Column
        scRow = scFieldCell.Row
        Set returnCell = Range(Sheets(scWorksheet).Cells(scStartingRow, 1), Sheets(scWorksheet).Cells(scSheetLR, scSheetLC)) _
        .Find(what:=returnfield)
        returnColumn = returnCell.Column
        returnRow = returnCell.Row
        Set returnRange = Range(Sheets(scWorksheet).Cells(returnRow, returnColumn), _
        Sheets(scWorksheet).Cells(scSheetLR, returnColumn))
        Set lookupRange = Range(Sheets(scWorksheet).Cells(scRow, scColumn), _
        Sheets(scWorksheet).Cells(scSheetLR, scColumn))
        Set comparisonFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=comparisonField)
        comparisonColumn = comparisonFieldCell.Column
        Set comparisonRange = Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, 1), Sheets _
        (ebsWorksheet).Cells(ebsSheetLR, comparisonColumn))
        Set thisRange = Range(Sheets(ebsWorksheet).Cells(ebsStartingRow, ebsColumn), Sheets(ebsWorksheet) _
        .Cells(ebsSheetLR, ebsColumn))


        For j = (ebsRow + 1) To ebsSheetLR
            If Not (Application.WorksheetFunction.IsNA(Application.Match _
                (Sheets(ebsWorksheet).Cells(j, ebsColumn), lookupRange, 0))) Then
                
                If Application.WorksheetFunction.Index(returnRange, _
                    Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
                    lookupRange, 0)) <> _
                    Application.WorksheetFunction.Index(comparisonRange, _
                    Application.Match(Sheets(scWorksheet).Cells(j, scColumn), _
                    thisRange, 0)) Then
                
                    Sheets(ebsWorksheet).Rows(j).EntireRow.Copy
                    Sheets("Weight Discrepancies").Range("A" & Rows.Count) _
                    .End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormats & xlPasteValues
                    
                    Set recordCellToDelete = reconcileRange.Find(what:=(Sheets(scWorksheet).Cells(j, scColumn).Value))
                    recordToDelete = recordCellToDelete.Row
                    Sheets(reconciledSheet).Rows(recordToDelete).Delete
                End If
            End If
        Next
    End If

End Function
