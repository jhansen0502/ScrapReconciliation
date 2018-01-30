Public Function indexMatch(ByVal ebsfield As String, _
ByVal scfield As String, ByVal returnfield As String, _
ByVal printColumn As Long, ByVal sheetName As String) As Variant
    
    'This function runs INDEX(MATCH) excel function to perform
    'lookups.  Its purpose is equivalent to VLOOKUP but is
    'more resilient if users rearrange columns in reference sheets
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
    Dim j As Long
    Dim errorWorksheet As String
    Dim errorLR As Long
    Dim errorLC As Long
    Dim errorRange As Range
    
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
    
    If sheetName = "sc" Then
        errorWorksheet = "Receipts Missing From Oracle"
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
        

        For j = (scRow + 1) To scSheetLR
            If Application.WorksheetFunction.IsNA(Application.Match _
                (Sheets(scWorksheet).Cells(j, scColumn), lookupRange, 0)) Then
                
                Sheets(scWorksheet).Rows(j).EntireRow.Copy
                Sheets(errorWorksheet).Range("A" & Rows.Count) _
                .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            End If
        Next
        
        errorLR = Sheets(errorWorksheet).UsedRange.Rows _
        (Sheets(errorWorksheet).UsedRange.Rows.Count).Row
        errorLC = Sheets(errorWorksheet).UsedRange.Columns _
        (Sheets(errorWorksheet).UsedRange.Columns.Count).Column
        Set errorRange = Sheets(errorWorksheet).Range(Sheets _
        (errorWorksheet).Cells(1, 1), Sheets(errorWorksheet).Cells _
        (errorLR, errorLC))
        
        With errorRange
            .Borders.LineStyle = xlContinuous
        End With

     ElseIf sheetName = "ebs" Then
        errorWorksheet = "Receipts Missing From SC"
        Set ebsFieldCell = Sheets(ebsWorksheet).Rows(ebsStartingRow).Find(what:=ebsfield)
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

        For j = (ebsRow + 1) To ebsSheetLR
            If Application.WorksheetFunction.IsNA(Application.Match _
                (Sheets(ebsWorksheet).Cells(j, ebsColumn), lookupRange, 0)) Then
                
                Sheets(ebsWorksheet).Rows(j).EntireRow.Copy
                Sheets(errorWorksheet).Range("A" & Rows.Count) _
                .End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            End If
        Next
        
        errorLR = Sheets(errorWorksheet).UsedRange.Rows _
        (Sheets(errorWorksheet).UsedRange.Rows.Count).Row
        errorLC = Sheets(errorWorksheet).UsedRange.Columns _
        (Sheets(errorWorksheet).UsedRange.Columns.Count).Column
        Set errorRange = Sheets(errorWorksheet).Range(Sheets _
        (errorWorksheet).Cells(1, 1), Sheets(errorWorksheet).Cells _
        (errorLR, errorLC))
        
        With errorRange
            .Borders.LineStyle = xlContinuous
        End With
      End If

End Function
