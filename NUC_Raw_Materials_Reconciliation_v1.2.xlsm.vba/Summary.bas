Sub printSummary()
        
    On Error Resume Next
    
    invworksheet = "Invoice Report"
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    
'*******************************************************  RESULTS SUMMARY ****************************************
    
    'This section calculates and prints summary data on "Home" page
    Dim lastRowMissingOracleSheet As Long
    Dim lastRowMissingSCSheet As Long
    
    
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
    
    
    varLR = Sheets("Void and Return to Vendor").UsedRange.Rows _
    (Sheets("Void and Return to Vendor").UsedRange.Rows.Count).Row
        
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
    
    reconciledLR = Sheets(reconciledSheet).UsedRange.Rows.Count
   
    Dim reconciledRange As Range
    Set reconciledRange = Sheets(reconciledSheet).Range(Sheets(reconciledSheet).Cells(1, 1), Sheets(reconciledSheet) _
    .Cells(reconciledLR, reconciledLC))
        
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
        
        With Sheets(i).UsedRange
            .Rows(1).Font.Bold = True
            .Columns.AutoFit
        End With
        
        ActiveWindow.FreezePanes = False
        Sheets(i).Rows(2).Select
        ActiveWindow.FreezePanes = True
             
        Sheets(i).Rows(1).EntireRow.Insert
        With Worksheets(i)
            .Hyperlinks.Add Anchor:=.Range("A1"), Address:="", SubAddress:="'" & Worksheets(1).Name _
            & "'" & "!A1", TextToDisplay:="Home"
            With .Range("A1")
                .Font.Bold = True
                .Font.Color = RGB(214, 214, 214)
                .Font.Size = 16
                .Font.Name = "arial"
                .RowHeight = 30
                .ColumnWidth = 15
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = RGB(0, 15, 230)
            End With
        End With
    Next i
    
    
    a = Application.Count(Sheets(scWorksheet).Range(Sheets(scWorksheet) _
    .Cells(scStartingRow, scColumn), Sheets(scWorksheet).Cells(scSheetLR, scColumn)))
    b = Application.Count(Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet) _
    .Cells(ebsStartingRow, ebsColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn)))
    c = Application.Count(Sheets("Receipts Missing From SC") _
    .Range(Sheets("Receipts Missing From SC").Cells(ebsStartingRow, ebsColumn), _
    Sheets("Receipts Missing From SC").Cells(lastRowMissingSCSheet, ebsColumn)))
    d = Application.Count(Sheets("Receipts Missing From Oracle") _
    .Range(Sheets("Receipts Missing From Oracle").Cells(scStartingRow, scColumn), _
    Sheets("Receipts Missing From Oracle").Cells(lastRowMissingOracleSheet, scColumn)))
    e = Application.Count(Sheets("Void and Return to Vendor").Range("A1:A" & varLR))
    f = Application.Count(Sheets("Weight Discrepancies").Columns(1))
    g = Application.Count(Sheets("Pending Receipts").Range("A1:A" & Sheets("Pending Receipts") _
    .UsedRange.Rows(Sheets("Pending Receipts").UsedRange.Rows.Count).Row))
    If UserForm1.OptionButton1.Value = "True" Then
    i = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A1:A" & Sheets(reconciledSheet) _
    .UsedRange.Rows.Count), ChrW(10006))
    j = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A1:A" & Sheets(reconciledSheet) _
    .UsedRange.Rows.Count), "ERROR")
    End If
    
    
'    g = Application.WorksheetFunction.CountIf(reconcileRange.Columns(Sheets(reconciledSheet) _
    .UsedRange.Find(what:="Invoice Total").Column), ">0")
    h = reconciledLR

    'display summary on Reconciliation page
'    With Sheets(1)
'        .Range("k2").Value = b - 1

    With Worksheets(1)
        .Hyperlinks.Add Anchor:=.Range("K2"), _
        Address:="", SubAddress:="'" & Worksheets(ebsWorksheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (b)
        .Range("l2").Value = "Total Oracle Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K3"), _
        Address:="", SubAddress:="'" & Worksheets(scWorksheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (a)
'        .Range("k3").Value = a - 1
        .Range("l3").Value = "Total ScrapConnect Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K4"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (h)
'        .Range("k4").Value = h - 1
        .Range("l4").Value = "Reconciled Receipts"
        
        If UserForm1.OptionButton1.Value = "True" Then
        .Hyperlinks.Add Anchor:=.Range("K5"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (i)
        .Range("l5").Value = "Uninvoiced Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K6"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (j)
        .Range("l6").Value = "Invoices with Errors"
        End If
        
        .Hyperlinks.Add Anchor:=.Range("K7"), _
        Address:="", SubAddress:="'" & Worksheets("Pending Receipts").Name & "'" & "!A1", _
        TextToDisplay:="'" & (g)
        .Range("l7").Value = "Pending Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K8"), _
        Address:="", SubAddress:="'" & Worksheets("Receipts Missing From SC").Name & "'" & "!A1", _
        TextToDisplay:="'" & (c)
'        .Range("k6").Value = c - 1
        .Range("l8").Value = "Receipts missing from ScrapConnect"
        
        .Hyperlinks.Add Anchor:=.Range("K9"), _
        Address:="", SubAddress:="'" & Worksheets("Receipts Missing From Oracle").Name & "'" & "!A1", _
        TextToDisplay:="'" & (d)
'        .Range("k7").Value = d - 1
        .Range("l9").Value = "Receipts missing from Oracle"
        
        .Hyperlinks.Add Anchor:=.Range("K10"), _
        Address:="", SubAddress:="'" & Worksheets("Void and Return to Vendor").Name & "'" & "!A1", _
        TextToDisplay:="'" & (e)
'        .Range("k8").Value = e - 2
        .Range("l10").Value = "Voided and Return to Vendor receipts"
       
       .Hyperlinks.Add Anchor:=.Range("K11"), _
        Address:="", SubAddress:="'" & Worksheets("Weight Discrepancies").Name & "'" & "!A1", _
        TextToDisplay:="'" & (f)
'        .Range("k9").Value = f - 1
        .Range("l11").Value = "Weight discrepancies"
    End With
     
    If UserForm1.OptionButton1.Value = "False" Then
    Sheets(1).Range("K5:L6").Delete Shift:=xlUp
    End If

                With Sheets(1).Range("K1")
                    .Value = "Summary - " & Format(Now, "mm/dd/yyyy HH:mm")
                    .Font.Size = 24
                    .Font.Bold = True
                    .Font.Name = "arial"
                    .Rows(1).AutoFit
                End With
                With Sheets(1).Range("k2:k11")
                    .Font.Bold = True
                    .Font.ColorIndex = 3
                End With
                With Sheets(1).Range("k2:l" & Sheets(1).Cells(Rows.Count, "L").End(xlUp).Row)
                    .Font.Size = 15
                    .Font.Bold = True
                    .Font.Name = "arial"
                    .Rows.AutoFit
                    .BorderAround ColorIndex:=0, Weight:=xlThick
                    .Columns.AutoFit
                End With
    
    Sheets(reconciledSheet).Visible = xlSheetHidden
    Sheets("Pending Receipts").Visible = xlSheetHidden
    Sheets("Weight Discrepancies").Visible = xlSheetHidden
    Sheets("Void and Return to Vendor").Visible = xlSheetHidden
    Sheets("Receipts Missing From Oracle").Visible = xlSheetHidden
    Sheets("Receipts Missing From SC").Visible = xlSheetHidden
    If UserForm1.OptionButton1.Value = "True" Then
    Sheets("Unmatched Invoices").Visible = xlSheetHidden
    End If
    
End Sub