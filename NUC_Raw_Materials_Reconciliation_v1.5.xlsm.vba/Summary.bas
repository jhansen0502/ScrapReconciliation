Sub printSummary()
        
On Error GoTo ErrorHandler
    
    invworksheet = "Invoice Report"
    ebsWorksheet = "Oracle Report"
    scWorksheet = "ScrapConnect Report"
    reconciledSheet = "Reconciled Receipts"
    reconciledInvoices = "Reconciled Invoices"
'*******************************************************  RESULTS SUMMARY ****************************************
    
    'This section calculates and prints summary data on "Home" page
    Dim lastRowMissingOracleSheet As Long
    Dim lastRowMissingSCSheet As Long
    
    
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
    
   
    
    
    a = Application.Count(Sheets(scWorksheet).Range(Sheets(scWorksheet) _
    .Cells(scStartingRow, scColumn), Sheets(scWorksheet).Cells(scSheetLR, scColumn)))
    b = Application.Count(Sheets(ebsWorksheet).Range(Sheets(ebsWorksheet) _
    .Cells(ebsStartingRow, ebsColumn), Sheets(ebsWorksheet).Cells(ebsSheetLR, ebsColumn)))
'    c = lastRowMissingSCSheet - ebsStartingRow
    c = Application.Count(Sheets("Receipts Missing From SC") _
    .Range("A1:A" & Sheets("Receipts Missing From SC") _
    .UsedRange.Rows.Count))
    d = Application.Count(Sheets("Receipts Missing From Oracle") _
    .Range("A1:A" & Sheets("Receipts Missing From Oracle").UsedRange.Rows.Count))
    e = Application.Count(Sheets("Void and Return to Vendor").Range("A1:A" & Sheets("Void and Return to Vendor") _
    .UsedRange.Rows.Count))
    f = Application.Count(Sheets("Weight Discrepancies").Range("A1:A" & Sheets("Weight Discrepancies").UsedRange _
    .Rows.Count))
    g = Application.Count(Sheets("Pending Receipts").Range("A1:A" & Sheets("Pending Receipts") _
    .UsedRange.Rows(Sheets("Pending Receipts").UsedRange.Rows.Count).Row))
    If UserForm1.OptionButton1.Value = "True" Then
    i = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A1:A" & Sheets(reconciledSheet) _
    .UsedRange.Rows.Count), ChrW(10006))
    j = Application.WorksheetFunction.CountIf(Sheets(reconciledInvoices).Range("A1:A" & Sheets(reconciledInvoices) _
    .UsedRange.Rows.Count), ChrW(10006))
    End If
    
    
'    g = Application.WorksheetFunction.CountIf(reconcileRange.Columns(Sheets(reconciledSheet) _
    .UsedRange.Find(what:="Invoice Total").Column), ">0")
'    h = Application.CountIf(Sheets(reconciledSheet).Range("B3:B" & Sheets(reconciledSheet).UsedRange.Rows.Count))
    
    If Sheets.Count < 10 Then
    h = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("A3:A" & Sheets(reconciledSheet). _
    UsedRange.Rows.Count), "Complete")
    Else
    h = Application.WorksheetFunction.CountIf(Sheets(reconciledSheet).Range("B3:B" & Sheets(reconciledSheet). _
    UsedRange.Rows.Count), "Complete")
    End If

    'create table links
    With Worksheets(1)
        .Hyperlinks.Add Anchor:=.Range("K2"), _
        Address:="", SubAddress:="'" & Worksheets(ebsWorksheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (b)
        .Range("l2").Value = "Total Oracle Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K3"), _
        Address:="", SubAddress:="'" & Worksheets(scWorksheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (a)
        .Range("l3").Value = "Total ScrapConnect Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K4"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (h)
        .Range("l4").Value = "Reconciled Receipts"
        
        If UserForm1.OptionButton1.Value = "True" Then
        .Hyperlinks.Add Anchor:=.Range("K5"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledSheet).Name & "'" & "!A1", _
        TextToDisplay:="'" & (i)
        .Range("l5").Value = "Uninvoiced Receipts"
        
        .Hyperlinks.Add Anchor:=.Range("K6"), _
        Address:="", SubAddress:="'" & Worksheets(reconciledInvoices).Name & "'" & "!A1", _
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
        .Range("l8").Value = "Receipts missing from ScrapConnect"
        
        .Hyperlinks.Add Anchor:=.Range("K9"), _
        Address:="", SubAddress:="'" & Worksheets("Receipts Missing From Oracle").Name & "'" & "!A1", _
        TextToDisplay:="'" & (d)
        .Range("l9").Value = "Receipts missing from Oracle"
        
        .Hyperlinks.Add Anchor:=.Range("K10"), _
        Address:="", SubAddress:="'" & Worksheets("Void and Return to Vendor").Name & "'" & "!A1", _
        TextToDisplay:="'" & (e)
        .Range("l10").Value = "Void and Return to Vendor receipts"
       
       .Hyperlinks.Add Anchor:=.Range("K11"), _
        Address:="", SubAddress:="'" & Worksheets("Weight Discrepancies").Name & "'" & "!A1", _
        TextToDisplay:="'" & (f)
        .Range("l11").Value = "Weight discrepancies"
    End With
     
    If UserForm1.OptionButton1.Value = "False" Then
    Sheets(1).Range("K5:L6").Delete shift:=xlUp
    End If

    For i = 2 To Sheets.Count
        Sheets(i).Activate
        
        sheetlr = Sheets(i).UsedRange.Rows _
        (Sheets(i).UsedRange.Rows.Count).Row
        
        sheetlc = Sheets(i).UsedRange.Columns.Count

        Set sheetRange = Sheets(i).Range(Sheets(i).Cells(1, 1), Sheets(i).Cells(sheetlr, sheetlc))
                            
        For j = 1 To sheetlc
            If InStr(1, Sheets(i).Cells(1, j).Value, "Date") <> 0 Then
            Sheets(i).Columns(j).NumberFormat = "mm/dd/yyyy"
            End If
        Next j
        
        With Sheets(i).UsedRange
            .Rows(1).Font.Bold = True
            If Not Sheets(i).Name = "Void and Return To Vendor" Then
                .AutoFilter
            End If
            .Columns.AutoFit
        End With
        
        If Not Sheets(i).Name = "Void and Return to Vendor" Then
        ActiveWindow.FreezePanes = False
        Sheets(i).Rows(2).Select
        ActiveWindow.FreezePanes = True
        End If
        
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
        .BorderAround ColorIndex:=0, weight:=xlThick
        .Columns.AutoFit
    End With
    
    Sheets(reconciledSheet).Visible = xlSheetHidden
    Sheets("Pending Receipts").Visible = xlSheetHidden
    Sheets("Weight Discrepancies").Visible = xlSheetHidden
    Sheets("Void and Return to Vendor").Visible = xlSheetHidden
    Sheets("Receipts Missing From Oracle").Visible = xlSheetHidden
    Sheets("Receipts Missing From SC").Visible = xlSheetHidden
    If UserForm1.OptionButton1.Value = "True" Then
    With Sheets("Reconciled Invoices")
        .Columns(Sheets("Reconciled Invoices").UsedRange _
        .Find(what:="Invoice Amount", lookat:=xlWhole).Column) _
        .NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
        .Columns(Sheets("Reconciled Invoices").UsedRange _
        .Find(what:="Invoice Dist Amount", lookat:=xlWhole).Column) _
        .NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
        .Visible = xlSheetHidden
    End With
    End If
Exit Sub
ErrorHandler: Call ErrorHandle

End Sub