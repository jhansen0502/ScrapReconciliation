Sub ExportToNew()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim NewName As String
    Dim nm As Name
    Dim ws As Worksheet
    
    'confirmation modal to export
    If MsgBox("Results will be exported to a new workbook.  Press OK to confirm.", vbOKCancel, _
    "NewCopy") = vbCancel Then Exit Sub
    
    'checks for invoice matching to copy correct tables to new file
    If UserForm1.OptionButton1.Value = "True" Then
    
    Sheets(Array("Home", "Reconciled Receipts", "Pending Receipts", "Oracle Report", "ScrapConnect Report", _
    "Receipts Missing From Oracle", "Receipts Missing From SC", "Void and Return to Vendor", "Weight Discrepancies", _
    "Invoice Report", "Reconciled Invoices")).Copy
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Copy
        ws.[A1].PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        Cells(2, 1).Select
        ws.Activate
    Next ws
    Cells(2, 1).Select
    
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next nm
    
    Else
    
    'if no invoice matching, these are the copied tables
    Sheets(Array("Home", "Reconciled Receipts", "Pending Receipts", "Oracle Report", "ScrapConnect Report", _
    "Receipts Missing From Oracle", "Receipts Missing From SC", "Void and Return to Vendor", "Weight Discrepancies")).Copy
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Copy
        ws.[A1].PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        Cells(2, 1).Select
        ws.Activate
    Next ws
    Cells(2, 1).Select
    
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next nm
    
    End If
    
    UserForm1.Hide
    
    'input box for new file name
    NewName = InputBox("Please enter the name of your new workbook" & vbCr & _
    "File will be saved in current folder.", "New Copy")
    
    If NewName = "" Then
        UserForm1.Show
        ActiveWorkbook.Close savechanges:=False
        Exit Sub
        
    End If
    
    With ActiveWorkbook
        Sheets(1).Activate
        Sheets(1).Shapes.SelectAll
        Selection.Delete
        Sheets(1).Columns("A:J").Delete
        
        For i = 2 To Sheets.Count
            Sheets(i).Visible = xlSheetVisible
        Next
        
    End With
    
    ActiveWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & NewName & ".xlsx"
    ActiveWorkbook.Close savechanges:=False
    
    Workbooks.Open ThisWorkbook.Path & "\" & NewName & ".xlsx"
    
    ThisWorkbook.Close savechanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
    With UserForm1
        .ExportToNewWB.Enabled = False
        .ExportToNewWB.BackColor = RGB(214, 214, 214)
    End With
    
Exit Sub
ErrorHandler:     Call ErrorHandle

End Sub
