Sub ExportToNew()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim NewName As String
    Dim nm As Name
    Dim ws As Worksheet
    
    If MsgBox("Results will be exported to a new workbook.  Press OK to confirm.", vbOKCancel, _
    "NewCopy") = vbCancel Then Exit Sub
    
    '   Copy User Sheets
    '   *Set sheet names to copy
    '   array("sheet name","sheet name 1","sheet name 2",...)
    On Error GoTo ErrCatcher
    
    If UserForm1.OptionButton1.Value = "True" Then
    
    Sheets(Array("Home", "Reconciled Receipts", "Pending Receipts", "Oracle Report", "ScrapConnect Report", _
    "Receipts Missing From Oracle", "Receipts Missing From SC", "Void and Return to Vendor", "Weight Discrepancies", _
    "Invoice Report", "Reconciled Invoices")).Copy
    On Error GoTo 0
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Copy
        ws.[A1].PasteSpecial xlPasteAll
'        ws.Cells.Hyperlinks.Delete
        Application.CutCopyMode = False
        Cells(2, 1).Select
        ws.Activate
    Next ws
    Cells(2, 1).Select
    
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next nm
    
    Else
    
    Sheets(Array("Home", "Reconciled Receipts", "Pending Receipts", "Oracle Report", "ScrapConnect Report", _
    "Receipts Missing From Oracle", "Receipts Missing From SC", "Void and Return to Vendor", "Weight Discrepancies")).Copy
    On Error GoTo 0
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Copy
        ws.[A1].PasteSpecial xlPasteAll
'        ws.Cells.Hyperlinks.Delete
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
ErrCatcher:
    MsgBox ("Specified sheets do not exist in this workbook.")
End Sub
