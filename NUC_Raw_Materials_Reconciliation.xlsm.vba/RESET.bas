Sub clearEverything()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Clear Summary data on "Home" sheet
    With Sheets(1).Range("k1:l11")
        .ClearContents
        .ClearFormats
        .Columns.AutoFit
        .Rows(1).RowHeight = 14.4
    End With
    
    'Delete all sheets except "Home" sheet
    For i = Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next
    
    'Reset userform buttons and text boxes.
    With UserForm1
        .OptionButton1.Value = "False"
        .OptionButton1.Enabled = True
        .OptionButton1.ForeColor = RGB(0, 0, 0)
        .TextBox1.Value = "Oracle Receipt Report File Path"
        .TextBox1.ForeColor = RGB(0, 0, 0)
        .TextBox1.BackColor = RGB(214, 214, 214)
        .TextBox2.Value = "ScrapConnect Receipt Report File Path"
        .TextBox2.ForeColor = RGB(0, 0, 0)
        .TextBox2.BackColor = RGB(214, 214, 214)
        .TextBox3.Value = "Invoice Report File Path"
        .TextBox3.ForeColor = RGB(0, 0, 0)
        .TextBox3.BackColor = RGB(214, 214, 214)
        .scReportUpload.Enabled = False
        .scReportUpload.BackColor = RGB(214, 214, 214)
        .findDiscrepancies.Enabled = False
        .findDiscrepancies.BackColor = RGB(214, 214, 214)
        .ebsReportUpload.Enabled = True
        .ebsReportUpload.BackColor = RGB(0, 0, 255)
        .ExportToNewWB.Enabled = False
        .ExportToNewWB.BackColor = RGB(214, 214, 214)
        .invReportUpload.Enabled = False
        .invReportUpload.BackColor = RGB(214, 214, 214)
        .invoiceMatch.Enabled = False
        .invoiceMatch.BackColor = RGB(214, 214, 214)
    End With

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub