Public ebsWorksheet As String
Public scWorksheet As String
Public reconciledSheet As String
Public reconciledLR As Long
Public reconciledLC As Long
Public ebsSheetLR As Integer
Public scSheetLR As Integer
Public ebsSheetLC As Integer
Public scSheetLC As Integer
Public reconcileRange As Range

Public ebsFieldCell As Range
Public ebsColumn As Integer
Public ebsRow As Integer
Public scFieldCell As Range
Public scColumn As Integer
Public scRow As Integer
Public ebsfield As String
Public scfield As String
Public recordCellToDelete As Range
Public recordToDelete As Long
Public ebsStartingRow As Long
Public scStartingRow As Long


Private Sub ebsReportUpload_Click()
getOracleReport
End Sub

Private Sub ExportToNewWB_Click()
ExportToNew
End Sub


Private Sub hideForm_Click()
ActiveWorkbook.Close savechanges:=False
End Sub

Private Sub findDiscrepancies_Click()
getDiscrepancies
End Sub

Private Sub scReportUpload_Click()
getScrapConnectReport
End Sub


Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub ResetButton_Click()
clearEverything
End Sub

Private Sub InvoiceSheet_click()
Reconcile
End Sub