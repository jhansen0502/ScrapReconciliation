VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Reconciliation User Form"
   ClientHeight    =   4068
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7092
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
