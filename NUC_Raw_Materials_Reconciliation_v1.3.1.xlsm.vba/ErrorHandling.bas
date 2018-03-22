Function ErrorHandle()
    Dim retType As Long
    Dim strMsg As String
    Dim strTitle As String
    Dim noType As Long
    Dim cancelType As Long
    Dim yesType As Long
    
    strMsg = "The program has encountered a critical error. Verify that upload reports are correct." & vbCrLf & _
    "Click Yes to send an incident report. Click No to reset form. Click Cancel to exit program."
    strTitle = "Critical Error"
    retType = MsgBox(strMsg, vbYesNoCancel + vbCritical, strTitle)
    Select Case retType
    Case 6
        yesType = MsgBox("Are you sure you want to send incident report?", vbYesNo, "Confirm")
        Select Case yesType
        Case 6
            Dim aOutlook As Object
            Dim aEmail As Object
    
            Set aOutlook = CreateObject("Outlook.Application")
            Set aEmail = aOutlook.CreateItem(0)
            With aEmail
                .importance = 2
                .Subject = "Critical Error - Reconciliation Template"
                .body = "Active Sheet:  " & ActiveSheet.Name & _
                vbCrLf & "Error:  " & Err.Number & " " & Err.Description & _
                vbCrLf & vbCrLf & _
                "Please add relevant information here."
                .attachments.Add ActiveWorkbook.FullName
                .To = "john.hansen@nucor.com"
                .display
            End With
            Set aOutlook = Nothing
            Set aEmail = Nothing
            Call clearEverything
        Case 7
            Call ErrorHandle
        End Select
    Case 7
        noType = MsgBox("Are you sure you want to reset form?", vbYesNo, "Confirm")
        Select Case noType
        Case 6
            Call clearEverything
        Case 7
            Call ErrorHandle
        End Select
    Case 2
        cancelType = MsgBox("Are you sure you want to exit?", vbYesNo, "Confirm")
        Select Case cancelType
        Case 6
'            Exit Function
            ActiveWorkbook.Close savechanges:=False
        Case 7
            Call ErrorHandle
        End Select
    End Select
    

End Function