Option Explicit
Sub ResetData()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    
    Response = MsgBox( _
        "You're about to delete the data from reporting sheet. Do you want to continue?", _
        vbYesNo, _
        "Reset Data" _
    )
    
    If Response = vbYes Then
        Call API(Week, "RESET")
        MsgBox ("Reporting is now reset and ready for filling in new data.")
    End If

End Sub
