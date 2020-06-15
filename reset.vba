Option Explicit
Sub ResetData()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    
    Response = MsgBox( _
        "You're about to delete the reporting sheet data. Do you want to continue?", _
        vbYesNo, _
        "Reset Data" _
    )
    
    If Response = vbYes Then
        Call SetCurrentDataToEmpty
    End If

End Sub