Option Explicit
Sub SaveData()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    Response = MsgBox("You're about to save the report data of week " + Week + ". Do you want to continue?", vbYesNo, "Save Data")
    
    If Response = vbYes Then
        Debug.Print ("Yes")
    End If

End Sub