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
        If WeekAlreadyExists(Week) Then
            Call UnProtectSheets
            Call API(Week, "RESET")
            Call CompareWeek(Week, "RESET")
            Range("B2") = Week + 1 'Increment week
            Call ProtectSheets
            MsgBox ("Reporting is now reset and ready for filling in new data.")
        Else
            Response = MsgBox( _
                "You can't reset a week which hasn't been validated ! Please check and validate report or contact your admin if any issues.", _
                vbCritical, _
                "ERROR" _
            )
        End If
    End If

End Sub
Sub CopyPasteData(Data, Destination)
    
    Data.Copy
    Destination.Select
    Selection.PasteSpecial _
            Paste:=xlPasteValues, _
            Operation:=xlNone, _
            SkipBlanks:=False, _
            Transpose:=False
    
End Sub

