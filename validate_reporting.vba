Option Explicit
Sub ValidateReporting()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    
    Response = MsgBox( _
        "You're about to save the report data of week " + Week + ". Do you want to continue?", _
        vbYesNo, _
        "Save Data" _
    )
    
    If Response = vbYes Then
        If WeekAlreadyExists(Week) Then
            Response = MsgBox( _
                "This week has already been imported. Do you want to continue and overwrite data ?", _
                vbYesNo, _
                "Week already imported" _
            )
            If Response = vbYes Then
                Call UnProtectSheets
                Call API(Week, "UPDATE")
                Call CompareWeek(Week, "UPDATE")
                Call ProtectSheets
                MsgBox ("Reporting is now up to date.")
            End If
        Else
            Call UnProtectSheets
            Call API(Week, "CREATE")
            Call AddWeekToParams(Week)
            Call CompareWeek(Week, "UPDATE")
            Call ProtectSheets
            MsgBox ("Reporting is now up to date.")
        End If
    End If

End Sub
Function WeekAlreadyExists(Week) As Boolean
    
    Dim ReportingSheet As String
    Dim Weeks As Range
    Dim Cell As Variant
    Dim i As Integer
    
    ReportingSheet = SetParams("ReportingSheet")
    i = 0
    
    Sheets("Weeks").Activate
    Set Weeks = Range("WEEKS[REPORT]")
    
    For Each Cell In Weeks
        If ("W" + Week) = Cell Then
            i = 1
        End If
        Next Cell
    
    If i = 1 Then
        WeekAlreadyExists = True
    Else
        WeekAlreadyExists = False
    End If
    
    Sheets(ReportingSheet).Activate
End Function
Sub AddWeekToParams(Week)

    Dim NewRow As ListRow
    Dim ReportingSheet As String
    
    Sheets("Weeks").Activate
    Set NewRow = ActiveSheet.ListObjects("WEEKS").ListRows.Add
        NewRow.Range(1) = "W" + Week
    
    ReportingSheet = SetParams("ReportingSheet")
    Sheets(ReportingSheet).Activate
    
End Sub

