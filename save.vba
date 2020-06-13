Option Explicit
Sub SaveData()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    Response = MsgBox("You're about to save the report data of week " + Week + ". Do you want to continue?", vbYesNo, "Save Data")
    
    If Response = vbYes Then
        If WeekAlreadyExists(Week) Then
            Response = MsgBox("This week has already been imported. Do you want to continue and overwrite data ?", vbYesNo, "Week already imported")
            If Response = vbYes Then
                Debug.Print ("Continue Overwritte")
            End If
        Else
            Debug.Print ("Continue New")
        End If
    End If

End Sub
Function WeekAlreadyExists(Week) As Boolean
    
    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim RepWeek As String
    
    DataSheet = SetDataSheet()
    ReportingSheet = SetReportingSheet()
    
    Sheets(DataSheet).Activate
    RepWeek = Range("K3")
    
    If RepWeek = ("W" & Week) Then
        WeekAlreadyExists = True
    Else
        WeekAlreadyExists = False
    End If
    
    Sheets(ReportingSheet).Activate
    
End Function
