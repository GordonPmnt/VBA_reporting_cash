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
                Debug.Print ("Continue Overwritte")
            End If
        Else
            Call ShiftPreviousWeeksData
        End If
    End If

End Sub
Function WeekAlreadyExists(Week) As Boolean
    
    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim RepWeek As String
    
    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    Sheets(DataSheet).Activate
    RepWeek = Range("K3")
    
    If RepWeek = ("W" & Week) Then
        WeekAlreadyExists = True
    Else
        WeekAlreadyExists = False
    End If
    
    Sheets(ReportingSheet).Activate
    
End Function
Private Sub ShiftPreviousWeeksData()

    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim PreviousSocialWeeks As Range
    Dim PreviousAgingClientsWeeks As Range
    Dim PreviousAgingSuppliersWeeks As Range
    Dim PreviousStockWeeks As Range
    Dim PreviousOrderBookWeeks As Range
    
    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    Sheets(DataSheet).Activate
    
    Set PreviousSocialWeeks = _
        Range(SetParams("PreviousSocialWeeks"))
        
    Set PreviousAgingClientsWeeks = _
        Range(SetParams("PreviousAgingClientsWeeks"))
        
    Set PreviousAgingSuppliersWeeks = _
        Range(SetParams("PreviousAgingSuppliersWeeks"))
        
    Set PreviousStockWeeks = _
        Range(SetParams("PreviousStockWeeks"))
        
    Set PreviousOrderBookWeeks = _
        Range(SetParams("PreviousOrderBookWeeks"))

    Call ShiftData(PreviousSocialWeeks)
    Call ShiftData(PreviousAgingClientsWeeks)
    Call ShiftData(PreviousAgingSuppliersWeeks)
    Call ShiftData(PreviousStockWeeks)
    Call ShiftData(PreviousOrderBookWeeks)
    
    Sheets(ReportingSheet).Activate

End Sub
Private Sub ShiftData(PreviousValues)

    PreviousValues.Copy
    PreviousValues.Offset(0, -1).Select
    ActiveSheet.PasteSpecial _
        Format:=3, _
        Link:=1, _
        DisplayAsIcon:=False, _
        IconFileName:=False

End Sub