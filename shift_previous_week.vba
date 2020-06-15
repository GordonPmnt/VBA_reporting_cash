Option Explicit
Option Private Module
Sub ShiftPreviousWeeksData()

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
Sub ShiftData(PreviousValues)

    PreviousValues.Copy
    PreviousValues.Offset(0, -1).Select
    ActiveSheet.PasteSpecial _
        Format:=3, _
        Link:=1, _
        DisplayAsIcon:=False, _
        IconFileName:=False

End Sub
