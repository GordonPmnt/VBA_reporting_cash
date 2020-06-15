Option Explicit
Option Private Module
Sub SaveData()

    Dim CurrentSocial As Range
    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    Dim DataCurrentSocial As Range
    Dim DataCurrentAgingClients As Range
    Dim DataCurrentAgingSuppliers As Range
    Dim DataCurrentStock As Range
    Dim DataCurrentOrderBook As Range
    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim Week As Range
    
    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    Sheets(ReportingSheet).Activate
    Set Week = Range("K3")
    Set CurrentSocial = _
        Range(SetParams("CurrentSocial"))
    Set CurrentAgingClients = _
        Range(SetParams("CurrentAgingClients"))
    Set CurrentAgingSuppliers = _
        Range(SetParams("CurrentAgingSuppliers"))
    Set CurrentStocks = _
        Range(SetParams("CurrentStocks"))
    Set CurrentOrderBook = _
        Range(SetParams("CurrentOrderBook"))
        
    Sheets(DataSheet).Activate
    Set DataCurrentSocial = _
        Intersect( _
            Range("K:K"), _
            Range(SetParams("PreviousSocialWeeks")) _
        )
    Set DataCurrentAgingClients = _
        Intersect( _
            Range("K:K"), _
            Range(SetParams("PreviousAgingClientsWeeks")) _
        )
    Set DataCurrentAgingSuppliers = _
        Intersect( _
            Range("K:K"), _
            Range(SetParams("PreviousAgingSuppliersWeeks")) _
        )
    Set DataCurrentStock = _
        Intersect( _
            Range("K:K"), _
            Range(SetParams("PreviousStockWeeks")) _
        )
    Set DataCurrentOrderBook = _
        Intersect( _
            Range("K:K"), _
            Range(SetParams("PreviousOrderBookWeeks")) _
        )
    
    Call CopyPasteCurrentValues( _
        CurrentSocial, _
        DataCurrentSocial, _
        2, _
        Week _
    )
    Call CopyPasteCurrentValues( _
        CurrentAgingClients, _
        DataCurrentAgingClients, _
        3, _
        Week _
    )
    Call CopyPasteCurrentValues( _
        CurrentAgingSuppliers, _
        DataCurrentAgingSuppliers, _
        3, _
        Week _
    )
    Call CopyPasteCurrentValues( _
        CurrentStocks, _
        DataCurrentStock, _
        2, _
        Week _
    )
    Call CopyPasteCurrentValues( _
        CurrentOrderBook, _
        DataCurrentOrderBook, _
        3, _
        Week _
    )
    
    Sheets(ReportingSheet).Activate

End Sub
Sub CopyPasteCurrentValues(CurrentValues, DataValues, StartRow, Week)

    CurrentValues.Copy
    DataValues.Cells(StartRow, 1).Select
    ActiveSheet.PasteSpecial _
        Format:=3, _
        Link:=1, _
        DisplayAsIcon:=False, _
        IconFileName:=False
    
    DataValues.Cells(1, 1) = Week

End Sub
