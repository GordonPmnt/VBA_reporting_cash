Option Explicit
Option Private Module
Sub SetCurrentDataToEmpty()

    Dim CurrentSocial As Range
    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    
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
    
    CurrentSocial.Value = ""
    CurrentAgingClients.Value = ""
    CurrentAgingSuppliers.Value = ""
    CurrentStocks.Value = ""
    CurrentOrderBook.Value = ""
    
End Sub
