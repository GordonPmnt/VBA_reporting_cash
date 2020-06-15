Option Explicit
Option Private Module
Sub SetCurrentDataToEmpty()

    Dim CurrentSocial As Range
    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    Dim TreasuryForecast As Range
    
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
    Set TreasuryForecast = _
        Range(SetParams("TreasuryForecast"))
    
    CurrentSocial.Value = ""
    CurrentAgingClients.Value = ""
    CurrentAgingSuppliers.Value = ""
    CurrentStocks.Value = ""
    CurrentOrderBook.Value = ""
    TreasuryForecast.Value = ""
    
    Range("B25:B26") = ""
    Range("B113:B114") = ""
    
End Sub
