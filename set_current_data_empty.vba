Option Explicit
Option Private Module
Sub SetCurrentDataToEmpty()

    Dim CurrentSocial As Range
    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    
    CurrentSocial.Value = ""
    CurrentAgingClients.Value = ""
    CurrentAgingSuppliers.Value = ""
    CurrentStocks.Value = ""
    CurrentOrderBook.Value = ""
    
End Sub