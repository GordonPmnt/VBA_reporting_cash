Option Explicit
Option Private Module
Sub SaveData()

    Call CollectCurrentWeekData

End Sub
Sub CollectCurrentWeekData()

    Dim CurrentSocial As Variant
    Dim CurrentAgingClients As Variant
    Dim CurrentAgingSuppliers As Variant
    Dim CurrentStocks As Variant
    Dim CurrentOrderBook As Variant
    
    CurrentSocial = _
        SetParams("CurrentSocial")
        
    CurrentAgingClients = _
        SetParams("CurrentAgingClients")
        
    CurrentAgingSuppliers = _
        SetParams("CurrentAgingSuppliers")
        
    CurrentStocks = _
        SetParams("CurrentStocks")
        
    CurrentOrderBook = _
        SetParams("CurrentOrderBook")

End Sub



