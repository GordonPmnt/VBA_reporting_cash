Option Explicit
Option Private Module
Sub SaveData()

    Call CollectData

End Sub
Sub CollectData()

    Dim CurrentSocialColumn As Variant
    Dim CurrentAgingClients As Variant
    Dim CurrentAgingSuppliers As Variant
    Dim CurrentStocks As Variant
    Dim CurrentOrderBook As Variant
    
    CurrentSocialColumn = Range("B10:B18")
    CurrentAgingClients = Range("B85:B89")
    CurrentAgingSuppliers = Range("B95:B99")
    CurrentStocks = Range("B105:B107")
    CurrentOrderBook = Range("B119:B124")

End Sub

