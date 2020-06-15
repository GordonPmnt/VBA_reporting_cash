Option Explicit
Option Private Module
Function SetParams(Param) As Variant

    Select Case Param
    
        Case "DataSheet"
            SetParams = "Data Simair"
            
        Case "ReportingSheet"
            SetParams = "Reporting Simair"
        
        Case "CurrentSocial"
            SetParams = Range("B10:B18")
            
        Case "CurrentAgingClients"
            SetParams = Range("B85:B89")
            
        Case "CurrentAgingSuppliers"
            SetParams = Range("B95:B99")
        
        Case "CurrentStocks"
            SetParams = Range("B105:B107")
            
        Case "CurrentOrderBook"
            SetParams = Range("B119:B124")
            
    End Select

End Function