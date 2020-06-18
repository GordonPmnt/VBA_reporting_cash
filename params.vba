Option Explicit
Option Private Module
Function SetParams(Param) As String

    Select Case Param
    
        Case "DataSheet"
            SetParams = "Data Simair"

        Case "ReportingSheet"
            SetParams = "Reporting Simair"
        
        Case "CurrentSocial"
            SetParams = "B10:B18"
            
        Case "CurrentAgingClients"
            SetParams = "B86:B91"
            
        Case "CurrentAgingSuppliers"
            SetParams = "B96:B101"
        
        Case "CurrentStocks"
            SetParams = "B106:B109"
            
        Case "CurrentOrderBook"
            SetParams = "B120:B126"
            
        Case "TreasuryForecast"
            SetParams = "C34:O46"
            
    End Select

End Function

