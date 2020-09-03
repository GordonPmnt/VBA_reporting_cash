Option Explicit
Option Private Module
Function SetParams(Param) As String

    Select Case Param
        
        'Password
        Case "Password"
            SetParams = "HOUGOUADMIN"
            
            
        'Sheets names
        Case "DataSheet"
            SetParams = "Data IMS"

        Case "ReportingSheet"
            SetParams = "Reporting IMS"
        
        
        'Ranges of week Data (Reporting sheet)
        Case "CurrentSocial"
            SetParams = "B10:B18"
            
        Case "CurrentAgingClients"
            SetParams = "B89:B93"
            
        Case "CurrentAgingSuppliers"
            SetParams = "B99:B103"
        
        Case "CurrentStocks"
            SetParams = "B109:B112"
            
        Case "CurrentOrderBook"
            SetParams = "B123:B129"
            
        Case "TreasuryForecast"
            SetParams = "C37:O49"
            
        Case "CurrentMonthTurnover"
            SetParams = "B117:B118"
            
            
        'Ranges where previous week data are paste for comparison (Reporting sheet)
        Case "CompareSocial"
            SetParams = "G10"
            
        Case "CompareAGClient"
            SetParams = "I88"
            
        Case "CompareAGSuppliers"
            SetParams = "I98"
            
        Case "CompareStocks"
            SetParams = "I108"
            
        Case "CompareOrderBook"
            SetParams = "I122"
        
            
    End Select

End Function
