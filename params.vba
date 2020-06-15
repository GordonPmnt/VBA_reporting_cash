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
            SetParams = "B86:B90"
            
        Case "CurrentAgingSuppliers"
            SetParams = "B96:B100"
        
        Case "CurrentStocks"
            SetParams = "B106:B108"
            
        Case "CurrentOrderBook"
            SetParams = "B120:B125"
            
        Case "TreasuryForecast"
            SetParams = "C36:O45"
        

        'IMPORTANT: When referencing previous week range, please always omit column B !!
        ' Column B is the oldest week, which is deleted when shifting data
        '------------------------------------------------------------------------
        
        Case "PreviousSocialWeeks"
            SetParams = "C3:K12"
            
        Case "PreviousAgingClientsWeeks"
            SetParams = "C21:K27"
        
        Case "PreviousAgingSuppliersWeeks"
            SetParams = "C30:K36"
        
        Case "PreviousStockWeeks"
            SetParams = "C40:K43"
        
        Case "PreviousOrderBookWeeks"
            SetParams = "C49:K56"
            
    End Select

End Function
