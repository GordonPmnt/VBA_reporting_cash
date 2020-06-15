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
            SetParams = "B85:B89"
            
        Case "CurrentAgingSuppliers"
            SetParams = "B95:B99"
        
        Case "CurrentStocks"
            SetParams = "B105:B107"
            
        Case "CurrentOrderBook"
            SetParams = "B119:B124"
        
        
        'IMPORTANT: When referencing previous week range, please always omit column B !!
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
