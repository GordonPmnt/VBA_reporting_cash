Option Explicit
Option Private Module
Sub UPDATE(Week)

    Dim DataSheet As String
    Dim ReportingSheet As String

    Dim SocialCol As String
    Dim AGClientsCol As String
    Dim AGSuppCol As String
    Dim StocksCol As String
    Dim OrdersCol As String
    
    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    SocialCol = "SOCIAL[W" + Week + "]"
    AGClientsCol = "AG_CLIENTS[W" + Week + "]"
    AGSuppCol = "AG_SUPPLIERS[W" + Week + "]"
    StocksCol = "STOCKS[W" + Week + "]"
    OrdersCol = "ORDERS_BOOK[W" + Week + "]"
    

    ' exemple
    Sheets("Data Simair").Activate
    Range(SocialCol).Select
    
    'To Do - For each table, check if week exists before pasting data
    'Here Call CopyPasteCurrentValues(CurrentValues, Column, StartRow)
    
End Sub
