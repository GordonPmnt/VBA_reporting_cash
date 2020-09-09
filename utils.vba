Option Explicit
Option Private Module
Sub ProtectSheets()

    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim Pwd As String

    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    Pwd = SetParams("Password")
    
'   Protect worksheet with a password:
    Sheets(DataSheet).Protect Password:=Pwd
    Sheets(ReportingSheet).Protect Password:=Pwd

End Sub
Sub UnProtectSheets()

    Dim DataSheet As String
    Dim ReportingSheet As String
    Dim Pwd As String

    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    Pwd = SetParams("Password")
    
'   Protect worksheet with a password:
    Sheets(DataSheet).Unprotect Password:=Pwd
    Sheets(ReportingSheet).Unprotect Password:=Pwd

End Sub
Sub CompareWeek(Week, Method)

    Dim DataSheet As String
    Dim ReportingSheet As String
    
    Dim SocialCol As Range
    Dim AGClientsCol As Range
    Dim AGSuppCol As Range
    Dim StocksCol As Range
    Dim OrdersCol As Range
    Dim MonthTurnoverCol As Range
    
    Dim RefWeek As String
    
    If Method = "UPDATE" Then
        RefWeek = Week - 1
    ElseIf Method = "RESET" Then
        RefWeek = Week
    End If

    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    Sheets(DataSheet).Activate
    
        Set SocialCol = Range("SOCIAL[W" + RefWeek + "]")
        Set AGClientsCol = Range("AG_CLIENTS[W" + RefWeek + "]")
        Set AGSuppCol = Range("AG_SUPPLIERS[W" + RefWeek + "]")
        Set StocksCol = Range("STOCKS[W" + RefWeek + "]")
        Set OrdersCol = Range("ORDERS_BOOK[W" + RefWeek + "]")
        Set MonthTurnoverCol = Range("MONTH_CA[W" + RefWeek + "]")
            
        
    Sheets(ReportingSheet).Activate
        
        Call CopyPasteData(SocialCol, Range(SetParams("CompareSocial")))
        Call CopyPasteData(AGClientsCol, Range(SetParams("CompareAGClient")))
        Call CopyPasteData(AGSuppCol, Range(SetParams("CompareAGSuppliers")))
        Call CopyPasteData(StocksCol, Range(SetParams("CompareStocks")))
        Call CopyPasteData(OrdersCol, Range(SetParams("CompareOrderBook")))
        Call CopyPasteData(MonthTurnoverCol, Range(SetParams("CompareMonthTurnover")))

End Sub

