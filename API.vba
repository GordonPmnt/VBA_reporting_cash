Option Explicit
Option Private Module
Sub API(Week, Method)

    Dim DataSheet As String
    Dim ReportingSheet As String

    Dim CurrentSocial As Range
    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    Dim TreasuryForecast As Range
    
    Dim SocialCol As Range
    Dim AGClientsCol As Range
    Dim AGSuppCol As Range
    Dim StocksCol As Range
    Dim OrdersCol As Range
    
    Dim NewColumn As ListColumn


    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    
    Sheets(ReportingSheet).Activate
    
        Set CurrentSocial = _
            Range(SetParams("CurrentSocial"))
            
        Set CurrentAgingClients = _
            Range(SetParams("CurrentAgingClients"))
            
        Set CurrentAgingSuppliers = _
            Range(SetParams("CurrentAgingSuppliers"))
            
        Set CurrentStocks = _
            Range(SetParams("CurrentStocks"))
            
        Set CurrentOrderBook = _
            Range(SetParams("CurrentOrderBook"))
            
        Set TreasuryForecast = _
            Range(SetParams("TreasuryForecast"))
    
    
    
    If Method = "CREATE" Then
    
        Sheets(DataSheet).Activate
    
        Set NewColumn = ActiveSheet.ListObjects("SOCIAL").ListColumns.Add
            Call CopyPasteCurrentValues( _
                CurrentSocial, NewColumn, 2, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("AG_CLIENTS").ListColumns.Add
            NewColumn.Range(2) = "CLIENTS"
            Call CopyPasteCurrentValues( _
                CurrentAgingClients, NewColumn, 3, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("AG_SUPPLIERS").ListColumns.Add
            NewColumn.Range(2) = "FOURNISSEURS"
            Call CopyPasteCurrentValues( _
                CurrentAgingSuppliers, NewColumn, 3, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("STOCKS").ListColumns.Add
            Call CopyPasteCurrentValues( _
                CurrentStocks, NewColumn, 2, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("ORDERS_BOOK").ListColumns.Add
            NewColumn.Range(2) = "Montant CA (K€)"
            Call CopyPasteCurrentValues( _
                CurrentOrderBook, NewColumn, 3, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("FTE_SUM").ListColumns.Add
            NewColumn.Range(2).Offset(0, -1).Select
            Range(Selection, Selection.End(xlDown)).Copy
            Selection.Offset(0, 1).Select
            Selection.PasteSpecial _
                Paste:=xlPasteFormulas, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
            
            TreasuryForecast.Copy
            Range("C64").Select
            Selection.PasteSpecial _
                Paste:=xlPasteValues, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
        
    End If
    
    
    
    If Method = "RESET" Then
    
        Sheets(ReportingSheet).Activate
        CurrentSocial.Value = ""
        CurrentAgingClients.Value = ""
        CurrentAgingSuppliers.Value = ""
        CurrentStocks.Value = ""
        CurrentOrderBook.Value = ""
        TreasuryForecast.Value = ""
        Range("B25:B26") = ""
        Range("B113:B114") = ""
        
    End If
    
    
    
    If Method = "UPDATE" Then
    
        Sheets(DataSheet).Activate
    
        Set SocialCol = Range("SOCIAL[W" + Week + "]")
        Set AGClientsCol = Range("AG_CLIENTS[W" + Week + "]")
        Set AGSuppCol = Range("AG_SUPPLIERS[W" + Week + "]")
        Set StocksCol = Range("STOCKS[W" + Week + "]")
        Set OrdersCol = Range("ORDERS_BOOK[W" + Week + "]")
        
        Call CopyPasteCurrentValues( _
            CurrentSocial, SocialCol, 2, False _
        )
        Call CopyPasteCurrentValues( _
            CurrentAgingClients, AGClientsCol, 3, False _
        )
        Call CopyPasteCurrentValues( _
            CurrentAgingSuppliers, AGSuppCol, 3, False _
        )
        Call CopyPasteCurrentValues( _
            CurrentStocks, StocksCol, 2, False _
        )
        Call CopyPasteCurrentValues( _
            CurrentOrderBook, OrdersCol, 3, False _
        )
    
        TreasuryForecast.Copy
        Range("C64").Select
        Selection.PasteSpecial _
                Paste:=xlPasteValues, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
    
    End If

End Sub
Sub CopyPasteCurrentValues(CurrentValues, Column, StartRow, Creation)
    
    CurrentValues.Copy
    
    If Creation Then
        Column.Range(StartRow).Select
    Else
        Column.Offset(StartRow - 2, 0).Select
    End If
    
    Selection.PasteSpecial _
            Paste:=xlPasteValues, _
            Operation:=xlNone, _
            SkipBlanks:=False, _
            Transpose:=False
    
End Sub
