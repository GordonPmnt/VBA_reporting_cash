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
    Dim CurrentMonthTurnover As Range
    
    Dim SocialCol As Range
    Dim AGClientsCol As Range
    Dim AGSuppCol As Range
    Dim StocksCol As Range
    Dim OrdersCol As Range
    Dim MonthTurnoverCol As Range
    
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
            
        Set CurrentMonthTurnover = _
            Range(SetParams("CurrentMonthTurnover"))
            
    
    
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
            NewColumn.Range(2) = "Montant CA (KÛ)"
            Call CopyPasteCurrentValues( _
                CurrentOrderBook, NewColumn, 3, True _
            )
        Set NewColumn = ActiveSheet.ListObjects("MONTH_CA").ListColumns.Add
            Call CopyPasteCurrentValues( _
                CurrentMonthTurnover, NewColumn, 2, True _
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
        CurrentMonthTurnover.Value = ""
        
  
        Union( _
            Intersect(Rows("36:41"), TreasuryForecast), _
            Intersect(Rows("43:45"), TreasuryForecast) _
        ).Select
        
        Selection = ""
        
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
        Set MonthTurnoverCol = Range("MONTH_CA[W" + Week + "]")
        
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
        Call CopyPasteCurrentValues( _
            CurrentMonthTurnover, MonthTurnoverCol, 2, False _
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
Sub ComputeAllTrends(StartWeek, Size, Method)

    Dim CurrentAgingClients As Range
    Dim CurrentAgingSuppliers As Range
    Dim CurrentStocks As Range
    Dim CurrentOrderBook As Range
    
    Dim RefWeek As String
    
    If Method = "UPDATE" Then
        RefWeek = StartWeek - 1
    ElseIf Method = "RESET" Then
        RefWeek = StartWeek
    End If
        
        
    Set CurrentAgingClients = Range(SetParams("CurrentAgingClients"))
    Call ComputeTrend( _
        RefWeek, _
        Size, _
        CurrentAgingClients, _
        "AG_CLIENTS", _
        RepOffset:=5 _
    )

    Set CurrentAgingSuppliers = Range(SetParams("CurrentAgingSuppliers"))
    Call ComputeTrend( _
        RefWeek, _
        Size, _
        CurrentAgingSuppliers, _
        "AG_SUPPLIERS", _
        RepOffset:=5 _
    )
            
    Set CurrentStocks = Range(SetParams("CurrentStocks"))
    Call ComputeTrend( _
        RefWeek, _
        Size, _
        CurrentStocks, _
        "STOCKS", _
        RepOffset:=5 _
    )
    
    Set CurrentOrderBook = Range(SetParams("CurrentOrderBook"))
    Call ComputeTrend( _
        RefWeek, _
        Size, _
        CurrentOrderBook, _
        "ORDERS_BOOK", _
        RepOffset:=5 _
    )

End Sub
Sub ComputeTrend(RefWeek, Size, DataRange, DataTableName, RepOffset)

    Dim StartCol As Range
    Dim ActiveCellRep As Range
    Dim ActiveCellData As Range
    Dim ActiveRowData As Range
    Dim DestCellRep As Range
    
    Dim i As Integer
    Dim j As Integer
    
    i = 1
    For Each ActiveCellRep In DataRange
        Set StartCol = Range(DataTableName + "[W" + RefWeek + "]")
        j = 1
        For Each ActiveCellData In StartCol
            If i = j Then
                Set ActiveRowData = Range(ActiveCellData, ActiveCellData.Offset(0, -Size))
                Set DestCellRep = ActiveCellRep.Offset(0, RepOffset)
                DestCellRep.Formula = "=IFERROR(LINEST(" + ActiveRowData.Address(External:=True) + "),"""" )"
            End If
            j = j + 1
        Next ActiveCellData
        i = i + 1
    Next ActiveCellRep

End Sub
