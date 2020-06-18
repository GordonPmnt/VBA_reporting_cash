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
            Call CopyPasteCurrentValues(CurrentSocial, NewColumn, 2)
        
        Set NewColumn = ActiveSheet.ListObjects("AG_CLIENTS").ListColumns.Add
            NewColumn.Range(2) = "CLIENTS"
            Call CopyPasteCurrentValues(CurrentAgingClients, NewColumn, 3)

        Set NewColumn = ActiveSheet.ListObjects("AG_SUPPLIERS").ListColumns.Add
            NewColumn.Range(2) = "FOURNISSEURS"
            Call CopyPasteCurrentValues(CurrentAgingSuppliers, NewColumn, 3)
        
        Set NewColumn = ActiveSheet.ListObjects("STOCKS").ListColumns.Add
            Call CopyPasteCurrentValues(CurrentStocks, NewColumn, 2)
            
        Set NewColumn = ActiveSheet.ListObjects("ORDERS_BOOK").ListColumns.Add
            NewColumn.Range(2) = "Montant CA (K€)"
            Call CopyPasteCurrentValues(CurrentOrderBook, NewColumn, 3)
            
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

End Sub


Sub CopyPasteCurrentValues(CurrentValues, Column, StartRow)
    
    CurrentValues.Copy
    
    Column.Range(StartRow).Select
    ActiveSheet.PasteSpecial _
        Format:=3, _
        Link:=1, _
        DisplayAsIcon:=False, _
        IconFileName:=False
    
End Sub
