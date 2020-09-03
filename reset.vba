Option Explicit
Sub ResetData()

    Dim Response As Integer
    Dim Week As String
    
    Week = Range("B2")
    
    Response = MsgBox( _
        "You're about to delete the data from reporting sheet. Do you want to continue?", _
        vbYesNo, _
        "Reset Data" _
    )
    
    If Response = vbYes Then
        If WeekAlreadyExists(Week) Then
            Call UnProtectSheets
            Call API(Week, "RESET")
            Call ShiftWeek(Week)
            Range("B2") = Week + 1 'Increment week
            Call ProtectSheets
            MsgBox ("Reporting is now reset and ready for filling in new data.")
        Else
            Response = MsgBox( _
                "You can't reset a week which hasn't been validated ! Please check and validate report or contact your admin if any issues.", _
                vbCritical, _
                "ERROR" _
            )
        End If
    End If

End Sub
Sub ShiftWeek(Week)

    Dim DataSheet As String
    Dim ReportingSheet As String
    
    Dim SocialCol As Range
    Dim AGClientsCol As Range
    Dim AGSuppCol As Range
    Dim StocksCol As Range
    Dim OrdersCol As Range

    DataSheet = SetParams("DataSheet")
    ReportingSheet = SetParams("ReportingSheet")
    
    Sheets(DataSheet).Activate
    
        Set SocialCol = Range("SOCIAL[W" + Week + "]")
        Set AGClientsCol = Range("AG_CLIENTS[W" + Week + "]")
        Set AGSuppCol = Range("AG_SUPPLIERS[W" + Week + "]")
        Set StocksCol = Range("STOCKS[W" + Week + "]")
        Set OrdersCol = Range("ORDERS_BOOK[W" + Week + "]")
            
        
    Sheets(ReportingSheet).Activate
        
        Call CopyPasteData(SocialCol, Range(SetParams("CompareSocial")))
        Call CopyPasteData(AGClientsCol, Range(SetParams("CompareAGClient")))
        Call CopyPasteData(AGSuppCol, Range(SetParams("CompareAGSuppliers")))
        Call CopyPasteData(StocksCol, Range(SetParams("CompareStocks")))
        Call CopyPasteData(OrdersCol, Range(SetParams("CompareOrderBook")))
        

End Sub
Sub CopyPasteData(Data, Destination)
    
    Data.Copy
    Destination.Select
    Selection.PasteSpecial _
            Paste:=xlPasteValues, _
            Operation:=xlNone, _
            SkipBlanks:=False, _
            Transpose:=False
    
End Sub

