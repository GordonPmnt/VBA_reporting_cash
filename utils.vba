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

