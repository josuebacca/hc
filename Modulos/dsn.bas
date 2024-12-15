Attribute VB_Name = "ModODBC"

Public Sub RegisterDataSource()
    Dim strAttribs As String
    strAttribs = "UserCommitSync = Yes" _
    & Chr$(13) & "Threads = 3" _
    & Chr$(13) & "SafeTransactions = 0" _
    & Chr$(13) & "ReadOnly = 0" _
    & Chr$(13) & "PageTimeout = 5" _
    & Chr$(13) & "MaxScanRows = 8" _
    & Chr$(13) & "MaxBufferSize = 2048" _
    & Chr$(13) & "FIL=MS Access;" _
    & Chr$(13) & "ExtendedAnsiSQL=0" _
    & Chr$(13) & "DBQ=" & App.Path & "\Fondos.mdb"
    ' Crea un nuevo DSN registrado.
    DBEngine.RegisterDatabase "FONDOS", "Microsoft Access Driver (*.mdb)", True, strAttribs
End Sub




