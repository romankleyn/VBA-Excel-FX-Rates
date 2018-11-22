Sub fxRate()
    'https://ca.investing.com/currencies/live-currency-cross-rates
    'Indices, Commodities, Forex, Crypto
    Dim ConnString As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks("fxrate.xlsm")
    Set ws = wb.Sheets("Data")
    ws.UsedRange.ClearContents
    ConnString = "URL;https://ca.investing.com/currencies/live-currency-cross-rates"
    http ConnString, "A1", ws
    
    'Clean
    ws.Range("A38:G80").Clear
    ws.Range("A1:H1").Delete
    ws.Range("A1:A40").Delete
    ws.Range("E1:F40").Delete
    
    'Titles
    ws.Range("A1") = "Indices"
    ws.Range("B1") = "Current Price($)"
    ws.Range("C1") = "Change($)"
    ws.Range("D1") = "Change(%)"
    ws.Range("A10") = "Commodities"
    ws.Range("A19") = "Forex"
    ws.Range("A28") = "Crypto"
    
    'Set Format
    ws.Columns("B").NumberFormat = "$#,##0.000"
    ws.Columns("C").NumberFormat = "$#,##0.000"
    
    'Remove Broken Hyperlinks
    RemoveHyperlink ws
    ws.Range("A1").Select
End Sub
Private Sub RemoveHyperlink(ws As Worksheet)
    Dim n As Integer
    For n = 1 To 50
        On Error Resume Next
        ws.Range("A" & n).Select
        Selection.Hyperlinks(1).Delete
    Next n
End Sub
Private Sub http(ConnString As String, rng As String, ws As Worksheet):
With ws.QueryTables.Add(Connection:=ConnString, Destination:=ws.Range(rng))
        .Name = _
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
        End With
End Sub