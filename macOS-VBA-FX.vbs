Function getHTTP()
Dim ConnString As String
ConnString = "URL;http://api.fixer.io/latest?base=USD"
With ActiveSheet.QueryTables.Add(Connection:=ConnString, Destination:=Range("A1"))
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
End Function

Sub ParseJSON()
Dim aRun, JSONinput As String
Dim RawData, Info, Rates As Variant
Sheets("FX").UsedRange.ClearContents
aRun = getHTTP()
JSONinput = Range("A1").Value
'clean data
RawData = Split(JSONinput, ":{")
Info = Split(Replace(RawData(0), "{", ""), ",")
Rates = Split(Replace(RawData(1), "}}", ""), ",")
'loop basic info data
On Error Resume Next
For a = 0 To UBound(Info)
setArray = Split(Info(a), ":")
Cells(a + 2, 1).Value = Replace(setArray(0), Chr(34), "")
Cells(a + 2, 2).Value = Replace(setArray(1), Chr(34), "")
Next a
'loop rates
For a = 0 To UBound(Rates)
setArray = Split(Rates(a), ":")
Cells(a + 5, 1).Value = Replace(setArray(0), Chr(34), "")
Cells(a + 5, 2).Value = Replace(setArray(1), Chr(34), "")
Next a
End Sub