Sub ClearContents():
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Columns(9).ClearContents
    ws.Columns(9).ClearFormats
    ws.Columns(10).ClearContents
    ws.Columns(10).ClearFormats
    ws.Columns(11).ClearContents
    ws.Columns(11).ClearFormats
    ws.Columns(12).ClearContents
    ws.Columns(12).ClearFormats
    ws.Columns(17).ClearContents
    ws.Columns(17).ClearFormats
    ws.Columns(18).ClearContents
    ws.Columns(18).ClearFormats
    ws.Columns(19).ClearContents
    ws.Columns(19).ClearFormats
Next ws
End Sub


Sub VBA():
Dim ws As Worksheet

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

'For each ws in worksheet
Dim columnA As String
Dim Stockopen As Double
Dim Stockclose As Double
Dim columnG As Double
Dim columnI As Integer
Dim Yearlychange As Double
Dim Percentchange As Double
Dim Totalstockvolume As Double

Dim R As Long
Dim Rowcount As Long
Dim Tickersummary As Long
Dim Counter As Integer

'set Percent change column format to percent with 2 decimal places
        ws.Columns(11).NumberFormat = "0.00%"
        ws.Range("s2").NumberFormat = "0.00%"
        ws.Range("s3").NumberFormat = "0.00%"

' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Counter = 1
Stockopen = ws.Cells(2, 3).Value
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearlychange"
ws.Cells(1, 11).Value = "Percentchange"
ws.Cells(1, 12).Value = "Totalstockvolume"

For Row = 2 To LastRow

    Totalstockvolume = Totalstockvolume + ws.Cells(Row, 7).Value
    columnA = ws.Cells(Row, 1).Value
    
    If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
            Counter = Counter + 1
              Stockclose = ws.Cells(Row, 6).Value
              Yearlychange = Stockclose - Stockopen
                If Stockopen <> 0 Then
                Percentchange = Yearlychange / Stockopen
                End If
            'Set Ticker column value
                ws.Cells(Counter, 9).Value = columnA
            'Set Yearlychange column value
                ws.Cells(Counter, 10).Value = Yearlychange
                If Yearlychange > 0 Then
                    ws.Cells(Counter, 10).Interior.ColorIndex = 4
                    Else: ws.Cells(Counter, 10).Interior.ColorIndex = 3
                End If
              'Set Totalstockvolume column value
                ws.Cells(Counter, 12).Value = Totalstockvolume
              'Set Percentchange column value
                If Stockopen <> 0 Then
            
                ws.Cells(Counter, 11).Value = Percentchange
                Else: ws.Cells(Counter, 11).Value = "NA"
                End If
              'Reset variables for next Ticker
                Totalstockvolume = 0
                Stockopen = ws.Cells(Row + 1, 3).Value
    End If
    
  Next Row
  'Autofit Ticker total columns
        ws.Columns("I:S").AutoFit

   ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws
    
    MsgBox ("Ticker Totals Complete")


End Sub




