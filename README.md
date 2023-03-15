# VBA-Challenge

Sub Ticker()

' Set variable for holding the Ticker
Dim Ticker As String

Ticker_Total = 0

' Keep track of the location for each ticker in the summary table

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

' Loop through all Tickers
For i = 2 To 759001

' Check if we are still within the same Ticker, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker
Ticker = Cells(i, 1).Value

' Add to the Ticker Total
Ticker_Total = Ticker_Total + Cells(i, 7).Value

' Print the Ticker Summary Table
Range("J" & Summary_Table_Row).Value = Ticker

' Print the Ticker Amount to the Summary Table
Range("K" & Summary_Table_Row).Value = Ticker_Total

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Ticker Total
Ticker_Total = 0

' If the cell immediately following a row is the same Ticker
Else

' Add to the Ticker Total

Ticker_Total = Ticker_Total + Cells(i, 7).Value


End If

Next i

End Sub


Sub Pricechange()

' Set variable for holding the Ticker
Dim Ticker As String

Ticker_Total = 0

' Keep track of the location for each ticker in the summary table

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

' Loop through all Tickers
For i = 2 To 759001

' Check if we are still within the same Ticker, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker
Ticker = Cells(i, 1).Value

' Add to the Ticker Total
Ticker_Total = Ticker_Total + Cells(i, 8).Value

' Print the Ticker Summary Table
Range("L" & Summary_Table_Row).Value = Ticker

' Print the Ticker Amount to the Summary Table
Range("M" & Summary_Table_Row).Value = Ticker_Total

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Ticker Total
Ticker_Total = 0

' If the cell immediately following a row is the same Ticker
Else

' Add to the Ticker Total

Ticker_Total = Ticker_Total + Cells(i, 8).Value


End If

Next i

End Sub


Sub Return_lowest_number()
'declare a variable
Dim ws As Worksheet
Set ws = Worksheets("2020")
'return lowest number in a range
ws.Range("Q19") = Application.WorksheetFunction.Min(ws.Range("m2:m3001"))
End Sub


Sub Return_highest_number()
'declare a variable
Dim ws As Worksheet
Set ws = Worksheets("2020")
'return highest number in a range
ws.Range("Q18") = Application.WorksheetFunction.Max(ws.Range("m2:m3001"))
End Sub

Sub Return_Volume_number()
'declare a variable
Dim ws As Worksheet
Set ws = Worksheets("2020")
'return highest number in a range
ws.Range("Q20") = Application.WorksheetFunction.Max(ws.Range("k2:k3001"))
End Sub
