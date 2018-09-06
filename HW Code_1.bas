Attribute VB_Name = "Module1"
Sub name_vol_close()

Dim ws As Worksheet

For Each ws In Worksheets

' Set an initial variable for holding the brand name
Dim Stock_Name As String

' Set an initial variable for holding the total volume per stock
Dim Stock_Volume As Double
Stock_Volume = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 3

'Name columns I, J, K, L
ws.Range("I2").Value = "Ticker"
ws.Range("J2").Value = "Total Stock Volume"
ws.Range("K2").Value = "Opening Price"
ws.Range("L2").Value = "Closing Price"
ws.Range("M2").Value = "Yearly Change"
ws.Range("N2").Value = "Percent Change"


For i = 3 To 75000

' Check if we are still within the same ticker value, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the Stock name
Stock_Name = ws.Cells(i, 1).Value

' Add to the Stock Volume
Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

'Set Closing Price

Closing_Price = ws.Cells(i, 6).Value
 
' Print the Stock Name in the Summary Table
ws.Range("I" & Summary_Table_Row).Value = Stock_Name

' Print the Stock Volume to the Summary Table
ws.Range("J" & Summary_Table_Row).Value = Stock_Volume

'Print Closing Price
ws.Range("L" & Summary_Table_Row).Value = Closing_Price

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Stock Volume
Stock_Volume = 0

Else
 ' If the cell immediately following a row is the same brand...Add to the Brand Total
 Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

End If


Next i

Next ws

End Sub

Sub opening()

Dim ws As Worksheet

For Each ws In Worksheets

' Set an initial variable for holding the brand name
Dim Stock_Name As String

' Set an initial variable for holding the total volume per stock
Dim Stock_Volume As Double
Stock_Volume = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 3

'Name columns I, J, K, L
ws.Range("I2").Value = "Ticker"
ws.Range("J2").Value = "Total Stock Volume"
ws.Range("K2").Value = "Opening Price"
ws.Range("L2").Value = "Closing Price"
ws.Range("M2").Value = "Yearly Change"
ws.Range("N2").Value = "Percent Change"


For i = 3 To 75000

'If Cell above is not equal to the cell
'If Cell above is not equal to the cell
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

'Opening price is equal to the cells price value
Opening_Price = ws.Cells(i, 3).Value

'Print Opening Price
ws.Range("K" & Summary_Table_Row).Value = Opening_Price

Summary_Table_Row = Summary_Table_Row + 1

ws.Range("M" & Summary_Table_Row).Value = ws.Range("L" & Summary_Table_Row).Value - ws.Range("K" & Summary_Table_Row).Value

End If


Next i

Next ws

End Sub

Sub change()

Dim ws As Worksheet

For Each ws In Worksheets

' Set an initial variable for holding the brand name
Dim Stock_Name As String

' Set an initial variable for holding the total volume per stock
Dim Stock_Volume As Double
Stock_Volume = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 3

'Name columns I, J, K, L
ws.Range("I2").Value = "Ticker"
ws.Range("J2").Value = "Total Stock Volume"
ws.Range("K2").Value = "Opening Price"
ws.Range("L2").Value = "Closing Price"
ws.Range("M2").Value = "Yearly Change"
ws.Range("N2").Value = "Percent Change"


For i = 3 To 5000

'Yearly change (Closing minus Opening)
ws.Range("M" & Summary_Table_Row).Value = ws.Range("L" & Summary_Table_Row).Value - ws.Range("K" & Summary_Table_Row).Value

'Print positive yearly change as colorindex 4
ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4

'Print negative yearly change as colorindex 3
If ws.Range("M" & Summary_Table_Row).Value < 0 Then

ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3

End If

'Print %change
If ws.Range("K" & Summary_Table_Row).Value <> 0 Then

ws.Range("N" & Summary_Table_Row).Value = ws.Range("M" & Summary_Table_Row).Value / ws.Range("K" & Summary_Table_Row).Value

ws.Range("N" & Summary_Table_Row).NumberFormat = "0.00%"

Else

'Print NA
ws.Range("N" & Summary_Table_Row).Value = "NA"

End If

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

Next i

Next ws

End Sub

