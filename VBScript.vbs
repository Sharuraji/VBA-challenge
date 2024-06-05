Sub stockData()


' Loop through all worksheets
 For Each ws In Worksheets


' Find last row in each worksheet
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row



' Set a variable to hold tickers
Dim ticker As String

' assign total stock volume
total = 0

' Set variable to hold open price and close price
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim QuarterlyChange As Double
Dim PercentageChange As Double


'Set header name to display
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly change"
ws.Range("K1").Value = "Percentage change"
ws.Range("L1").Value = "Total stock volume"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' set a variable to keep track of different stock's open price
Dim price_row As Long
price_row = 2


' Keep track of different stock names
Dim summary_table_row As Integer
summary_table_row = 2



' Loop through ticker column to find tickers with total volume
For i = 2 To lastrow

 total = total + ws.Range("G" & i).Value
' add total volume for each ticker
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 ticker = ws.Cells(i, 1).Value

 
 ' fill appropriate columns
    ws.Range("L" & summary_table_row).Value = total
    ws.Range("I" & summary_table_row).Value = ticker
    
 
 ' Calculate quarterly change and percentage
    OpenPrice = ws.Cells(price_row, 3).Value
    
    ClosePrice = ws.Cells(i, 6).Value
    
    QuarterlyChange = ClosePrice - OpenPrice
    
    ws.Cells(summary_table_row, 10).Value = QuarterlyChange
    
    PercentageChange = QuarterlyChange / OpenPrice
    
    ws.Cells(summary_table_row, 11).Value = PercentageChange
    
    ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
    price_row = i + 1
    summary_table_row = summary_table_row + 1
     QuarterlyChange = 0
     ClosePrice = 0
     OpenPrice = 0
     total = 0
 End If
 
 
 ' conditional formatting to make color change on "quarterly change" column
 If ws.Range("J" & summary_table_row).Value > 0 Then
 ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
 
 Else
 ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
 
 End If
 
 Next i
 
 ' find greatest increase, greatest decrease and greatest total volume
    greatestIncrease = ws.Range("K2").Value
    greatestDecrease = ws.Range("K2").Value
    greatestTotal = ws.Range("L2").Value
    
        ' find last row of summary table's percentage change
        lastrow_summary_table = ws.Cells(ws.Rows.Count, 11).End(xlUp).row
       
        
        For j = 2 To lastrow_summary_table
        
        If ws.Range("K" & j + 1).Value > greatestIncrease Then
        greatestIncrease = ws.Range("K" & j + 1).Value
        greatestIncrease_ticker = ws.Range("I" & j + 1).Value
        
        ElseIf ws.Range("K" & j + 1).Value < greatestDecrease Then
        greatestDecrease = ws.Range("K" & j + 1).Value
        greatestDecrease_ticker = ws.Range("I" & j + 1).Value
        
        ElseIf ws.Range("L" & j + 1).Value > greatestTotal Then
        greatestTotal = ws.Range("L" & j + 1).Value
        greatestTotal_ticker = ws.Range("I" & j + 1).Value
        
        End If
        Next j
        
        ' Assign the values in appropriate columns
        ws.Range("P2").Value = greatestIncrease_ticker
        ws.Range("P3").Value = greatestDecrease_ticker
        ws.Range("P4").Value = greatestTotal_ticker
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q4").Value = greatestTotal
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
  
 

Next ws
 
End Sub






