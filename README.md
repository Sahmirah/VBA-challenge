Sub YearlyStockData()
 ' Declare Variable for holding Ticker Name
    Dim Ticker As String
 ' Declare Variable for holding Year Change
    Dim Open_Year As Single
    Open_Year = 0
    Dim Close_Year As Single
    Close_Year = 0
 ' Declare Variable for holding Percent Change
    Dim Percent_Change As Double
    Percent_Change = 0
 ' Declare Variable for holding Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
 ' Declare Variable for holding Ticker Row
    Dim Ticker_Row As Double
    Ticker_Row = 2
 ' Declare Variable for holding new ticker
    Dim New_Ticker As Double
    New_Ticker = 2
    
    
    
    
 ' Assign headers to columns for all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
 ' Check for when the ticker name changes and place it in the summary table
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
        Ticker = ws.Cells(i, 1).Value
        DiffTicker = ws.Cells(i + 1, 1).Value
        If Ticker <> DiffTicker Then
            Cells(Ticker_Row, 9).Value = Ticker
            Ticker_Row = Ticker_Row + 1
        End If
    Next i
 ' Add Volume for Tickers
    For i = 2 To lastrow + 1
        Ticker = ws.Cells(i, 1).Value
        DiffTicker = ws.Cells(i + 1, 1).Value
        If Ticker = DiffTicker And i > 2 Then
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ElseIf i > 2 Then
            Cells(New_Ticker, 12).Value = Total_Stock_Volume
            New_Ticker = New_Ticker + 1
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i
  
  
  ' Find Open and close values
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Close_Year = ws.Cells(i, 6).Value
            Open_Year = ws.Cells(i, 3).Value
        End If
  ' Find  Yearly Change and Percent Change
        If Open_Year > 0 And Close_Year > 0 Then
            Change = Close_Year - Open_Year
            Percent_Change = Change / Open_Year
            ws.Cells(i, 10).Value = Change
            ws.Cells(i, 11).Value = Percent_Change
        End If
    Next i
    ' Conditional Formatting
    For i = 2 To lastrow
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
 Next ws
End Sub

![Screenshot of Results](https://user-images.githubusercontent.com/118565186/209271200-a48ed058-324a-45ba-995d-dd0461796423.PNG)

