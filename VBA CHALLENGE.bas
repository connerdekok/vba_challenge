Attribute VB_Name = "Module5"
Sub StockFindings112()
'Define Variables
    Dim ws As Worksheet
    Dim i As Long
    Dim SummaryRowTable As Long
    SummaryRowTable = 2
    Dim Counter As Integer
    Dim Ticker As String
    Dim QuarterOpenPrice As Double
    Dim QuarterClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
      'Loop sheets in Workbook
    
    For Each ws In Worksheets
    
'Define Greatest Or Least Variables
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestVolume As Double
    GreatestVolume = 0
    Dim TickerGreatestIncrease As String
    Dim TickerGreatestDecrease As String
    Dim TickerGreatestVolume As String
 
    
    'Establish Last Row for each sheet
    
    SummaryRowTable = 2
    
    Dim RowsCount As Long
    RowsCount = ws.Range("A1").End(xlDown).Row
    
    'Create New Columns
    
    ws.Range("I1").EntireColumn.Insert
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Range("J1").EntireColumn.Insert
    ws.Cells(1, 10).Value = "Quarterly Change"
    
    ws.Range("K1").EntireColumn.Insert
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Range("L1").EntireColumn.Insert
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    

    ' Establish Last Row for each sheet
    SummaryRowTable = 2
    RowsCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create New Columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ' Summarized Columns
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 17) = "Ticker"
    ws.Cells(1, 18) = "Value"
            

    ' Start Loop Through Stocks
    For i = 2 To RowsCount

        ' Check Stock Data And Tickers
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryRowTable).Value = Ticker

            QuarterClosingPrice = ws.Cells(i, 6).Value

            QuarterlyChange = QuarterClosingPrice - QuarterOpenPrice
            ws.Range("J" & SummaryRowTable).Value = QuarterlyChange
            

            PercentChange = (QuarterlyChange / QuarterOpenPrice) * 1
            ws.Range("K" & SummaryRowTable).Value = PercentChange
            
            'Format to percentages in K
            ws.Range("K" & SummaryRowTable).NumberFormat = "0.00%"

            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Range("L" & SummaryRowTable).Value = TotalStockVolume

            ' Advance Summary Table To next Stock
            SummaryRowTable = SummaryRowTable + 1

            ' Reset Stock Volume for each new Ticker
            TotalStockVolume = 0

        Else

            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            ' Set opening price for new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                QuarterOpenPrice = ws.Cells(i, 3).Value
            End If
    
        End If
        
    Next i
    
    ' Add Color to quarterly change collumns
    For i = 2 To RowsCount
        If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            
        End If
    Next i
        
     'New Loop to find greatest values
    
        'Check rows percentages
    For i = 2 To SummaryRowTable
    'Find greatest increase in spreadsheet
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            TickerGreatestIncrease = ws.Cells(i, 9).Value
        End If
         
         ' Find greatest decrease in spreadsheet
        If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            TickerGreatestDecrease = ws.Cells(i, 9).Value
        End If
   
    ' Check for values that are above the current highest seleceted value
        If ws.Cells(i, 12).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(i, 12).Value
            TickerGreatestVolume = ws.Cells(i, 9).Value
        End If
        
    Next i
    
' Output the greatest increase, decrease, volume, and format
ws.Cells(2, 17).Value = TickerGreatestIncrease
ws.Cells(2, 18).Value = GreatestIncrease
ws.Cells(2, 18).NumberFormat = "0.00%"

ws.Cells(3, 17).Value = TickerGreatestDecrease
ws.Cells(3, 18).Value = GreatestDecrease
ws.Cells(3, 18).NumberFormat = "0.00%"

ws.Cells(4, 17).Value = TickerGreatestVolume
ws.Cells(4, 18).Value = GreatestVolume


    
    'Autofit the columns
    ws.Columns("J:L").AutoFit
    ws.Columns("O").AutoFit
    

            
    
Next ws
End Sub

