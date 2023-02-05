Attribute VB_Name = "Module1"
Sub TickerSymbol():

' We will be taking the Stock Symbol and figuring out the Ticker, Yearly Change, Percent Change over the year, and total Stock Value
' We will also be calculating Greatest % Increase, Greatest % decrease, and Greatest Volume based on the Stock Symbol

' track stock symbol changes in column a
' add up volume totals in column g

' each time the stock symbol changes in column A
' populate the name of the stock symbol in I
' track the yearly change in column j
' percent change from first open and last close - Yearly change in column K
' display the total in L
' column O - List Greatest % Increase, Greatest % Decrease and Greatest Volume
' reset the total and start tracking for the next stock symbol

  

For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    Dim Stocksymbol As String

    ' variable to hold the totals for the stock symbol
    Dim ssTotal As Double
    ssTotal = 0 ' start the initial total at 0

    ' variable holds the rows in the total columns (Columns I and L)
    Dim ssRows As Integer
    ssRows = 2 ' first row to populate in Column I and L will be row 2

    ' variable holds the yearly change
    Dim yearlychange As Double
    yearlychange = 2

    ' declare variable to hold the rows
    Dim row As Double

    ' declare variable to hold open
    Dim tickeropen As Double
    Dim tickeropenrow As Double
    tickeropenrow = 2

    ' declare variable to hold close
    Dim tickerclose As Double
    Dim stockcount As Double
    stockcount = 2

    ' declare variable to hold PercentChange
    Dim PercentChange As Double
    PercentChange = 2
    
    ' label columns and autoformat width
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").EntireColumn.AutoFit
   
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
 
    
    

    ' loop through the rows and check the changes in stock symbol
        For row = 2 To lastrow

    ' check the changes in the stock symbol
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
    
      
            ' set the stock symbol name
            Stocksymbol = ws.Cells(row, 1).Value ' grabs the value from Column A before the Change
               
            ' add to the total stock volume
            ssTotal = ssTotal + ws.Cells(row, 7).Value ' grabs the value from Column G before the change
        
            ' display the stock symbol name value on the current row of the total columns
            ws.Cells(ssRows, 9).Value = Stocksymbol
        
            ' display the stock symbol volume total on the current row of the total columns
            ws.Cells(ssRows, 12).Value = ssTotal
        
            ' grab the first open for the ticker
            tickeropen = ws.Cells(tickeropenrow, 3).Value
            tickeropenrow = row + 1

            ' grab the last close for the ticker
            tickerclose = ws.Cells(row, 6).Value
        
            ' add 1 to the stock symbol row to go to the next row
            ssRows = ssRows + 1
        
            ' reset the stock symbol total for the next stock symbol
            ssTotal = 0
        
            ' yearlychange caluculation
        
            If tickeropen = 0 Then
                yearlychange = 0
                PercentChange = 0
            Else
                yearlychange = tickerclose - tickeropen
                PercentChange = yearlychange / tickeropen
        
            End If
        
            ' display the yearly change on the current row of the total columns
            ws.Cells(ssRows - 1, 10).Value = yearlychange
                  
            ' color the yearly change red is negative green is positive
            If ws.Cells(ssRows - 1, 10).Value <= 0 Then
                ws.Cells(ssRows - 1, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(ssRows - 1, 10).Value >= 0 Then
                ws.Cells(ssRows - 1, 10).Interior.ColorIndex = 4
            End If
        
            ' display the percent change on the current row of the total columns
            ws.Cells(ssRows - 1, 11).Value = PercentChange
            
            ' format the percent change column
            ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
        
        
            Else
            'if there is no change in the stock symbol, keep adding to the total
            ssTotal = ssTotal + ws.Cells(row, 7).Value 'Grabs the value from Column G
               
            End If

    Next row
    
    'Calculate Greatest % Increase, Greatest % Decreast, Greatest Total Volume
        
        'declare variable to hold Greatest % increase and index
        Dim gi As Double
        gi = 0
        Dim giIndex As Integer
        giIndex = 0
        
        ' declare variable to hold Greatest % decrease and index
        Dim gd As Double
        gd = 0
        Dim gdIndex As Integer
        gdIndex = 0
        
        'declare variable to hold Greatest Total Volume and index
        Dim gtv As Double
        gtv = 0
        Dim gtvIndex As Integer
        gtvIndex = 0
        




        'Need to find Greatest % Increase
        gi = Application.Max(ws.Range("K2:K3000"))
        giIndex = Application.Match(gi, ws.Range("K2:K3001"), 0)
        ws.Range("P2").Value = ws.Range("I" & giIndex + 1).Value
        ws.Range("Q2").Value = gi
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ' Need to find Greatest % Decrease
        gd = Application.Min(ws.Range("K2:K3001"))
        gdIndex = Application.Match(gd, ws.Range("K2:K3001"), 0)
        ws.Range("P3").Value = ws.Range("I" & gdIndex + 1).Value
        ws.Range("Q3").Value = gd
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ' Need to find the Greatest Total Volume
        gtv = Application.Max(ws.Range("L2:L3001"))
        gtvIndex = Application.Match(gtv, ws.Range("L2:L3001"), 0)
        ws.Range("P4").Value = ws.Range("I" & gtvIndex + 1).Value
        ws.Range("Q4").Value = gtv

        ws.Range("O:Q").EntireColumn.AutoFit
    Next ws

End Sub
