Sub Master_Ticker_M0d2()
            'set varibles
Dim ws As Worksheet
            'start for loop
For Each ws In Worksheets
            'Create the column headings for each sheet in workbook
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
            'Fill in data in each worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer
Dim TickerRow As Long
Dim Ticker As String

Dim openPrice As Double
Dim closePrice As Double
Dim priceChange As Double
Dim price_chg_pct As Double
Dim vol As Double
Dim opencounter As Integer

Ticker = ""
openPrice = 0
closePrice = 0
priceChange = 0
price_chg_pct = 0
opencounter = 2


            'Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TickerRow = 1

            'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow
openPrice = ws.Cells(opencounter, 3).Value

            'Ticker symbol to correct "I" col in each sheet
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerRow = TickerRow + 1
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(TickerRow, "I").Value = Ticker
        
    
     closePrice = ws.Cells(i, 6).Value
     vol = ws.Cells(i, 7).Value
     priceChange = (closePrice - openPrice)
     
     price_chg_pct = (price_chg_pct / openPrice) * 100

     price_chg_pct = (closePrice - openPrice) / openPrice
     
    priceChange = ws.Cells(TickerRow, 10).Value
    price_chg_pct = ws.Cells(TickerRow, 11).Value
    vol = ws.Cells(TickerRow, 12).Value
            
    TickerRow = TickerRow + 1
    
    ElseIf open_price = 0 Then
     
        openPrice = Null
            
            
            vol = 0
            opencounter = i + 1
    End If
    
    
    
     'Calculate change in Price
     
        

    

            'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.



            'Set new variables for prices and percent changes


Next i
    
Next ws

End Sub

'This is the revised copy of the code



Sub Master2_ticker()
            'This is the revised copy of the code****

            'set varibles
Dim ws As Worksheet
            'start for loop
For Each ws In Worksheets
            'Create the column headings for each sheet in workbook
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
            'Fill in data in each worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer
Dim TickerRow As Long
Dim Ticker As String

Dim openPrice As Double
Dim closePrice As Double
Dim priceChange As Double
Dim price_chg_pct As Double
Dim vol As Double
Dim opencounter As Integer

Ticker = ""
openPrice = 0
closePrice = 0
priceChange = 0
price_chg_pct = 0
opencounter = 2


'Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'TickerRow keeps track of the row for each ticker in summary table
TickerRow = 2

            'Do loop of each worksheet to Lastrow
For i = 2 To Lastrow
            'Add to the volume total
    vol = vol + ws.Cells(i, 7).Value
            'Set the ticker name
    'Ticker = ws.Cells(i, 1).Value
            'Set the opening price
    'openPrice = ws.Cells(opencounter, 3).Value
            'Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set the ticker name
    Ticker = ws.Cells(i, 1).Value
            'Set the opening price
    openPrice = ws.Cells(opencounter, 3).Value
                'Set the closing price
        closePrice = ws.Cells(i, 6).Value
                'Print the ticker name in the Summary Table
                'ws.Cells(TickerRow, 9).Value = Ticker
                'Calculate the price change
        priceChange = closePrice - openPrice
                'Calculate the percent change
        price_chg_pct = priceChange / openPrice
                'Print the results in the Summary Table
        ws.Range("I" & TickerRow).Value = Ticker
        ws.Range("J" & TickerRow).Value = priceChange
        ws.Range("K" & TickerRow).Value = price_chg_pct
        ws.Range("J" & TickerRow).NumberFormat = "0.00"
        ws.Range("K" & TickerRow).NumberFormat = "0.00%"
        ws.Range("L" & TickerRow).Value = vol
      
      'Change color index of cell for pos and neg vales withing "J"
      
        If ws.Range("J" & TickerRow).Value > 0 Then
            ws.Range("J" & TickerRow).Interior.ColorIndex = 4
        Else
            ws.Range("J" & TickerRow).Interior.ColorIndex = 3
     
        End If
        
        
                'Update TickerRow so that each new ticker is printed on a new row
        TickerRow = TickerRow + 1
                'Reset the volume total
        vol = 0
                'Set the opening price for the next ticker based on current row + 1
        opencounter = i + 1
    End If
            
            
    
    
            'Greatest % Increase
     'ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & Lastrow))
     'ws.Range("Q2").NumberFormat = "0.00%"
     Range("Q2").Value = "%" & WorksheetFunction.Max(Range("K2:K" & Lastrow)) * 100

            'Greatest % Decrease
     'ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & Lastrow))
     'ws.Range("Q3").NumberFormat = "0.00%"
     Range("Q3").Value = "%" & WorksheetFunction.Min(Range("K2:K" & Lastrow)) * 100

            'Greatest Total Volume
            
     ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))
     

    maxincreaseindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    maxdecreaseindex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    maxvolumeindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Lastrow)), ws.Range("L2:L" & Lastrow), 0)
 
 

     ws.Range("P2").Value = ws.Cells(maxincreaseindex + 1, 9)
     ws.Range("P3").Value = ws.Cells(maxdecreaseindex + 1, 9)
     ws.Range("P4").Value = ws.Cells(maxvolumeindex + 1, 9)


            


Next i
    
Next ws

End Sub
