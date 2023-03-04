Attribute VB_Name = "Module2"
Sub Loop_Ticker_YearlyChange_PercentChange_TotalVolume()
    
'Loop through all worksheets
For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
WorksheetName = ws.Name
  
    'Define variables
    output_row = 2
    yearlychange = 0
    percentchange = 0
    totalvolume = 0
    open_price = ws.Cells(2, 3).Value
    
'set last row variable
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

For input_row = 2 To LastRow

    'loop through until ticker changes value
    If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1).Value Then
    
    ws.Cells(output_row, 9).Value = ws.Cells(input_row, 1).Value
    
    'Calculate change from open to close of year and add to column
    yearlychange = (ws.Cells(input_row, 6).Value - open_price)
    ws.Cells(output_row, 10).Value = yearlychange
    
    'Calculate yearly percentage change from open to close and add to column
    percentchange = yearlychange / open_price
    ws.Cells(output_row, 11).Value = percentchange
    ws.Cells(output_row, 11).NumberFormat = "0.00%"
    
    
    'Add Total stock volume for the year to column
    totalvolume = totalvolume + ws.Cells(input_row, 7).Value
    ws.Cells(output_row, 12).Value = totalvolume
    
    output_row = output_row + 1
    yearlychange = 0
    percentchange = 0
    totalvolume = 0
    open_price = ws.Cells(input_row + 1, 3).Value
    
    Else
        totalvolume = totalvolume + ws.Cells(input_row, 7).Value

        End If

Next input_row

'Add conditional for yearlychange column to add green if positive, and red if negative
For i = 2 To output_row

    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
   
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 4
 
        End If
        
      Next i
    
    Next ws
    
End Sub
Sub Loop_Greatest_Increase_Decrease_TotalVolume()

'Loop through all worksheets
For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
WorksheetName = ws.Name

'set variables for greatest increase, greatest decrease, total volume and their tickers
    increase_ticker = ""
    greatest_increase = -1E+20
    decrease_ticker = ""
    greatest_decrease = 1E+20
    volume_ticker = ""
    greatest_volume = -1E+24

'set last row variable
LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row

'start For loop
For input_row = 2 To LastRow
    
    'set conditional to loop through to find greatest percent increase
    If ws.Cells(input_row, 11).Value > greatest_increase Then
    greatest_increase = ws.Cells(input_row, 11).Value
    increase_ticker = ws.Cells(input_row, 9).Value
    
        End If
    
    'set conditional to loop through and find greatest percent decrease
    If ws.Cells(input_row, 11).Value < greatest_decrease Then
    greatest_decrease = ws.Cells(input_row, 11).Value
    decrease_ticker = ws.Cells(input_row, 9).Value
    
        End If
        
    'set conditional to loop through and find the greatest Total Stock Volume
    If ws.Cells(input_row, 12).Value > greatest_volume Then
    greatest_volume = ws.Cells(input_row, 12).Value
    volume_ticker = ws.Cells(input_row, 9).Value
        
        End If
            

  Next input_row

    'set output rows for greatest increase
    ws.Cells(2, 16).Value = increase_ticker
    ws.Cells(2, 17).Value = greatest_increase
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'set output rows for greatest decrease
    ws.Cells(3, 16).Value = decrease_ticker
    ws.Cells(3, 17).Value = greatest_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    
    'set output rows for greatest total stock volume
    ws.Cells(4, 16).Value = volume_ticker
    ws.Cells(4, 17).Value = greatest_volume

Next ws

End Sub
Sub Start_Call_Modules()

'call the order of the modules to run
Call Loop_Ticker_YearlyChange_PercentChange_TotalVolume
Call Loop_Greatest_Increase_Decrease_TotalVolume


End Sub
