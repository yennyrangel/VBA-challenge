Sub stockTicker()

'-------------------------------------------------------------------
' Declare variables
'-------------------------------------------------------------------
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim greatestPercentIncrease As Double
Dim greatestPercentDecrease As Double
Dim greatestTotalVolume As Double

' Set an initial variable for holding the total volume
Dim totalVolume As Double
totalVolume = 0

' Set an initial variable for calcular el openprice
Dim openPriceRow As Long

' Keep track of the location for each ticker in the summary table
Dim summary_table_row As Integer



' -------------------------------------------------------------------
    ' LOOP THROUGH ALL SHEETS
' -------------------------------------------------------------------

' Loop through all worksheets in the workbook

For Each ws In Worksheets

summary_table_row = 2
' =====================================================================================
' =====================================================================================
    ' First summary
' =====================================================================================
' =====================================================================================

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Set the header row for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Set the start row for the stock data
    openPriceRow = 2

    ' Loop through all rows in the worksheet
    For i = 2 To LastRow

        ' Check if the current row is the start of a new stock
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Get the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' print the results to the summary table
             ws.Range("I" & summary_table_row).Value = ticker

            ' Get the opening price
            openPrice = ws.Cells(openPriceRow, 3).Value

            ' Get the closing price
            closePrice = ws.Cells(i, 6).Value

            ' Calculate the yearly change
            yearlyChange = closePrice - openPrice
            
            ' print
            ws.Range("J" & summary_table_row).Value = yearlyChange
            
            ' Return the formatted two significant numbers YearlyChange
             ws.Range("J" & summary_table_row).NumberFormat = "0.00"
             
            ' Add conditional formatting to highlight positive and negative changes
             If ws.Range("J" & summary_table_row).Value > 0 Then
               
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            
             Else
                   
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            
             End If

            ' Calculate the percentage change
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openPrice
                 
            End If
            
            ' print
            ws.Range("K" & summary_table_row).Value = percentChange
            
            ' Return the formatted percentage percentChange
             ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            
            ' Calculate the total stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' print
            ws.Range("L" & summary_table_row).Value = totalVolume
               
             ' Increment the summary table row
             summary_table_row = summary_table_row + 1
            
             ' Reset the open price row
             openPriceRow = i + 1
            
             ' Reset the total volume
             totalVolume = 0
            
             ' If the cell immediately following a row is the totalVolume...
         Else
        
             ' Add to the total volume
             totalVolume = totalVolume + ws.Cells(i, 7).Value
           
         End If
         
    Next i
    

    
' =====================================================================================
' =====================================================================================
    ' Second summary
' =====================================================================================
' =====================================================================================
    
' Determine the Last Row
LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
' Set the header row for the othersummary table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
    
'----------------------------------------------------------------------------------
' Compare the percent change of each stock to the current greatest percent increase
'----------------------------------------------------------------------------------

greatestPercentIncrease = 0
        
For j = 2 To LastRow2

    If ws.Cells(j, 11).Value > greatestPercentIncrease Then
            greatestPercentIncrease = ws.Cells(j, 11).Value
    Else
            greatestPercentIncrease = greatestPercentIncrease
         
    End If
                    
Next j
        
' print the results to the summary table
ws.Range("Q" & 2).Value = greatestPercentIncrease
    
For j = 2 To LastRow2
    If greatestPercentIncrease = ws.Cells(j, 11).Value Then
        ws.Range("P" & 2).Value = ws.Cells(j, 9).Value
    End If
Next j
    
' Return the formatted percentage percentChange
ws.Range("Q" & 2).NumberFormat = "0.00%"
    
    
'----------------------------------------------------------------------------------
' Compare the percent change of each stock to the current greatest percent decrease
'----------------------------------------------------------------------------------
    
greatestPercentDecrease = 0
    
For j = 2 To LastRow2

    If ws.Cells(j, 11).Value < greatestPercentDecrease Then
        greatestPercentDecrease = ws.Cells(j, 11).Value
    Else
        greatestPercentDecrease = greatestPercentDecrease
         
    End If
        
Next j
        
' print the results to the summary table
ws.Range("Q" & 3).Value = greatestPercentDecrease
    
For j = 2 To LastRow2
    If greatestPercentDecrease = ws.Cells(j, 11).Value Then
            ws.Range("P" & 3).Value = ws.Cells(j, 9).Value
    End If
Next j
    
' Return the formatted percentage percentChange
ws.Range("Q" & 3).NumberFormat = "0.00%"
    

'------------------------------------------------------------------------------
' Compare the percent change of each stock to the current greatest total volume
'------------------------------------------------------------------------------

greatestTotalVolume = 0
      
For j = 2 To LastRow2

    If ws.Cells(j, 12).Value > greatestTotalVolume Then
        greatestTotalVolume = ws.Cells(j, 12).Value
    Else
        greatestTotalVolume = greatestTotalVolume
         
    End If
             
Next j
    
' print the results to the summary table
ws.Range("Q" & 4).Value = greatestTotalVolume
    
For j = 2 To LastRow2
        If greatestTotalVolume = ws.Cells(j, 12).Value Then
            ws.Range("P" & 4).Value = ws.Cells(j, 9).Value
        End If
Next j
               

'--------------------------------------------------------------------
    'Format the summary tables
'--------------------------------------------------------------------

' Autofit to display data

ws.Columns("J:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws


End Sub






