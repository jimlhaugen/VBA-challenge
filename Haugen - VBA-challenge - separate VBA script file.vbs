Sub MultipleYearStockData()

        ' Loop thru all worksheets
    For Each WS In ActiveWorkbook.Worksheets
          WS.Activate
                    
'------------------------------------------ FIRST OUTPUT ------------------------------------
          
          
            ' Determine last row in worksheet
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
            ' Set cell headers for first output as shown in Challenge instructions
            ' Autofit cell widths
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Range("I1:L1").Columns.AutoFit
        
        Dim i, j As Integer
        Dim TickerCount, TickerRowLocation, NewLastRow As Integer
        Dim OpeningPrice, ClosingPrice, YearlyChange, PercentageChange, VolCount As Double
    
            ' Initialize variables for the first interation of the following i loop
            ' Counts the number of rows for each ticker
        TickerCount = 0
            ' Rows of Ticker, Yearly Change, Percent Change & Total Volume columns
        TickerRowLocation = 1
            ' Opening price of the first ticker
        OpeningPrice = Cells(2, 3).Value
            ' Closing price at last row of each ticker
        ClosingPrice = 0
            ' Running count of Volume
        VolCount = 0
            ' Count for finding the last row for the j loop below for the number of iterations
            ' Unable to apply WS.Cells(Rows.Count, 1).End(xlUp).Row to column "I"

        NewLastRow = 1
        
            ' Loop performs each row until LastRow
        For i = 2 To LastRow
                ' Checks if ticker changes
                ' If so, populates the cells of a row of the following in the shown format:
                ' Ticker, Yearly Change, Percent Change, & Total Stock Volume cells
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ClosingPrice = Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                VolCount = VolCount + Cells(i, 7).Value
                TickerRowLocation = TickerRowLocation + 1
                Cells(TickerRowLocation, 9).Value = Cells(i, 1).Value
                Cells(TickerRowLocation, 10).NumberFormat = "#,##0"       ' format shown in instructions
                PercentageChange = YearlyChange / OpeningPrice
                Cells(TickerRowLocation, 10).Value = YearlyChange
                Cells(TickerRowLocation, 10).NumberFormat = "#,##0.00"    ' format shown in instructions
                Cells(TickerRowLocation, 12).Value = VolCount
                NewLastRow = NewLastRow + 1
                
                        ' Color code cells for positive or negagtive values
                    If YearlyChange > 0 Then
                        Cells(TickerRowLocation, 10).Interior.ColorIndex = 4    ' green
                    Else
                        Cells(TickerRowLocation, 10).Interior.ColorIndex = 3    ' red
                    End If
                    
                Cells(TickerRowLocation, 11).Value = PercentageChange
                Cells(TickerRowLocation, 11).NumberFormat = "0.00%"       ' format shown in instructions
                
                    ' Reset varaibles for each iteration of i loop
                VolCount = 0
                TickerCount = 0
                OpeningPrice = Cells(i + 1, 3).Value
                ClosingPrice = 0
                
                ' Keep running count of ticker & volume until change of ticker
            Else
                TickerCount = TickerCount + 1
                VolCount = VolCount + Cells(i, 7).Value
            End If
            
        Next i

'------------------------------------------ SECOND OUTPUT ------------------------------------
    
            ' Set cell headers for second output as shown in Challenge instructions
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Range("O2:O4").Columns.AutoFit
        Range("P1:Q1").Columns.AutoFit
        
            ' Initialize variables for the first interation of the following j loop
        GreatestPercentIncrease = Cells(2, 11)
        GreatestPercentIncreaseTicker = Cells(2, 9)
        GreatestPercentDecrease = Cells(2, 11)
        GreatestPercentDecreaseTicker = Cells(2, 9)
        GreatestTotalVolume = Cells(2, 12)
        GreatestTotalVolumeTicker = Cells(2, 9)

            
            ' Determine Greatest%Increase, Greatest%Decrease, & Greateset Volume among all Tickers
        For j = 3 To TickerRowLocation
            If Cells(j + 1, 11).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = Cells(j + 1, 11).Value
                GreatestPercentIncreaseTicker = Cells(j + 1, 9).Value
            End If
            
            If Cells(j + 1, 11).Value < GreatestPercentDecrease Then
                GreatestPercentDecrease = Cells(j + 1, 11).Value
                GreatestPercentDecreaseTicker = Cells(j + 1, 9).Value
            End If
            
            If Cells(j + 1, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = Cells(j + 1, 12).Value
                GreatestTotalVolumeTicker = Cells(j + 1, 9).Value
            End If
        
        Next j
        
            ' Display Greatest%Increase, Greatest%Decrease, & Greateset Volume as instructed
        Cells(2, 16).Value = GreatestPercentIncreaseTicker
        Cells(2, 17).Value = GreatestPercentIncrease
        Cells(2, 17).NumberFormat = "0.00%"                 ' format shown in instructions
        Cells(3, 16).Value = GreatestPercentDecreaseTicker
        Cells(3, 17).Value = GreatestPercentDecrease
        Cells(3, 17).NumberFormat = "0.00%"                 ' format shown in instructions
        Cells(4, 16).Value = GreatestTotalVolumeTicker
        Cells(4, 17).Value = GreatestTotalVolume
        Cells(4, 17).NumberFormat = "0.00E+00"              ' format shown in instructions
                
    Next WS
   
End Sub





