# Module 2 Challenge

An Excel file contains a dataset related to characteristics of stocks comprised of the following seven groups of cells: ticker, date, open, high, low, close, and volume.  The data set is provided if comprised of three worksheets representing the  years 2018-2020, one for each year.

## Instructions

As instructed by the challenge, a script is to be created that will loops through all the stocks for one year and outputs the following information:

* The ticker symbol.

* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year, where positive and negative changes are formatted with green and red highlighters, respectively.

* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock.

* Additional functionality that will return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume.

* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

## Solution

The script is comprised of three loops.

### First Loop "i"

Each ticker is represented by a plurality of rows, where each row corresponds to one date.  For each ticker, the first loop i reads each row per day to produce one row per ticker which presents the following information across four columns I-L, inclusive: (1) the ticker symbol, (2) the yearly change, (3) the percentage change, and (4) the total volume of stock.  

Prior to loop i being performed, headers idenfiying these will be added to row 1.  To identify the number of iterations which loop i will perform, the number of rows of the data set is determined using the .Cells(Rows.Count, 1).End(xlUp).  In addition, variables are initialized prior to begin subjected to the loop.  For the  initial loop, the opening price is obtained.

Loop i begins with a conditional "if" statement to determine whether the ticker has reached its last row by comparing with the ticker of the next row.  If the condition has not been met, loop i jumps to its corresponding "Else" statement where variables representative of the ticker count and volume count are increased by 1 and the actual numerical volume, respectively, so that the number of rows occupied by the ticker are known and an accumulated value of the numerical volume is increased until the condition has been met.

If the ticker condition has been met, then the closing price in the row is obtained and subtracted from the opening price from which the yearly change ticker's final and the percentage change between the two can be determined for the ticker.  Given this information, ticker, yearly change, and percentage change may be presented in Columns I-K, respectively, of the top-most available row under the headers, where the yearly change is highlighted with green format for a positive change and red format if the change is negative.  In addition, the volume of the ticker's last row is added to the accumulated value and then presented in column L next to percentage change.  Afterwards, variables representative of volume count, ticker count, and closing price are reset to zero for the next internation of loop i, and the value of the opening price is set for its corresponding value.

Loop i is repeated until the data of the last row has been read and subject to the script of the loop.

It should be noted loop i included a count of the number of tickers so that the number of rows created in columns data presented in Columns I-J, inclusive, would be known for defining the number of iterations that loop j would have to perform.

### Second Loop "j"

Prior to loop j being performed, headers identifying the following are presented in the cells as instructed: (1) Greatest % Increase, (2) Greatest % Decrease, (3) Greatest Total Volume, (4) Ticker, and (5) Value.  In addition, variables are initialized using the row 2 values under Columns I-L, inclusive.

Loop j will perform the number of iterations that was determined in loop i (as discussed above) in which each row will be subjected to three "If" conditional tests to compare values of percentage changes in Column K and total stock volumes of Column L.  As the loop progresses, values of greatest percentage increase, greatest percentage decrease, and greatest total volume will be updated as the condition warrants.  After each ticker in Column I has been subjected to the three "IF" rows, the tickers and corresponding values of greatest percentage increase, greatest percentage decrease, and greatest total volume will be presented in Columns O-Q, inclusive

### Loop WS

Loop WS is included in the script to ensure the script of the first and second loops are subjected to the three worksheets corresponding to years 2018-2020, inclusive.  It is the outer-most loop within which the first and second loops are found. 
