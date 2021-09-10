# Stock Analysis
Performing analysis on stocks using VBA in Excel.
## Overview of Project
In this assignment we were given two data sets for stocks for 2017 and 2018. These data sets included dates, opening and closing prices, highest and lowest prices and volume amount for a give day for each respective stock. The purpose of the analysis was to see how the individual stocks performed in each year and whether they increased or dropped in value from the starting closing price to the ending closing price. Using VBA formatting, a visual result was given to see the percentage increase or decrease in valuation for these stocks. The total daily volume was also calculated to see how many shares of the respective stock was shared in a year.
# Results
## Analysis and Coding
A VBA macro was written to loop through the numerous rows of data for a year and note the starting price, ending price and total volume for a specific stock. For this assignment, the original code was refactored to reduce time taken for it to perform analysis. An input option was added to the script to allow the user to choose the year they wish to perform the analysis on; 2017 or 2018. 
### Coding
Initialize array of all tickers:
    
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

The names of the stock tickers were intialised as an array which could then be indexed by the code to read and store information on respective prices and volume for each stock. The respective values were then added to the initialised arrays of:
* tickerVolumes
* tickerStartingPrices
* tickerEndingPrices

The following code is how the values were added to the arrays:

    For tickerIndex = 0 To 11
    ticker = tickers(tickerIndex)
    tickerVolumes(tickerIndex) = 0
 
        For i = 2 To RowCount
            If Cells(i, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
                'Find the starting price for the current ticker.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
             End If
        Next i
        
    Next tickerIndex

The first part of this code explains the module to complete the loop for the 12 different stock indexes as shown above. The inner loop iterating over all the rows of the data set allows the code to read every row and see if the stock index is the required one, then:
1) Increase the total Volume of that stock ticker
2) If the current row is the first one with the respective stock ticker then set starting price as the closing price of that row, else continue forward with the code.
3) If the current row is the last one with the respective stock ticker then set ending price as the closing price of that row, else continue forward with the code.

After the inner loop is completed, move on to the next stock ticker. Some things to note, RowCount is calculated using the function: **RowCount = Cells(Rows.Count, "A").End(xlUp).Row**. Another important part of this code is the formatting of the cells to relate to the values inside. Using the starting and ending prices, a percentage increase or decrease is calculated which is also formatted to have a cell colour of green if its a positive percentage or increase in price and red if its a negative percentage or decrease in price. These values are outputted according to the index from all four arrays.

## Results
The following two images are screenshots of the output of the code and show the performance of the stocks for the years 2017 and 2018.

![Refactored Code 2017](https://user-images.githubusercontent.com/87828174/132926691-612bd55e-2459-4bc3-8b69-9ff96b332616.png)
![Refactored Code 2018](https://user-images.githubusercontent.com/87828174/132926704-fcf23781-b4f2-484d-b4c5-85242b49a471.png)

### Analysis of the results



time elapsed comparisons

