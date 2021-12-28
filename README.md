# VBA Stock Analysis

## Overview of Project

### Purpose
The purpose of this project was to collect stock information on various stocks to calculate the yearly volume of the stocks and their returns to determine whether or not they are worth investing in.  The goal here was to refactor existing Excel VBA code to make the code more efficient and run quicker.  

### Data
The data set utilized in this analysis consisted of two tabs of yearly stock data.  Each tab included values for a stock ticker value, the date, the opening, high, low, closing, and adjusted closing price of the stock on that date, and the daily volume of the stock.

![data](https://user-images.githubusercontent.com/95199679/147530650-cd3c493b-f2bc-43fe-b066-3fb04f64fcb6.png)

## Results

### Analysis
The script first asks the user to input the year they would like to see the data outputs for:

![input](https://user-images.githubusercontent.com/95199679/147530380-96f166b8-1185-4d10-80aa-723d5d3402f3.png)

The data outputs include each ticker value, the total daily volume for each ticker, as well as the return.  After a year is keyed in, the refactored code creates output arrays that will display the stock data values.  Next, the total volume for each ticker is initialized to zero using a loop.  Then, each row in the data set is looped through to calculate the total volume of each ticker (by summing each ticker's daily volumes together), as well as the starting and ending prices for each ticker. Finally, the outputs are displayed on the "All Stocks Analysis" tab, as seen below:

![2017](https://user-images.githubusercontent.com/95199679/147529561-5359da47-2f12-4406-9e61-3b525454c930.png)
![2018](https://user-images.githubusercontent.com/95199679/147529565-ae1db261-22f0-44f3-90eb-32212ceebc5b.png)

As you can see, the Total Daily Volume for all stocks in 2017 is slightly less than all stocks in 2018, however the returns were significantly better in 2017 vs 2018 for all stocks except RUN and TERP.

Below are the run times for the refactored script for 2017 & 2018:

![VBA_Challenge_2017](https://github.com/bekkaross/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/bekkaross/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

These run times from the refactored code are faster than the run times for the original script (below):

![2017 original](https://user-images.githubusercontent.com/95199679/147530331-92201a82-d957-4c15-9779-80ef20a4b4f5.png)
![2018 original](https://user-images.githubusercontent.com/95199679/147530335-7a57d5f2-ed9e-4cfd-aa0b-682f1d91cd5f.png)

### Main Body of Refactored Code:

    '1a) Create a ticker Index
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If
                    
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

## Summary

### Advantages of Refactoring Code
The main advantages of refactoring code is to make the code more efficicent.  This can be accomplished in a number of ways, including taking fewer steps, improving logic to make code easier to read and understand for other users, and using less memory to help the code execute quicker.

### Advantage of Refactoring the Original Stock Analysis VBA Code
The main advantage of refactoring the original code was that it helped the code execute much quicker (by looping through all of the data only once).  As was shown above, the run times for each year were almost 6x as long in the original code vs the refactored code.  While the difference between 0.1 and 0.8 seconds may not seem hugely significant, the data set only contained values for 12 different stocks - the analysis time savings with the refactored code would be much more noticeable with a much larger data set.
