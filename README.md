# Stock Analysis

## Performing analysis on stocks data to determine whether Daqo is worth investing in

### Overview of Project

The basis of this project was to help Steve research a handful of green energy stocks in addition to *Daqo New Energy Corporation*, a company that makes silicon wafers for solar panels. This is the company. To further his research, Steve decided to expand the dataset to include the entire stock market over the last few years. The original code we created worked well for a dozen stocks, but might not be efficient to analyze thousands of stocks. So, we were tasked with refactoring the code to loop through all the data one time, collect stock information for the years 2017 and 2018, and determine whether refactoring the code successfully made the VBA script run faster/more efficiently.

### Results

Following the steps outlined in the Module 2 instructions, I successfully refactored the solution code (see below):

    '1a) Create a ticker Index    tickerIndex = 0        '1b) Create three output arrays    Dim tickerVolumes(12) As Long    Dim tickerStartingPrices(12) As Single    Dim tickerEndingPrices(12) As Single        ''2a) Create a for loop to initialize the tickerVolumes to zero.    For i = 0 To 11        tickerVolumes(i) = 0    Next i    ''2b) Loop over all the rows in the spreadsheet.    For i = 2 To RowCount            '3a) Increase volume for current ticker        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value                    '3b) Check if the current row is the first row with the selected tickerIndex.                'If Then            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value            End If                    'End If                '3c) check if the current row is the last row with the selected ticker         'If the next row’s ticker doesn’t match, increase the tickerIndex.                 'If  Then                       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value                      End If            '3d Increase the tickerIndex.                        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then            tickerIndex = tickerIndex + 1                        End If                    'End If        Next i        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.    For i = 0 To 11                Worksheets("All Stocks Analysis").Activate             Cells(4 + i, 1).Value = tickers(i)             Cells(4 + i, 2).Value = tickerVolumes(i)             Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1            Next i

This code produced the same outputs as the original code, informing us that DQ had a Total Daily Volume of 35,796, 200, and a 199.4% Return in 2017. In 2018, DQ had a Total Daily Volume of 107,873,900 and a -62.6% Return. The refactored code ran significantly faster than the original code– 2017 refactored code run time was 0.1328125 versus original code run time of 0.4960938. The 2018 refactored code run time was 0.140625 versus the original code run time of 0.484375. The following screenshots show the run time of the refactored code:

![2017](https://github.com/MichaelaAnastasiaAustin/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018](https://github.com/MichaelaAnastasiaAustin/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Summary

#### Advantages and Disadvantages of Refactoring Overall
The advantage of refactoring code is that it helps make code run more efficiently—by reducing the number of steps, using less memory, or simplifying  the code to make it more user-friendly for future users. As a new student to VBA, the refactoring process also helped me become more acquainted with the code. A disadvantage of code might be how time consuming the process is. Also, if you are not familiar with the code, and are assigned to refactor it, you might risk making a significant change to the code that alters the desired outputs.

#### Advantages and Disadvantages of Refactoring Stock Analysis
 In our case, refactoring the original VBA code was successful in significantly reducing the run time (as detailed with screenshots above). By reducing the run time, we directly solved the concern listed in the Challenge 2 introduction– that the original code might take a long time to execute. One disadvantage is that the refactored code added several steps, so it might be difficult for an external user to follow.
