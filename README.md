# Green Stocks Analysis
## Overview of Project
This project analyzes **12 stocks** including the **DQ stock**, by measuring their ***average daily volumes*** and their ***returns*** for the years ***2017*** and ***2018*** using VBA

## Result
*The output of analysis of the stocks for 2017*
![2017 Image](/resources/VBA_Challenge_2017.png)

*The output of the analysis of the stocks for 2018*
![2018 Image](/resources/VBA_Challenge_2018.png)


## Summary

Arrays are defined to store the values need from the spreadsheet for each stock
```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
Then rows are traversed to find the starting price for each stock (represented by a ticker) and the ending price. The sum of the averages of the rows with the same tickers are calculated and stored for later use.
```
'3b) Check if the current row is the first row with the selected tickerIndex.
        If (Cells(i, 1) = ticker) And (Cells(i - 1, 1) <> ticker) Then
```
```
'3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If (Cells(i, 1) = ticker) And (Cells(i + 1, 1) <> ticker) Then
```
The information gathered from this worksheet is displayed in another worksheet called **All Stocks Analysis**.

