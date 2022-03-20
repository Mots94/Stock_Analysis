# Stock Market Analysis

## Using VBA to explore stock data

---
## Purpose
Steve works in finance and is searching for some specific information on stocks for his parents.  They are looking to invest in the stock market, and would like to know more about the total volume and yearly returns for the stock data that Steve has available.  In order to help Steve, we are searching an Excel workbook with stock data from the years 2017 and 2018.  This workbook has about 3000 rows of data for 12 stocks in both the 2017 and 2018 worksheets.  That is a considerable amount of data for analysis, so VBA was used to search our 2017 and 2018 worksheets for relevant financial data.  The original code created to accomplish this goal utilized a nested "for loop", which loops through all rows of data as well as a list of stock tickers.  This code runs fairly quickly for 12 stocks, but Steve would like to look at a much larger cross-section of the stock market for future analyses.  If our data set consisted of 40,000 rows instead of 3,000 the current code may take much longer to run.  Therefore, the challenge here was to re-factor our current code to only utilize one loop searching through data rows instead of looping through data rows and our list of stock tickers. 

---
## Methods
In order to utilize one loop that could capture all data for a stock ticker at once, there must be a way to reference which stock ticker is currently being searched for in the data.  This was addressed by creating a reference variable called tickerIndex with the code `tickerIndex = 0`.   This new variable was set to 0 since our tickers array begins at 0.  Additionally, three output arrays were created to hold the total volume, starting price, and ending price for each of the 12 stocks in our tickers array.  This was done using the code 
``` 
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single  
```
One more task that had to be accomplished before looping through the rows of data was setting all items of the totalVolumes output array to 0. 
```
For i = tickerArrayStart To tickerArrayEnd
        
    tickerVolumes(i) = 0
        
Next i
```
Finally, a code block was written to loop through all rows, using conditionals to add up tickerVolumes, find the tickerStartingPrices, and find the tickerEndingPrices for each ticker.  Within the conditional statements, the tickerIndex variable was used as an index to pull the correct data for each of the three output arrays.  This can be seen in the following code block
```
For i = 2 To RowCount
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
    End If
        
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerIndex = tickerIndex + 1
        
    End If
    
Next i
```
One other essential part of this code is the conditional statement at the end.  This statement increases the tickerIndex value by 1 if a new ticker symbol is found in a row following the current ticker.  Rather than having to loop through both the array of tickers and the rows in a spreadsheet, this index variable acts as a reference point so data can just be collected by looping through the data rows. 

---
## Results
When running the original code that was built throughout the module, the compilation time was about .81 seconds for both the 2017 and 2018 worksheets.
![VBAO2017](https://github.com/Mots94/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.PNG)
![VBAO2018](https://github.com/Mots94/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.PNG)
After refactoring the original code to its current state, the time to compile data dropped to about .11 seconds in both the 2017 and 2018 worksheets. This constitutes an 86% decrease in time taken to execute this code and compile data.  
![VBAR2017](https://github.com/Mots94/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Refactored.PNG)
![VBAR2018](https://github.com/Mots94/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored.PNG)
Even after completing this refactoring, the data captured for all 12 tickers remained the same as the original analysis
![DO2017](https://github.com/Mots94/stock-analysis/blob/main/Resources/Data_Output_2017.PNG)
![DO2018](https://github.com/Mots94/stock-analysis/blob/main/Resources/Data_Output_2018.PNG) 

---
## Summary
