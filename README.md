#Stock Market Analysis

##Using VBA to explore stock data
---
##Purpose
Steve works in finance and is searching for some specific information on stocks for his parents.  They are looking to invest in the stock market, and would like to know more about the volume and yearly returns for the stock data that Steve has available.  In order to help Steve, we are searching an Excel workbook with stock data from the years 2017 and 2018.  This workbook has about 3000 rows of data for 12 stocks in both the 2017 and 2018 worksheets.  That is a considerable amount of data for analysis, so VBA was used to search our 2017 and 2018 worksheets for relevant financial data.  Within the code for the original analysis, a nested "for loop" was used to loop through all rows of data as well as a list of stock tickers to search for relevant volume and pricing information.  This analysis runs fairly quickly, but our challenge is to find a way to make the data compile even faster by only using one loop instead of two.  
##Methods
In order to accomplish what Steve is asking for, a detailed VBA subroutine was written to analyze stock data for 2017 and 2018.  An overview of this code can be explained in four major parts: year input, looping through data, output of stored data, and formatting output sheet.  In the first part, year input, code was written to create a variable that allows the individual running this macro to input the year for analysis.  This variable is then used to activate either the 2017 or 2018 worksheets for analysis.  Then    
---
##Results
A detailed VBA subroutine was written for Steve 
