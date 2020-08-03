# stock-analysis
Perform stock Analysis by using Excel Macro with Visual Basic language

Challenge

Project Background
The purpose of this stock analysis is to help Steve compare Total Daily Volume and Yearly Return of each target stock in particular year. By designing a  Macro to calculate returns for different years. Refactoring of code is meant to test speed of macro and to gauge whether the VBA script will run faster.  

Conclusion
In 2018, ENPH and RUN stocks had positive yearly Return as well as large Total Daily Volume. 
![2018 Stock Analysis](https://user-images.githubusercontent.com/59589015/89142649-01896600-d516-11ea-9f03-89110fe1a924.png)


In 2017, all of stocks had positive Return except TERP (-7.2%). 


￼![2017 Stock Analysis](https://user-images.githubusercontent.com/59589015/89142675-149c3600-d516-11ea-98e5-2e2cd947691a.png)


Program Design
There are Four Loops:
* (A) is the Main Loop for going through all data and assigned tickerIndex for 12 stock.
* (B) is a nested loop in the main loop (A), go through stocks original data and get ticker name, startingPrices and endingPrices, and save information to each related tickerIndex.
* (C) a nested loop in (B) loop, in order to get volume information for each Index.
* (D) a new loop for putting all saved output information into an analysis sheet.

Logical Flow
1. Request users input which year they would like to analyze stock performance.
	yearValue = InputBox("What year would you like to run the 	analysis on?") 
2. Create and activate an analysis worksheet to keep all information retrieved.
3. Declare 1 array for ticker and 3 outputs arrays for saving data, as well as a variable named tickerIndex.
    
4. Create a main loop to assigned tickerIndex from 0 to 11. Initialize index as zero before loops.
    tickerIndex = 0
    For tickerIndex = 0 to 11
        if meet some criteria then
            tickerIndex = tickerIndex + 1
    Next tickerIndex

5. Make a loop go through all stocks data.
    Worksheets(yearValue).Activate
  
6. Make a nested loop to get incremental Daily volumes for each stock, then put into Volume(tickerIndex).

7. Create new loop for putting outcomes into analysis Worksheet which created on step 2
    	Worksheets("Challenge_All Stocks Analysis").Activate

8. Apply Font Formatting and conditional color Formatting to analysis Worksheets
