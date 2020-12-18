# VBA_Stock_Analysis
Learning to code in VBA using stock market data
##NOTE: There are comments written in my VBA code that align with the notes below.
##If reader cannot open the .bas files with the Macros code, the code is also placed in the uploaded Word file.

For this assignment, I needed to summarize for each stock ticker the yearly change, percent change, and total volume. 
Plus, the code had to applied across all three worksheets of data for quick automation.

Step-by-Step thought process for the FIRST SUMMARY TABLE:
1. Name the Macros: Sub StockMarketAnalysis()

2. Name the variables and assign the types

3. Looking at the raw data, there are a number of assumptions made. That the ticker data are in alphabetical order, chronological date order (January to December), and the data represent the data types. Thus, there is no cleaning required with regards to these areas.  

4. For each worksheet, set the variables that will count to zero

5. Create the column labels for the results to be entered. I call this the summary table for the Ticker, Yearly Change, Percent Change, and Total Stock Volume

6. Because each worksheet has varying number of rows, I used the code to start at the end of the worksheet and move up to the first row with data. Another way of thinking of it, this is getting the last row of data.

7. Next, needed to create conditional loop statements that compares the values from bottom to the top row (aka start row). 
First: figuring out the values for "Total Volume" for each ticker
  7a. There is some bad data, though. Where the volume is zero, the data in the other columns are static (the same). Thus, we do not want to include these rows in        the counts.
  7b. Once the zeros are taken care of, I then instructed the computer to look at all the non-zero total stock volumn for each ticker. The results were stored in a holding variable. 

Second: figuring out the yearly change and percdent change for each ticker
  7c. Calculates at the closing price at the end of the year minus the opening price at the beginning of the same year. 
  7d. Calculates the percent change ([(new amount - old amount)/ old amount]*100) with two decimal places out.
 
Third: Do 7a-7d again by begining on the next stock ticker

8. Place the results for each ticker in the same worksheet next to the raw data (columns I-L).

9. Help the reader quickly see stocks that were in the positive or negative at the end of the calendar year by placing conditional stylistic color coding. Green for in the positive and red in the negative. If there was zero (0) change, the cell will stay white.

10. Before closing the conditional loop, reset the counter for a new stock ticker.

11. Enter the closing commands for each the for statement, move on to the next worksheet when finished with the current one, and then end program.


BONUS Assignment
Somewhat similar to the main assignment by creating another summary table of the greatest percent decrease, greastest percent increase, and greatest volumne for each worksheet/year of data.

Step-by-Step thought process for the SECOND SUMMARY TABLE:
1. Name the Macros: Sub Bonus()

2. Name the variables and assign the types

3. Looking at the raw data, there are a number of assumptions made. That the ticker data are in alphabetical order, chronological date order (January to December), and the data represent the data types. Thus, there is no cleaning required with regards to these areas.  

4. For each worksheet, set the variables that will count to zero

5. Create the row labels (Greatest % Increase, Greatest % Decrease, and Greatest Volume) and column labels (Ticker and Value) for the results to be entered. I call this the second summary table. 

6. Figure out the last row from the first summary table bsaed on column "I".

7. Define two variables that will hold the data as the computer runs the code (currentTicker and Value)

8. For the three data points of interest, need to do comparisons to find out which is greatest between one row of data and the next, keeping the one that meets the condition (greater increase, greater decrease, greater volume). 

9. Please the results of the computaton in a new SECOND summary table (it will be placed to the right of the FIRST summary table).

10. Once completed, move on to the end worksheet and compute the same loop. End the Macros once there are no more worksheets.
