# VBA scripting to analyze stock market data - Solution

Script location - https://github.com/kanhagithub/VBA-Challenge/blob/main/Private%20Sub%20multiple_year_stock().vbs

screeenshot - 

Quarter1 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q1%20screenshot.png

Quarter2 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q2%20screenshot.png

Quarter3 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q3%20screenshot.png

Quarter4 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q4%20screenshot.png


Retrieval of Data and calculation of quarterly change , % change and total stock volume 

The script loops through all the stocks data once and the following information is displayed.
   o ticker: The script stores distinct ticker symbol in one column in column "I" with a column header "Ticker”.
   o Total stock volume: The total stock volume is also generated on "L" column, this is done by adding "<vol>" column (G) values in loop for one ticker. 
   o	quarterly change: The script finds quarterly change by deducting opening price at the beginning of a given quarter to the closing price at the end of that 
   quarter and put the value on "J" column. 
   o percent change: The script also calculates percentage change which is (quarterly change)/(opening price) and put the value on "K" column. Then this column is formatted to represent percentage. 

Conditional Formatting 
•	The script applies an IF Condition to the quarterly change column and that highlights positive change in green and negative change in red in column J. Also the script 
Formats the percentage change values to display in percentage format by using Numberformat function in column "k".
  

Calculated Values 
•	Lastly the solution also provide the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".For this we use a for loop to go through the previous dataset.To calculate the gretest increase and decrease  values the script applies an IF condition to the percent change, and stores which value is greater. To get the percentage format of "Greatest % increase", "Greatest % decrease" the script applies a Numberformat function in column "P2" and "P3".


