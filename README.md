# VBA scripting to analyze stock market data - Solution

Script location - https://github.com/kanhagithub/VBA-Challenge/blob/main/Private%20Sub%20multiple_year_stock().vbs

screeenshot - 
Quarter1 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q1%20screenshot.png

Quarter2 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q2%20screenshot.png

Quarter3 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q3%20screenshot.png

Quarter4 - https://github.com/kanhagithub/VBA-Challenge/blob/main/Q4%20screenshot.png


Retrieval of Data and 
•	The script loops through one quarter of stock data and stores all the values from each row ticker, volume of stock, open price and close price.
Column Creation 
The script loops through all the stocks data once and the following information is displayed.
o	ticker: The script will sort the distinct ticker symbol in one column in column "I" with a column header "Ticker”.
o	Total stock volume: The total stock volume is also generated on "L" column. 
o	quarterly change: The script will execute quarterly change from opening price at the beginning of a given quarter to the closing price at the end of that quarter and put the value on "J" column. For this task the code added a conditional formatting that highlighted positive change in green and negative change in red.
o	 percent change: The script also percent perform a change from opening price at the beginning of a given quarter to the closing price at the end of that quarter and put the value on "K" column.



Conditional Formatting 
•	The script applied an IF Condition correctly and appropriately to the quarterly change and percent change column. 



Calculated Values 
•	Last, not least the solution also provide the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".


