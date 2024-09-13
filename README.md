# VBA-challenge
Module 2 Challenge - EXCEL and VBA Scripting


Under this Module, you will have to use VBA script to figure out following the thing:


1) The ticker symbol 
	(as you can see in excel, there is a lot of AAF, make it only one AAF and the only one next ticker and so on)


2) Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
   	(formula = open price (AAF, first date) - Close price (AAF, not the same date) = value (under the same Q1))



3) The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
	(formula = (percent Change = ((close price - open price) / open price)) * 100))


4) The total stock volume of the stock. 
	(this is a sum of all the total stock in the Q1 under the same ticker name AAF for example)



5) Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
	(this is a script to show max and min of percent change, also the mas of total volume into a new cells place WITH corresponding ticker name)


You will be mark under those requirements 
Requirements
Retrieval of Data (20 points)
The script loops through one quarter of stock data and reads/ stores all of the following values from each row:

ticker symbol (5 points)

volume of stock (5 points)

open price (5 points)

close price (5 points)

Column Creation (10 points)
On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:

ticker symbol (2.5 points)

total stock volume (2.5 points)

quarterly change ($) (2.5 points)

percent change (2.5 points)

Conditional Formatting (20 points)
Conditional formatting is applied correctly and appropriately to the quarterly change column (10 points)

Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:

Greatest % Increase (5 points)

Greatest % Decrease (5 points)

Greatest Total Volume (5 points)

Looping Across Worksheet (20 points)
The VBA script can run on all sheets successfully.
GitHub/GitLab Submission (15 points)
All three of the following are uploaded to GitHub/GitLab:

Screenshots of the results (5 points)

Separate VBA script files (5 points)

README file (5 points)