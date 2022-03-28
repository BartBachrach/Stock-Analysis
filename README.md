# Stock-Analysis
# **VBA Challenge**

##  **Project Overview**
The objective of this project was to refactor code we have previously written to include all stocks from years 2017 to 2018. The previous code was written to include one stock for those years, DAQO; the refactored code goes through all the stocks for those two years, so it required some adjustment. The refactored code was also supposed to run faster than the previous code, which for one instance it did but mostly ran at about the same speed as the previous code. The reason we used Visual Basic for Applications (VBA) was because we wanted to design this code to work on the list regardless of whether additional information was added, and could again be refactored for future use. The advantage of using VBA is that, unlike regular excel formulas that depend on using hard-coded instructions, VBA is flexible and can be programmed to adjust to its new spreadsheets even if they’re different sizes. 

## **Code Refactoring**
One of the blocks of code that needed to be refactored was the array declaring the conditions the for-loop would run through. Initially it was set to just run through all the times DAQO had been traded and its prices, but this time we had to include all 12 different stock tickers, indexed in the array from 0-11. 
Another difference between this code and the original code is that we wanted to know three things about the stocks: starting prices, ending prices, as well as the total volume on the stocks (how many times they were traded during that year). We needed the starting and ending prices because the difference between the starting and ending prices would give us their annual return, adding another level of insight into their performance. 

## **Code Performance**
	The time difference between the original DQ analysis code and the refactored code can be explained as follows: the raw time it took to perform both analyses initially was about the same. However, the refactored code includes more steps and runs through much more information (12 times as much to be exact). The code now looks like the following.

`Dim tickers(12) As String
tickers(0) =
tickers(1) =
tickers(2) =
tickers(3) =
tickers(4) =
tickers(5) =
tickers(6) =
tickers(7) =
tickers(8) =
tickers(9) =
tickers(10) =
tickers(11) =`
 and returns the results in about the same amount of time. Given that information the code was much more efficient and each step required less time to complete than the original code. After the program had run through once, the next time it ran it took much less time than it did the first, completing in about .05 seconds. 
	

## **Results**
The code refactoring was successful, with the program executing the expanded code in the same amount of time. The code was only expanded to the extent required which contributed to the success of the refactor. 
Regarding the stocks themselves, it seems 2017 was a much better year for this series of stocks than 2018, and the young man should consider investing his money in other stocks. ![This is an image]![Uploading Screen Shot 2022-03-26 at 11.49.55 AM.png…]


