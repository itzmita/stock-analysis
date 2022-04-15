
# Green Stock Analysis
In this Module lesson we are learning how to use VBA to analyze stock data.

## Project Overview

### Background
In this project, Steve's parents are planning to invest in Green energy stock and we are helping Steve to analyze the stock data in order for him to have a better understanding to help his parent's investments. The Module lesson covered analyzing the data for one year. However, this assignment is to expand the dataset to include the entire stock market, over 2017-2018.
### Purpose
The purpose of this analysis was to help Steve understand which green energy stock would be better for his parents to invest. His parents are interested in the DQ stocks, so we did our analysis using only this Ticker "DQ". We looked into the performance for this stock over the year of 2017 and 2018 and compared it with other stocks as well. We used VBA code to derive Total Daily Volume and Yearly Return for all the 12 stocks including DQ and measure the performance of running the script. In this final challenge, we are re-factoring the earlier VBA code to increase performance so that it can be run faster with more efficiency with more data and not just these 12 stocks. 


## Results
Displayed below are the images of 12 green energy stocks: Analysis was done considering the below few fields from the dataset.
	• Ticker name
	• Total Daily Volume for a given year
	• Percentage of the yearly return for each stock in a given year	

![image](https://user-images.githubusercontent.com/3753839/163502555-df1e172e-8ac5-49fb-85df-d337afda5c91.png)
![image](https://user-images.githubusercontent.com/3753839/163502568-02bae3d8-dcab-45f0-88c8-21d52dff6223.png)

If you see the above data, 2017 had been a good year for the stocks in consideration for this project compared to 2018.
Most of the stocks except TERP shows positive yearly return in 2017. Daily volume for DQ has been lower than 2018. Also DQ has a yearly return of almost 200% which might indicate it’s a good stock to invest. But when we see the year for 2018, even though the Daily volume has increased for DQ which means more transactions has been happening, the yearly return is showing negative 63% which indicates it’s a risky investment as its not stable.  However if we see ENPH which has increased in Daily volume and also had positive Yearly returns might be a good investment to look for. Considering the volatility diversification and investment in multiple stocks might be a better option for Steve's parents.


## Comparing codes
The original VBA solution in "AllStockAnalysis” and the refactored code in “AllStockAnalysisRefactored” have the same output in terms of Yearly returns and Total Daily volumes. We are refactoring this code in order to make it better at performance and better readability of the code as well. 

The  "AllStockAnalysis"  code included two loops. the inner loop with iteration "i" calculates the Ticker Daily volume and the Return. These are considered in variables. The outer loop with iteration "t" has the control shifted to "All Stock Anlaysis" worksheet to output the data for each of the Tickers. For each of the 12 Tickers the control keeps going back and forth between worksheets to calculate in one sheet and then to output that data in another sheet and again come back to calculate the next Ticker.

![image](https://user-images.githubusercontent.com/3753839/163502661-21d21f31-43ee-40d1-a911-0778b853dbb8.png)

The re-factored VBA solution had arrays for Daily volume, Yearly returns and Tickers and not variable like the earlier code had. The code was also modified to have one loop where all the calculations are done and stored in arrays which were indexed by a new variable called tickerIndex. Once the data is fetched for all the stocks and arrays are populated, then the output worksheet is activated to write the results in the other sheet. This reduced the back and forth between worksheets like it was happening for the earlier version of the code. 

![image](https://user-images.githubusercontent.com/3753839/163502738-cafbadb4-3c13-49e1-9158-4f1092153f33.png)


Both the scripts "AllStockAnalysis" and "AllStockAnalysisRefactored" have the same output in terms of it's values for each green stock.
### Execution time for the code
When the refactored code was executed against 2017 and 2018 stock market data set, both ran in approx 0.12 seconds with re-factored code as compared to the original code that ran approx in 7-8 seconds, which is so much slower than the refactored code.


Earlier code for the year of 2017 ran as showed below

![image](https://user-images.githubusercontent.com/3753839/163502783-d80abb07-9751-46ed-abdf-f2ff96322ad8.png)


But the refactored code for 2017 ran in much less time

![image](https://user-images.githubusercontent.com/3753839/163502800-ee9a8421-fbda-4795-953c-8303b0922922.png)



 The same when run for 2018 with earlier code showed below run time
 
![image](https://user-images.githubusercontent.com/3753839/163502811-51320ea0-3a6d-4bfb-a67d-82030554ea16.png)


And after the refactoring it showed significant improvement as you see below

![image](https://user-images.githubusercontent.com/3753839/163502824-1d292ab9-71fb-4a44-80ed-813aa2c6bcdd.png)


## Summary
1. What are the advantages or disadvantages of refactoring code?
	• Advantages of refactoring a code - 
		○ Performance - refactoring allows us to modify code in an optimal way and improve performance. Without changing the functionality of the code, we look for recoding it to have better methods, remove redundant logic and make functions which can be reusable by multiple steps. This increases performance.
		○ Readability - by refactoring a code, we can write it in a cleaner manner for better understanding and readability. We can write better error handling methods to enable efficient post production operations. 
		○ Fixes - Refactoring allows us to fix mistakes in the code or remove dead codes if present. 
	• Drawbacks of refactoring a code -
		○ Time consuming - if we haven't written a code and been asked to refactor, it’s a huge challenge as we may not understand the logic and it can take a long time to understand the code. If there are no well written documentation within the code, then it is all the more difficult. That’s why manual refactoring is the least preferred method for any type of modernization/migration projects. 
		○ Risky - Refactoring a code may end up impacting the Business logic if we don’t understand how and why it was written in the first place which is very risky for Business purposes. 
		
2. How do these pros and cons apply to refactoring the original VBA script?
	• Refactoring a code is a huge task when it comes to real life production code. The assignment given in this project was very well documented with each step of the logic. The guidance provided in the Module was very detailed which doesn’t usually happen in an actual production scenario. 
	• We could very well see the advantages of refactoring 
		○ As the performance was immensely improved when I ran the refactored code in comparison to the earlier code
		○ Readability was better as we could avoid the nested for loop
	• Cons for this refactoring at least for me was that it was time consuming for me not, not knowing the logic, but being new to VBA coding it was challenging for me. But the documentation was great so it helped me a lot to finish this challenge. Still I faced some difficulties getting the desired output for some time because for example I had the 2018 sheet filtered by the Ticker "DQ", so it was not giving me correct outputs. 

