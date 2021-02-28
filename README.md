# Stock Analyis with VBA

## OVERVIEW: VBA Stock Analysis Project

### Purpose
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## RESULTS: Refactor VBA Code and Measure Performance
 
### Code Examples, Compare Stock Performance and Timestamp procedure below:

Based on the code, the stock performance was better in 2017 compared to 2018 overall. 

Refactored code

![AllStocksAnalysisRefactored](https://user-images.githubusercontent.com/76754655/109437618-bffd6280-79f3-11eb-93a1-b453f3e62ade.PNG)

![AllStocksAnalysisRefactored2](https://user-images.githubusercontent.com/76754655/109437623-ca1f6100-79f3-11eb-9bef-a9d71a98d462.PNG)

All Stocks 2017

![All Stocks (2017)](https://user-images.githubusercontent.com/76754655/109437512-23d35b80-79f3-11eb-9e57-22e4d015d409.PNG)

All Stocks 2018

![All Stocks (2018)](https://user-images.githubusercontent.com/76754655/109437523-2d5cc380-79f3-11eb-99d2-9a3a40b3762b.PNG)

The execution times for 2017 was better than 2018 as well for the refactored script. Further, the execution times of the original script were better than the refactored script.

Refactored execution time 2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/76754655/109437530-39488580-79f3-11eb-8e84-4146520bfdb1.PNG)

Refactored execution time 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/76754655/109437542-46fe0b00-79f3-11eb-83eb-aca31ae291c2.PNG)

Original execution time 2017

![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/76754655/109437561-64cb7000-79f3-11eb-8f57-cae83772cac7.PNG)

Original execution time 2018

![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/76754655/109437569-6c8b1480-79f3-11eb-98c9-bf5c5bb956c4.PNG)

## SUMMARY:

The advantages of refactoring code is that the code becomes more flexible for future edits in case there are additional tickers that may need to be analyzed as the array can add additional tickers easier than manually keeping track of the total number of tickers in the original script.

The disadvantages of refactoring code is the execution time is slightly longer and the array may not be as easily readible compared to the original script.

These pros and cons apply to refactoring the original VBA script by future maintenance required for the code in terms of if there are additional updates required then those updates would have to be made within the existing code for either the original script or the array script. In addition, if there are any new requirements for the original VBA script, it may be easier to make the edits from that original VBA script given the original requirements.
