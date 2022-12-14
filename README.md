### stock-analysis

## Overview
# This automated spreadsheet was created to enable our client Steve to generate data for his stock portfolio for 2017 and 2018. Our client would like to expand his portfolio and needed a workbook that would be better suited to handle more stocks. Refactoring our code would allow it to run more efficiently therefore being able to run through more tickers for Steves bigger portfolio. 

# Our clients data was for daily information for each ticker in his list. Using excel macros we were able to pull and organize each tickers data and give a return and other results for each desired years. Incorporating equations and for loops in our code to check the data allowed this to happen

## Results
# When refactoring our code, first a plan was created via steps in the comments to facilitate borrowing and repurposing code from a previous macro. THis enabled me to keep track of the flow of the logic. A timer was also setup to time our code and check to see if it did in fact run more efficiently. The results for each year using the refactored code are as follows:



## Summary
# When refactoring code we can see many advantages. Making the language more efficient can create an easier to understand flow of logic which is easier to build on when we need to make changes. This will also allow the program to run much faster freeing up memory and computing power while saving time. Disadvantges on the other hand include introducing new bugs into the code as well as breaking an already working script. In this case the benefits outweight the cons.
# THe original VBA script ran at around half a second for both years. As we can see from the results, the refactored code perfomed astronomically faster at 0.06 seconds. On a small portfolio like Steves this may not seem like a big difference but when working on much larger data set this refactored code can save a large amount of time. The previous code also had a harder flow to comprehend. The refactored code is cleaner and the logic is easier to follow when revisiting or making changes in the future.
