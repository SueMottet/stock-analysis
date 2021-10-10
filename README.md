# Stock Analysis in Microsoft Excel using VBA

## Overview
Using Microsoft Excel VBA scripts, analyze green energy stocks to a help a finance graduate Steve advise his parents on the best ones to invest in base on how often the stocks are traded and returns. Steve's parent have taken a particular interest in a green energy stock DAQO (ticker DQ) so he would like to look at that stock first but also compare it to other green energy stocks. Once he completes his green energy stock analysis for his parents, he is hoping to be able to use this spreadsheet configuration and VBA script for larger longer lists of stock data that he wants to analyze. With that in mind, he would like the VBA script to be refactored after the initial green energy stock analysis to improve run time as much as possible

### Steps completed to perform intial green energy stock analysis
1.  Download green enery stock data
2.  Enable macros
3.  Create a worksheet for the analysis
4.  Write a VBA script (macro) to calculate DQ's stock daily and yearly volumes and to calculate DQ's stock yearly return for 2018
6.  Provide DAQO results to Steve
7.  Write a reusable VBA script (macro) to calculate daily and yearly volumes and to calculate yearly return for 2018 and 2017 for any additional handful of green energy stock
9.  Format the results data to make it easier for Steve to read
10. Provide a button in the analysis spreadsheet to prompt for a year to run for and then run the analysis for entered year as needed
11. Provide a button to clear the worksheet prior to rerunning when needed
12. Provide spreadsheet with the buttons and formatted results to Steve on the handful of green energy stocks including DQ

### Refactoring completed to tune the VBA 
In the case of this project, the refactoring goal was to tune the macro for optimum performance for analyzing larger longer lists of stocks. The initial VBA script iterated through the data in the data set for one stock and then output the results for that stock to the analysis spreadsheet. 


The refactored VBA script instead creates output arrays and captures the output data in them efficiently. Using these output arrays, allows this refactored VBA script to iterate through the rows in the data sheet only once and then output the result efficiently be pulling the data out of the arrays versus having to go back to the data sheet over and over again.

![Refactored code snippit](/Resources/Refactor_code_snip.png)  

![2018 refactored time](/Resources/VBA_Challenge_2018.png)  

![2017 refactor time](/Resources/VBA_Challenge_2017.png)  

## Project Resources
- Data Source: green_stocks.xls
- Software: Microsoft Excel Ofiice 365

## Summary
Refactoring code requires not only understanding what the code is used for and will be used for but also having a good working knowledge of coding options.The goal of refactoring is to improve the sturcture of the code without changing it's overall functionality. 

Improved maintainablity, improved performance, increased scalability and making the code more secure can all be reasons to refactor. The point of the refactor the intial code for this project was to improve preformance decreasing the run time of the VBA script.



### Refactoring Advantages

### Refactoring Advantages

### Stcok Analysis highlights
The analysis of the stocks show that:
- DQ had high returns in 2017 and poor returns in 2018
- That most of the stocks had positive returns in 2017 excel TERP
- Only two stocks had positive returns in 2018 and those two were ENPH and RUN
