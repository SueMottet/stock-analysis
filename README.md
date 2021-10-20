# Improving performance of Microsoft Excel VBA scripts using arrays

## Overview
Refactor a VBA script used for stock analysis to improve run time and increase script usability for larger datasets.

### Project Background
Using Microsoft Excel VBA scripts, the initial VBA script automated gathering and formatting data for the analysis of green energy stocks to a help a finance graduate Steve advise his parents on the best ones to invest in based on how often the stocks are traded and returns. Steve's parents had taken a particular interest in a green energy stock DAQO (ticker DQ) so he requested looking at that stock first but also wanted to compare it to a handful of other green energy stocks. Once he completes his green energy stock analysis for his parents, he needs to be able to use this spreadsheet configuration and VBA script for larger stock dataset that he wants to analyze. With that in mind, he would like the VBA script to be refactored after the initial green energy stock analysis to improve run time as much as possible.

### Project Resources
- Data Source: green_stocks.xls
- Software: Microsoft Excel Office 365

### Steps completed to perform initial green energy stock analysis
1.  Download green energy stock data
2.  Enable macros
3.  Create a worksheet for the analysis
4.  Write a VBA script (macro) to calculate DQ's stock daily and yearly volumes and to calculate DQ's stock yearly return for 2018
6.  Provide DAQO results to Steve
7.  Write a reusable VBA script (macro) to calculate daily and yearly volumes and to calculate yearly return for 2018 and 2017 for any additional handful of green energy stock
9.  Format the results data to make it easier for Steve to read
10. Provide a button in the analysis spreadsheet to prompt for a year to run for and then run the analysis for entered year as needed
11. Provide a button to clear the worksheet prior to rerunning when needed
12. Provide spreadsheet with the buttons and formatted results to Steve on the handful of green energy stocks including DQ

#### Stock analysis highlights
The analysis of the stocks shows that:
- DQ had high returns in 2017 and poor returns in 2018
- That most of the stocks had positive returns in 2017 except TERP
- Only two green energy stocks had positive returns in 2018 and those two were ENPH and RUN

### Steps completed to perform refactoring
1. Create output arrays and alter initial code using these arrays to capture results for the data more efficiently
2. Change results output to leverage arrays

## Refactoring Results

### Refactoring background
Improved maintainability, improved performance, increased scalability and making code more secure can all be reasons to refactor. The point of the refactor of the initial code for this project was to improve performance decreasing the run time of the VBA script.

#### Initial code summary:
The initial VBA script iterated through the data in the data set for one stock and then output the results for that stock to the analysis spreadsheet. 

##### Initial code sample:
![Initial_code snippit](/Resources/Initial_code_snip.png)  

##### Initial 2017 run time

![2017 initial time](/Resources/VBA_Initial_2017.png)  

##### Initial 2018 run time

![2018 initial time](/Resources/VBA_Initial_2018.png)  

#### Refactored code summary:
The refactored VBA script instead creates output arrays and captures the output data in them more efficiently. Using these output arrays, allows this refactored VBA script to iterate through the rows in the data sheet only once and then output the result efficiently by pulling the data out of the arrays versus having to go back to the data sheet over and over again.

##### Refactored code sample:

![Refactored code snippit](/Resources/Refactor_code_snip.png)  

##### Refactored 2017 run time

![2017 refactor time](/Resources/VBA_Challenge_2017.png)  

##### Initial 2018 run time

![2018 refactor time](/Resources/VBA_Challenge_2018.png)  

## Refactoring Summary
Refactoring code requires not only understanding what the code is used for and will be used for but also having a good working knowledge of coding options. The overall goal of refactoring is to improve the structure of the code without changing it's functionality. 

### General Refactoring Advantages
1. Leverage new technologies that were not available when code was initially written
2. Improve code with set goals in mind: improved maintainability, improved performance, increased scalability and/or making code more secure
3. Potential reduction in complexity and readability

### General Refactoring Disadvantages
1. Time investment into code that works as is
2. Requires impact analysis, release management and testing rework to go with change 
3. Risk of the introduction of new unintended functionality changes
4. Interoperability concerns and risks

### Project Refactoring Advantages
1. Refactoring can improve performance
2. Refactoring with performance enhancements in mind can expand the usability of code

### Project Refactoring Disadvantages
1. Refactoring requires a good working knowledge of coding options that may make this code harder to maintain for less skilled programmers
2. Refactoring code can increase code complexity

### References
* VBA documentation https://docs.microsoft.com/en-us/office/vba/api/excel.font%28object%29
* Basic writing and formatting syntax for GitHubhttps://help.github.com/en/articles/basic-writing-and-formatting-syntax
