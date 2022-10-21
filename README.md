# Stocks Analysis Template

In this challenge, we'll be helping Steve to analyze a stock market dataset with just a click and entry data for the year to run the analysis on, by means of Visual Basic for Aplications(VBA). 

##Overview of the project

Selecting and formatting the data to display the performance of the stocks in different years is done by coding variables and making conditions with the information on the cells. Also, we are going to determine the influence of refactoring a code, which is a readjustment for the computer to read it in a more efficient way, evaluating computing times.


##Results 

(usar donde?) This code is a complex of different conditions and variables, by using Cells values it can be done the display of the information we will be working on. 

In this code, it needs to be run an analysis to evaluate the ticker information. It is done by comparing the value of the cells from the range of data for any input year, then with every different ticker, it looks up for the values of starting price, ending price and volume by the means of conditional 'If' statements. 

Every time theres a sequence of similar code, it could mean a sense that it's a process that can be automated. In this case, look up for the starting and ending price of all the stocks, its the variable that makes the code take the time to run. 

The way to see this is the nested 'for' loop in AllStocksAnalysis(), this function uses the loop to search for the data and assing a value to the ticker, volume and prices. If we assign the values to where to look before, the machine only look for that specific values instead of taking time to assing a new value after each search. Putting it simple, when you know what to look for, you dont waste time evaluating other data. Similar to establishing conditions  

###Images and examples of code

Compare stock performance 2017 vs 2018 as well as execution times for the code.

![Stocks_2017vs2018.ong](/Resources/Stocks_2017vs2018.png)

From this tables, the formatting made it clear to see the performance of all the stocks, with a positive returns in ENPH and RUN for both years. Now Steve could have a better display or arrangement of the information to decide what to invest on.  

In the following pictures it could be appreciated the difference in running time of the code, as seen on the year 2017.
![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

Same case with year 2018.
![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

This makes it clear that a refactor process for this code, made it process the information faster. 

## Summary

###Advantages and disadvantages of refactoring a code

In general terms, refactoring a code could be very advantageous if its a repetitive process where a lot of data is taken into considetation, making an iterative code run through the information and saving up time by assinging values. The counterside is that it might take time to analyze how the code can be improved and whether or not that improvement is something significant for the future use of the code.  

++ add that nested for loops though helpful, it might take more memory to process in the ram 

###How Pros/cons apply to refactoring the original code 

Even though the process might look a little bit more complicated by the making of arrays and many loops, this loops were in fact a shorter information for the computer to read and run, as it was assigning values meanwhile it was iterating. Not replacing information to overwrite the variables created.  
