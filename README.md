# Stocks Analysis 

In this challenge, we'll be helping Steve to analyze a stock market dataset with just a click and entry data for the year to run the analysis on, by means of Visual Basic for Aplications(VBA). 

## Overview of the project

Selecting and formatting the data to display the performance of the stocks in different years is done by coding variables and making conditions with the information on the cells. Also, we are going to determine the influence of refactoring a code, which is a readjustment for the computer to read it in a more efficient way, evaluating computing times.


## Results 
 
In this code, it needs to be run an analysis to evaluate the ticker information. It is done by comparing the value of the cells from the range of data for any input year, then with every different ticker, it looks up for the values of starting price, ending price and volume by the means of conditional 'If' statements. 

Every time theres a sequence of similar code, it could mean a sense that it's a process that can be automated. In this case, look up for the volumes, starting and ending price of all the stocks, are the variables that makes the code take the time to run. 

The way to see this is the nested 'for' loop in AllStocksAnalysis(), this function uses the loop to search for the data and assing a value to the ticker, volume and prices. As seen in the following code example, it nested a 'for' loop of ~3,000 values in a 'for' loop of 12 iterations, being the total of the repetition of the code more than 36,000 times. Compared to the AllStocksAnalysisRefactored(), the same ~3000 iterations for the row values and a few more 'for' loops to create the arrays that serves to collect information for display.   
```
  For i = 0 To 11
    
        ticker = tickers(i)
        
        totalVolume = 0
        
        For j = 2 To RowCount
```

The presented table is the analysis run through for every year, 2017 and 2018 accordingly.

![Stocks_2017vs2018.ong](/Resources/Stocks_2017vs2018.png)

From this tables, the formatting made it clear to see the performance of all the stocks, with a positive returns in ENPH and RUN for both years. Now Steve could have a better display or arrangement of the information to decide what to invest on.  

In the following pictures it could be appreciated the difference in running time of the code, as seen on the year 2017.
![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

Same case with year 2018.
![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

This makes it clear that a refactor process for this code, made it process the information for a fraction of the time. 

## Summary

### Advantages and disadvantages of refactoring a code

In general terms, refactoring a code could be very advantageous if it is a repetitive process where a lot of data is taken into considetation, making an iterative code run through the information and saving up time by assinging values. The counterside is that it might take time to analyze how the code can be improved and whether or not that improvement is something significant for the future use of the code.  

### How Pros/cons apply to refactoring the original code 

Even though the process might look a little bit more complicated by the making of arrays and many loops, this loops were in fact a shorter information for the computer to read and run, as it was assigning values meanwhile it was iterating. Not replacing information to overwrite the variables created.  
