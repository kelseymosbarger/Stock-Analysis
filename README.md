# VBA Challenge

## Overview of Project

### Purpose

The purpose of VBA Challenge is to write a VBA code to analyse returns of stocks for 2017 and 2018 and create a summary 
output comparing the volume and return for each stock for Steve's parents to make a sound investment decision. THe VBA 
code uses inputs and arrays in for loops and if then statements to sumaraze and analyze the data.

## Results

### Results of Stock Data

Clearly 2017 was a better year for stocks than 2018, as 2017 had more "green" returns or positive returns compared to 2018 
for the same stocks. This analysis started by only analyzing DQ data, and though the 2017 return or 199% seems promising it 
had a negative return, validating the next step of comparing it to all stocks over the 2 year. There are only 2 stocks that 
help a positive return in both 2017 and 2018, ENPH and RUN - these are the 2 stocks that Steve should encourage his parents 
to invest/research those stock more rather than invest in DQ. 

See link to 2017 and 2018 screenshot of stock results for 2018 and 2017 and run time of hte refactored code.
https://github.com/kelseymosbarger/Stock-Analysis/commit/854336be5fefa79525f210c8cced4829f36b19ba



### Results of code differences
Both sets of "all stock analysis" delivered the same outcome. However the original code took more time as it is a nested for loop.
so the overarching result of between the two codes is that the refactored script can run much faster, and uses less processing power
this refactored code could work with much larger datasets, where as the orginal script would increase in time and processing power
as your data set got larger.

## Conclusions

### Advantages / Disadvantages of refactorings the code in general
The advantage of refacorting is to make the code as efficient as possible. In this case to not force the line level loops to reiterate
each time but to define arrays and an index so teh redundant part in the process is taken out to drive faster results. They both follow
a similar code structure, but in the original it is a nested for loop and the refactored is a series of 3 for loops, in which one of them
replicated the result as the nested for loop.


### Advantages / Disadvantages of original and refacored VBA script
The original process was more intuitive and eaiser to understand and test each step.Once I fully understood how to work with the 
index and the arrays, I kept getting an overflow error. VBA Debugger highlighted my row in #4 in the code as the source of the error. 
However after working through it with the tutor and other students, the cause of the overflow is beacuse the if then statements had 
an AND in it which was causing an error with the refactored code, as the AND is not necessary when using an index in the loop.