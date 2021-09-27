# Stock Market Analysis with VBA

## Overview of Project
Steve's parents want to invest in green energy companies, which is why they asked him to analyze and choose the best options, while Steve wants them to diversify their portfolio. This is where the analysis through VBA comes in: It is our job to create a VBA subroutine capable of summarizing the information provided by Steve while remaining fast and adaptable. 

A first iteration of the process was successful at analyzing the data but had some room for improvement, which is what will be discussed throughout this document


## Results
The original code would iterate over each row using _For_ loops and nested conditionals based on the ticker to calculate _totalVolume_, _startingPrice_ and _endingPrice_. Each time the loop ran the variables would be reset and recalculated. The results of this first version of the code can be seen next:

![All Stocks Analysis OG Code 2017](https://github.com/claud-e/Stock-analysis/blob/main/Resources/VBA_2017_OG.png) 

![All Stocks Analysis OG Code 2018](https://github.com/claud-e/Stock-analysis/blob/main/Resources/VBA_2018_OG.png) 

The second version of the code has one major difference: The use of arrays for the outputs instead of single variables. These arrays and a ticker index are declared as follows:

````
Dim tickerIndex as Single
    tickerIndex = 0
Dim tickerVolumes(12) as Long
Dim tickerStartingPrices(12) as Single
Dim tickerEndingPrices(12) as Single
````
Each iteration of the main loop now uses the _tickerIndex_ to access the arrays and make the calculations there. For example to increase the total volume:

````
For j= 2 to RowCount
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j,8).Value
.....


````
Once the algorithm determines that no more rows correspond to the current index the _tickerIndex_ is increased by one, which will make it access the next item in the array.

This change in how the variables are managed means the variables are created and accessible from the start, which will decrease the time needed to loop over the whole process. The results of this new process are as follow:

![All Stocks Analysis refactored Code 2017](https://github.com/claud-e/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) 

![All Stocks Analysis refactored Code 2018](https://github.com/claud-e/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png) 

It is clear that the refactored code works much faster, even when it has the added fuctionality of conditional formatting.

## Summary

1. What are the advantages and disadvantages of refactoring code?

- Advantages: in theory after refactoring the code should be more efficient and faster, with resources properly allocated and the same or improved results.

-Disadvantages: A certain level of proficiency is required to be able to refactor code, and if the code is not properly commented refactoring it could be a lenghty process that might end up reducing its functionality. 


2. How do these pros and cons apply to refactoring the original VBA script?

- The time required and functionality were improved, and while this change might seem small, if the data set were large enough the difference in performance would be necessary to make the analysis.

- After my first attempt I encountered a number of bugs I did not know how to solve, which means many hours were spent finding the appropriate changes. Also, while in this case it does not make much difference, the memory needed to create the arrays could become large enough that better hardware is required to run the code.
