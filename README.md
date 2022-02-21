# Stock Analysis

**OVERVIEW OF PROJECT**
The purpose of this analysis was to determine which stocks are "green", or worth investing in, using Microsoft Excel and VBA.  Twelve stocks were selected and their performance (in both 2017 and 2018) was analyzed to determine whether they would be wise investments.  Total daily volume and return on investment were compiled to determine the profitability of each stock.  In order to ensure more efficient and orderly analysis, refactoring (or structuring the code to run more efficiently) was used in the VBA code.  The runtimes for each version of VBA code (original vs. refactored) were compared to demonstrate the benefit of code refactoring.

**RESULTS**
After running the original VBA code for the year 2017, the following information was calculated:
![2017 Table](https://user-images.githubusercontent.com/99574730/155021376-5c51d7f9-b725-4cc5-b63e-db69d36bd13f.png)

From this data, it appears that all 12 stocks except for TERP generated positive return on investment.  In essence, 11 stocks would have been wise investments if Steve's aprents had invested in them in 2017.

For the year 2018, the total daily volume and return of the 12 stocks was the following:
![2018 Table](https://user-images.githubusercontent.com/99574730/155021530-ae792679-7308-47be-a4b6-d8b46e8e3681.png)
In 2018, a different general trend was observed.  With the exception of ENPH and RUN, all other stocks generated negative return on investment.  In other words, those who invested in AY, CSIQ, DQ, FSLR, HASI, HKS, SEDG, SPWR, TERP, and VSLR in 2018 lost money.  Therefore, though they may have decided to invest in the above stocks in 2017, Steve's parents should not have invested in them in 2018.

In addition, the following runtimes were given for the **original, nonrefactored** code:

![2017 Original](https://user-images.githubusercontent.com/99574730/155034128-65ad43a8-f455-4d26-af7c-10118551bda0.png)
![2018 Original](https://user-images.githubusercontent.com/99574730/155034240-cbfbeb15-a523-40b9-b7e7-c380b76dc5b1.png)

After refactoring the VBA script, the following runtimes were observed:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/99574730/155034492-c4540cdd-d9d2-41ef-9200-cee7fbb4b290.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/99574730/155034373-5dbc8512-3e7a-43e3-a1d2-5e392d8e24d4.png)


As seen above, the refactored code for 2017 data ran roughly 7 times faster than unrefactored code.  In 2018, the refactored code ran roughly 10 times faster.



**SUMMARY**
Refactoring code has many benefits, including more efficient runtime as demonstrated in this challenge.  In addition to faster execution, refactoring allows for code to be written more concisely, allowing for programmers to finish a task more quickly.  Since refactoring allows programmers to perform the same function in fewer steps, it also limits the possibility of errors due to a shorter amount of code.

However, refactoring also presents certain disadvantages.  For less experienced programmers, refactoring may be difficult and result in bugs or opportunities for error that could be avoided in a more familiar format (despite a less efficient execution).  While completing this challenge, for example, I had several bugs to correct while refactoring.  In addition, refactoring can be a time-consuming exercise.  Although it allows for code to execute more efficiently, following proper syntax and debugging (if necessary) would likely take more time than simply writing code in a less sophisticated format.
