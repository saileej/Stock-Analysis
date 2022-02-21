# Stock Analysis

**OVERVIEW OF PROJECT**
The purpose of this analysis was to determine which stocks are "green", or worth investing in, using Microsoft Excel and VBA.  Twelve stocks were selected and their performance (in both 2017 and 2018) was analyzed to determine whether they would be wise investments.  Total daily volume and return on investment were compiled to determine the profitability of each stock.  In order to ensure more efficient and orderly analysis, refactoring (or structuring the code to run more efficiently) was used in the VBA code.  The runtimes for each version of VBA code (original vs. refactored) were compared to demonstrate the benefit of code refactoring.

**RESULTS**
After running the **original** VBA code for the year 2017, the following information was calculated:
![2017 Table](https://user-images.githubusercontent.com/99574730/155021376-5c51d7f9-b725-4cc5-b63e-db69d36bd13f.png)

In addition, the following runtime was given for original 2017 code:
![2017 Original](https://user-images.githubusercontent.com/99574730/155021482-10e49666-1408-4192-93ca-e0fb206f8ca5.png)

In 2018, the total daily volume and return of the 12 stocks was the following:
![2018 Table](https://user-images.githubusercontent.com/99574730/155021530-ae792679-7308-47be-a4b6-d8b46e8e3681.png)

Using the **original (nonrefactored)** code, the runtime was approximately .9 seconds:
![2018 Original](https://user-images.githubusercontent.com/99574730/155021605-0563addd-485a-47f1-bd37-21e4992a8f2f.png)


**SUMMARY**
Refactoring code has many benefits, including more efficient runtime as demonstrated in this challenge.  In addition to faster execution, refactoring allows for code to be written more concisely, allowing for programmers to finish a task more quickly.  Since refactoring allows programmers to perform the same function in fewer steps, it also limits the possibility of errors due to a shorter amount of code.

However, refactoring also presents certain disadvantages.  For less experienced programmers, refactoring may be difficult and result in bugs or opportunities for error that could be avoided in a more familiar format (despite a less efficient execution).  While completing this challenge, for example, I had several bugs to correct while refactoring.  In addition, refactoring can be a time-consuming exercise.  Although it allows for code to execute more efficiently, following proper syntax and debugging (if necessary) would likely take more time than simply writing code in a less sophisticated format.
