# VBA of Wall Street stock-analysis

## 1.	Overview of Project: 
### Explain the purpose of this analysis.
Excel may be enough to run quick analysis on different kinds of industries. But as the world moves forward so the data and its complexity to be analyzed. In this module we learned how to use VBA, which is a programming language built in Excel that allows you to run complex projects while reducing the chances for errors and looking for automations that will leave you with more time to actually analyze the data than spending time in repetitive tasks.
So as we dive into programming we may encounter different scenarios were existing code might not be the best way to solve a specific problem, and as the problem grows in complexity also the code needs to be as lean and as efficient as possible to optimize computing resources. This is the case that we are solving with the stock analysis.
As per the introduction to this module we know Steve is looking for help to understand how stocks have been behaving over the years in order to invest his parents money. Let’s help him out!
## 2.	Results: 
### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.



| All Stock Analysis  | All Stock Analysis Refactored |
| ------------- | ------------- |
| 2017 (3.41 s) | 2017 (0.59 s)  |
| 2018 (3.25 s) | 2018 (0.55 s)  |

As we can notice refactoring improves the performance of the code. In this case the improvement goes up to 82% less time. With a similar result on the 2018 analysis. 

<img width="227" alt="2017" src="https://user-images.githubusercontent.com/31755703/149712435-e0c3331b-c88d-47f0-9beb-8ae3ee2d25cb.png">
<img width="214" alt="2018" src="https://user-images.githubusercontent.com/31755703/149712440-fa73f61b-96f0-45de-b110-30ea1b872e81.png">

<img width="236" alt="2017 refactored" src="https://user-images.githubusercontent.com/31755703/149711913-b9b0ddb0-9146-4237-804d-df952b759b10.png">
<img width="235" alt="2018 refactored" src="https://user-images.githubusercontent.com/31755703/149711919-80078d15-744a-4c88-8ff3-aeb779da9553.png">

One example of refactoring in the code is the fact of automating tasks as finding the value of the last cell without a fixed (or magical) number

RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
Another example is the use of variables to get the code organized in different arrays instead of using lines and lines of code to store values that otherwise can be managed through the use of these tools.


    '1a) Create a ticker Index
    tickerIndex = 0
   

    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single

    For i = 0 To 11
      tickerIndex = tickers(i)
      tickerVolumes = 0

## 3.	Summary: 
### In a summary statement, address the following questions.

Some of the advantages of refactoring code is that you cand find easier ways to perform a specific action as well as to make the code more readable so in case you revisit your code in the future you will spend less time trying to remember what a line of code was intended to do. Also you may optimize the computing resources as the problems grow in the complexity.
I consider that as you gain some experience the disadvantages of refactoring will be less. Sometimes it is difficult to find a better way to perform certain actions or you may find yourself without the knowledge or tools needed to improve the code. 
Refactoring is time consuming but we need to see it as an investment of time that will give you less work to do in the future.
As the module mentions when you find a solution that can be used as a general solution for different problems try to add it to your toolkit so you don’t have to rewrite everytime the code that you already know works in some situations. Always look for general solutions instead of specific ones.

After refactoring the stock-analysis code we can notice a major difference in the time used to process the code. We moved from 3.5 to 0.5 seconds, as I was refactoring the code I notice there are several ways to perform certain actions and this also may have an impact on the efficiency of the code. For sure as we move through the bootcamp we’ll gain the knowledge needed to perform better and faster analysis.
