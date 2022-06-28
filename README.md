# Stock Analysis: Understanding VBA 
## Overview

### Background

Steve has been tasked by his parents to help them invest in a green energy stock. He obtained stock data for 12 green stocks for the years 2017 and 2018 to create an analysis for his parents. He wanted to looked at the yearly return and volume traded as the basis for his recommendations but the data sets were quite large and overwhelming. Calcualting those metircs for each stock for each year was a time consuming and a very manual process. He needed a better way to perform this analysis. This was an opportunity to leverage the power of Visual Basic for Applications or VBA, with excel.  VBA can retrieve the information needed from the data sets and format it in an easily read summary at the click of a button. VBA allows us to automate tedious time consuming processes and limit errors that can emerge from doing repetitive processes manaually. VBA is a programming language thats often used with Excel. It relys on for loops, conditional statments(if-then statments), logic operators and then the process of refractoring code to create a subsroutine which improves effieciency and readability.

### Purpose

We wrote a macro that enabled steve to perform an analysis of the 12 stocks hes considering and summarized the yearly return and total volume for each stock at the click of a button. It worked well for a dozen stocks but Steve may want to expand his analysis in the future to thousands of stocks and needs code that runs more efficiently. This is where refractoring or editing comes into play. *Refractoring* is the process of editing the logic of a code to increase how efficiently it runs and/or how easy it is to read and understand. It has a number of benefits and some drawbacks which I'll elaborate more on later. In this project we were given the task of refractoring our VBA code called "AllStocksAnalysis". This original code looped through the data multiple times which is not efficient so the challenge was to find a way to loop through the data one time which uses less memory and cuts run time. 


## Results

### VBA Code
In our "AllStocksAnalysisRefactored" to improve the efficiency of the script we created a tickerindex variable which we set with a starting value 0. This enables the code to reference the stocks in our ticker array, starting with the first ticker(0), and pull in the applicabale data for the our output arrays.   We then created our 3 output arrays for tickervolume(12), tickerstartingprice(12) and tickerendingprice(12) followed by a for loop 0 to 11 for these variables where we set the inititial value to 0.

![image](https://user-images.githubusercontent.com/107006216/176061968-31cc727e-5bb9-4f67-bfa7-4bd3c856333d.png)

Our next for loop if a for loop for all the rows in our data sets and we define tickervolumes using the tickerindex in order reference each individual stock ticker and calculate total volume by pulling if from the column define. Note that the tickerindex was set equal to 0 earlier so we'll need write some conditional formatting for the tickerindex to increase once its totaled all the relevant data for ticker(0).

![image](https://user-images.githubusercontent.com/107006216/176077581-674e6bb8-e79c-4e6f-a46c-93b4cb80750b.png)

Within this same for loop write conditional script to calculate tickerstartingprice and tickerendingprice again using the tickerindex to pull data for the appropriate stock.

![image](https://user-images.githubusercontent.com/107006216/176078569-d484287b-25c5-4915-98c7-7e15ef476725.png)

The last change within this for loop is conditional code to increase the tickerindex after its looped through all the rows of data for ticker(0).
 
![image](https://user-images.githubusercontent.com/107006216/176080027-31da7af6-680f-464b-9fbf-70aa411e9c7e.png)

These were the primary changes in our refractoring process. There are some other subtle differences btw the code "AllStocksAnalysis" and "AllStocksAnalysisRefactored"
but for the purposes of this challenge im not going to go into those.

### Measure of Performance

The results of our efforts can easily be seen if we add timer script to clock the time it takes the original code and refactored code.  The Original code took between .63 and .7 seconds. 

![image](https://user-images.githubusercontent.com/107006216/176082022-b359c673-5e08-4989-8166-81ae512972e0.png)
![image](https://user-images.githubusercontent.com/107006216/176082099-f71822de-530a-4d96-93ab-9aa71bff03e4.png)

The refactored code in comparision ran much quicker. .13 to .25 seconds. This may not seem like a significant enought improvement to justify editing our code but it could be if we analyize a much larger number of stocks and our data set increases significantly.

![image](https://user-images.githubusercontent.com/107006216/176082374-4b59bb5f-11e9-452e-a67a-f473b3483545.png)
![image](https://user-images.githubusercontent.com/107006216/176082645-2ffc2324-c6ad-4dda-865b-454f177abed4.png)

## Summary

### Original Code: Advantages and Disadvantages

The primary advantage I found with the old code was its simplicity, it was more intuitive and very easy to follow and it accomplished what it was designed to in a reasonable time frame.  Although the code was much slower, which might be a significant disadvantage with a larger data set, in this instance the effieciency difference was negligible.

### Refactored Code: Advantages and Disadvantages

There are in fact both pros and cons to refactoring code and they need to be considered and wieghed against each other.  The main advantage in this example is the speed in which the code was executed. By removing nested loops from our script we reuduced the amount of processing required(more efficient) and the code completed 3-5 times faster. This efficiency improvement also makes the code more useful because we could use it on a much larger data set.  The main downside to our refactored code in this example efficiency gain really isn't worth the amoutn time spent reworking the code.  Granted in the future a significantly larger data set it might be worth it but in this case it wasn't. Also, and this perhaps is more related to my limited VBA experience, the refactored code was much more challenging to follow and work out.  The readability did not improve.















