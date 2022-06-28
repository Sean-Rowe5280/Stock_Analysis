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

These were the primary changes in our refractoring process. THere are some other subtle differences btw the code "AllStocksAnalysis" and "AllStocksAnalysisRefactored"
but for the purposes of this challenge im not going to go into those.

### Measure of Performance

The results of our efforts can easily be see if you add timer script to clock the time it takes the original code and refactored code to run.  Our Original code took between .63 and .7 seconds. 
