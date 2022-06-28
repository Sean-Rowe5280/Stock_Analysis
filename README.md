# Stock Analysis: Understanding VBA 
## Overview

### Background

Steve has been tasked by his parents to help them invest in a green energy stock. He obtained stock data for 12 green stocks for the years 2017 and 2018 to create an analysis for his parents. He wanted to looked at the yearly return and volume traded as the basis for his recommendations but the data sets were quite large and overwhelming. Calcualting those metircs for each stock for each year was a time consuming and a very manual process. He needed a better way to perform this analysis. This was an opportunity to leverage the power of Visual Basic for Applications or VBA, with excel.  VBA can retrieve the information needed from the data sets and format it in an easily read summary at the click of a button. VBA allows us to automate tedious time consuming processes and limit errors that can emerge from doing repetitive processes manaually. VBA is a programming language thats often used with Excel. It relys on for loops, conditional statments(if-then statments), logic operators and then the process of refractoring code to create a subsroutine which improves effieciency and readability.

### Purpose

We wrote a macro that enabled steve to perform an analysis of the 12 stocks hes considering and summarized the yearly return and total volume for each stock at the click of a button. It worked well for a dozen stocks but Steve may want to expand his analysis in the future to thousands of stocks and needs code that runs more efficiently. This is where refractoring or editing comes into play. *Refractoring* is the process of editing the logic of a code to increase how efficiently it runs and/or how easy it is to read and understand. It has a number of benefits and some drawbacks which I'll elaborate more on later. In this project we were given the task of refractoring our VBA code called "AllStocksAnalysis". This original code looped through the data multiple times which is not efficient so the challenge was to find a way to loop through the data one time which uses less memory and cuts run time. 


## Results

In our "AllStocksAnalysisRefractored" to improve the efficiency of the script we created a tickerindex variable which we set with a starting value 0.  We then created 3 output arrays for tickervolume(12), tickerstartingprice(12) and tickerendingprice(12) followed by a for loop 0 to 11 for these variables and set tickervolumes to zero.
![image](https://user-images.githubusercontent.com/107006216/176061968-31cc727e-5bb9-4f67-bfa7-4bd3c856333d.png)


 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107006216/176060950-f583fced-4358-45e0-b94f-1eab92dd1a1c.png)
![image](https://user-images.githubusercontent.com/107006216/176061051-60bc4109-3077-4748-9af5-b6f4f704a4da.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/107006216/176060958-25bdadc2-499f-4179-a7dc-0a412e7b6a84.png)
![image](https://user-images.githubusercontent.com/107006216/176061004-796c3d9e-d267-4e20-9fb6-15a3cb532d3c.png)
