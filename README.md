
# Stock Analysis With Excel VBA
Click here to view the Excel file: VBA Challenge - Stock Analysis

# Overview : VBA Stock Analysis Project

## Purpose

In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

### Analysis and Challenges

Prepare our dataser VBA_Challenge.vbs file for the project.

Create our resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.

Create and convert our XLSM file from *.vbs dataset that you used in this module as VBA_Challenge.xlsm.

Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.

Use the steps Refactor VBA code and measure performance to add code where indicated by the numbered comments in the starter code file.


### Results:Refactor VBA Code and Measure Performance

Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1. The tickerIndex is set equal to zero before looping over the rows.

Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

![Screenshot 2022-07-09 175643](https://user-images.githubusercontent.com/108489186/178125260-cc921da9-9d77-465b-a3d4-4498fff3f567.png)

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

![Screenshot 2022-07-09 175835](https://user-images.githubusercontent.com/108489186/178125273-d36901b2-4bd0-497e-9ce5-853327171709.png)

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

![Screenshot 2022-07-09 180035](https://user-images.githubusercontent.com/108489186/178125286-b2669390-cd5f-4c54-b271-cb722c9ae1e0.png)

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

![Screenshot 2022-07-09 180930](https://user-images.githubusercontent.com/108489186/178125300-bc11c85e-20d4-4921-9cac-aff5bf7dce71.png)

5. Code for formatting the cells in the spreadsheet is working.

We make positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.

![Screenshot 2022-07-09 181208](https://user-images.githubusercontent.com/108489186/178125307-cfbc6303-f07e-4a3f-b3e7-6f25e6e21ebc.png)


##Analysis

Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.

![Screenshot 2022-07-09 181645](https://user-images.githubusercontent.com/108489186/178125328-4d975e34-8069-4c81-a9bc-2433bd33df4f.png)

 ###The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named Dataset Examples Provided). In adition, in our resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook.

Final VBA Analysis 2017

![1VBA_Challenge_2017](https://user-images.githubusercontent.com/108489186/178125358-fe6ca921-bc7c-4f47-9ce8-7f9215998938.png)

Final VBA Analysis 2018

![1VBA _Challenge_2018](https://user-images.githubusercontent.com/108489186/178125370-76462a0f-9011-4f6c-a156-7f70c1d114ff.png)

## SUMMARY

###Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

###Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A complex unstructured code is usually best to split in several functions.
Refactoring process can affect the testing outcomes.

###Advantages:

Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.



