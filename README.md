# Stock-Analysis with VBA 
## Overview 
The following workbook is prepared for Steve to help him in advice his parents concerning the stock market with carefully done analysis and with clear and precise visualization. However, to do this, a previously used code comes in handy since the data being handled at this stage is quite huge and a code that saves time is crucial; and the fact that it saves some memory. This previously used code must be adjusted accordingly to perform its new function without any major issues. In this case, the code is being reused and attention to detail is very important to avoid misinterpretation of data.  This is called refactoring and is what will be done in this analysis. 

## Analysis
As it is going to be illustrated below, it is safe to say the stock market can be a very volatile place as seen by major changes that occurred in between 2017 and 2018.  A drastic drop in yearly returns between the two years can be very discouraging to prospective investors. 

The first thing to do when refactoring this code is to set the tickerIndex to have a value of Zero which is the first step before you go over the following rows. This is displayed as shown below:

![SD1](https://user-images.githubusercontent.com/101376325/161449461-46b29444-a3d4-4505-b8a8-a6508e17d623.png)


Three output arrays for easy visualization are created to help analyze data in the simplest way but informative way possible.  The arrays are tickerVolumes, tickerStartingPrices and tickerEndingPrices which are Long, Single and Single data types as displayed by the following image:

![SD2](https://user-images.githubusercontent.com/101376325/161449498-c5961585-2d6f-4244-a58b-992b32f9ba07.png)


A loop is created to initialize tickerVolumes to zero and if it does not match, the tickerIndex is increased which is displayed below:


![SD3](https://user-images.githubusercontent.com/101376325/161449540-289394bd-c6c1-4ba9-951a-18b75bf0e65d.png)





A script containing the existing tickerVolumes variable has a loop that goes over all rows in the spreadsheet which in turn adds the ticker volume for the current ticker


When the if-then statement is executed, it checks if the existing row is the first one with the selected tickerIndex, if that is the case, the current price is assigned as tickerStartingPrices and tickerEndingPrices variable shown by the image below:


![SD4](https://user-images.githubusercontent.com/101376325/161449577-3faaad10-823b-4319-9626-dbea7084f6ee.png)




To inform a prospective investor who is not technically savvy about general trends of a stock market, use of colors can be a good way of telling them if they really should invest in a particular stock and when can be the best time to do so. It is also important to clearly show numbers which is the most important thing when financial analysis is concerned hence the AutoFit function comes in handy. The bold is used to distinguish the header from the data, rounding off is important to avoid wasting time reading decimal points that are not so significant and use of commas to clearly read the values of numbers since Wall Street investors invest in the hundreds of millions and hence easy readability is recommended. 2017 was generally a good year in returns as compared to 2018. This is illustrated below:

![SD5](https://user-images.githubusercontent.com/101376325/161449637-06754853-ef87-4948-a08c-0192ecd2372f.png)





This refactored code was able to use a shorter time as compared to the original code. This makes sense since the original code had to run the Analysis code first then the static and conditional formatting code which consumes memory and time.  Merging more than a code can be a real lifesaver especially when dealing with huge amounts of data.  This is displayed in detail with the following images:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/101376325/161449372-511b4bce-5d29-4f76-8885-0cfbe4c43ba1.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/101376325/161449405-fed67bd7-9dac-4d1d-92bf-5e90954659a9.png)


The original code had the run time as shown:

![Code Perfomance for 2017](https://user-images.githubusercontent.com/101376325/161449785-862e64ad-fea1-4307-a9eb-93fb33916620.png)


![Code Perfomance for 2018](https://user-images.githubusercontent.com/101376325/161449829-f984c3e0-afe6-4519-8e31-c6ce53649b50.png)

## Summary
### Advantages of refactored code
1.	It saves memory and time- In this case, having one macro instead of two or more saves a considerable amount of time and a coder can execute multiple processes at the same time. 
2.	 The complexity of trying to restructure a code reveals patterns that a single macro could not display

### Disadvantages of refactored code
1.	Despite saving time, it takes time to create one. Any single missed detail can prove fatal in the sense that data can be displayed incorrectly and can mislead investors.
2.	Having a slow computer can be a real blow since processing time of different machines can be different hence you can not tell if a certain code is really time efficient as compared to another. 

### Advantages and disadvantages of original and refactored VBA code
After completion of stock analysis, it’s a refactored code clear that advantages of a refactored VBA script outweigh its disadvantages. One big example is a refactored code saves a coder the hustle to have to create multiple macros. As a result, a client or a prospective investor only receives one set of data that covers everything.  Let’s have this in bit more perspective in relation to the data that has just been analyzed, creating two macros, one showing individual years without formatting to show trends and another showing the same information with the relevant trends OR one macro serving the functionality of two? It is obvious that handling two tasks at the same time is time efficient for both the client and the developer. 
