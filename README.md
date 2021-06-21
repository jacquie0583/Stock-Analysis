#  A View of the Stock Market the Past Year
## Overview of Project
###  Purpose
Grateful, Steve, that you have appreciated our work so far and would like us to expand the dataset to include the entire stock market over the last few years.  We will check to see if the information is the same as previously delivered then we will refactor so the VBA script runs in a timely fashion by adjusting the code to loop through all the data one time in order to collect the same information in the Data Analysis Section.  Other possible techniques to accomplished this is by taking fewer steps, using less memeory, or improving logic code to a straight forward approach.
###  Background
To recap, this project is a result of your parent’s desire to invest in alternative energy sources which lead to their investment in DAQO New Energy Corp.  You promised them an assessment of the company’s performance as well as you have concerns about diversity.  Your request of a new program will help provide you with data to assist you in possible choices so you can diversify.
##  Analysis
### Refactor VBA Code Process Explainedw with Examples
####  TickerIndex was created as a variable and set to zero before iterating over all the rows.  It will be used to access the appropriate index across the four different arrays.
      For i = 0 To 11
        tickerIndex = tickers(i)
####  Output Arrays were designated as Long data type (ticker Volumes) and Single data type (ticker starting Prices and ticker ending prices).
      Dim tickerVolumes As Long
      Dim tickerStartingPrices As Single
      Dim tickerEndingPrices As Single
####  A loop initiates the tickerVolume to Zero and accessed by the tickerIndex.  Providing a filter if the next row's ticker doesn't match, the tickerIdeex is increased.
       tickerVolumes = 0
     For j = 2 To RowCount
        If Cells(j, 1).Value = tickerIndex Then
####  The code loops through the stock data, reading and storing the data from all the rows.  Modifying withing the loop, allows for the increase of the tickerVolume variable thereby increasing the tickerVolume for the stock ticker. We utilized if-then statements as a filter if to assign the closing price to the ticker starting prices and ticker ending prices variable.
      If Cells(j, 1).Value = tickerIndex Then
        tickerVolumes = tickerVolumes + Cells(j, 8).Value
         If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
              tickerStartingPrices = Cells(j, 6).Value
         End If
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
              tickerEndingPrices = Cells(j, 6).Value
####  Formatting the cells with codes added color and self-explanatory visual appeal.
         If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
         Else
             Cells(i, 3).Interior.Color = vbRed
####  Comments summited according to Best Practices for Writing Readable Code
##  Results
The All Stock Analysis dataset was consistent with that of the Refactored Analysis dataset.  However the runtime varies in successful outcomes.
####The initial 2017 code ran .08125 seconds.

the refactored 2017 code ran

