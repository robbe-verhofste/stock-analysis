# Stock Analysis With Excel VBA

## Overview of Project
This project seeks to analyze stock data for 12 publicly traded companies in the renewable sector. We used data on each stock's daily volume, opening price and closing price for the year 2017 and 2018.
### Purpose
The purpose was to determine the best stocks to purchase, based on the trends from 2017 and 2018.
### The Data

## Results

#### Why Refactor?

##### Benefits
Refactoring is a technique to clean code and should be attempted as frequently as possible. This cleaning process reduces redundancy and clutter within code by combining scripts wherever possible. This reduces the potential for error, and reduces what developers call "code rot" [1](https://digital.ai/resources/agile-101/code-refactoring#:~:text=Code%20Refactoring%20is%20the%20process,this%20is%20hard%20to%20do.). Code rot occurs when new items of code have been written with similar functions as previous items of code - the coder changes one item but cannot identify a previous section. As a result, the code "rots", and begins to function abnormally.

###### When code should be refactored [Ershad, 2017](https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/):

* *Chances of Enhancement are high*
 * *If modules have chances to add new features or functionalities then make sure design and current code is good and following Open Close Principle*
* *Code Smell is Detected*
 * *Sometimes bad patterns like tight coupling, duplicate code, long methods, large classes, etc. are detected in the code;  the code should be refactored in this case.*
* *Bug Fixing*
 * *Codes are written badly in some cases, and so many bugs are raised. In this case, fixing of bugs take too much effort. So, the root cause of bugs can be code smell. So, before fixing bugs code should be refactored.*
* *Peer Review*
 * *Peer review is an important part of code refactoring. If the peer-reviewer finds some code smell then code should be refactored during peer review.*


##### Drawbacks
Code should not be refactored if deadlines are near (the process can be time intensive), if the costs of rewriting the code are less than the cost of refactoring, or if the code is already stable.

#### Code

      Sub AllStocksAnalysisRefactored()

      Dim startTime As Single

      Dim endTime  As Single

      yearValue = InputBox("What year would you like to run the analysis on?")

      startTime = Timer

      'Format the output sheet on All Stocks Analysis worksheet

      Worksheets("All Stocks Analysis").Activate

      Range("A1").Value = "All Stocks (" + yearValue + ")"

      'Create a header row

      Cells(3, 1).Value = "Ticker"

      Cells(3, 2).Value = "Total Daily Volume"

      Cells(3, 3).Value = "Return"

      'Initialize array of all tickers

      Dim tickers(12) As String

      tickers(0) = "AY"
      tickers(1) = "CSIQ"
      tickers(2) = "DQ"
      tickers(3) = "ENPH"
      tickers(4) = "FSLR"
      tickers(5) = "HASI"
      tickers(6) = "JKS"
      tickers(7) = "RUN"
      tickers(8) = "SEDG"
      tickers(9) = "SPWR"
      tickers(10) = "TERP"
      tickers(11) = "VSLR"

      'Activate data worksheet

      Worksheets(yearValue).Activate

      'Get the number of rows to loop over

      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      '1a) Create a ticker Index

      tickerIndex = 0

      '1b) Create three output arrays

      Dim tickerVolumes(12) As Long

      Dim tickerStartingPrices(12) As Single

      Dim tickerEndingPrices(12) As Single

      ''2a) Create a for loop to initialize the tickerVolumes to zero.
      ' If the next row’s ticker doesn’t match, increase the tickerIndex.

      For i = 0 To 11
          tickerVolumes(i) = 0
          tickerStartingPrices(i) = 0
          tickerEndingPrices(i) = 0
      Next i

      ''2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount

          '3a) Increase volume for current ticker

          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

          '3b) Check if the current row is the first row with the selected tickerIndex.
          'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
              tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          End If

          '3c) check if the current row is the last row with the selected ticker
          'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
           End If

              '3d Increase the tickerIndex.
               If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                  tickerIndex = tickerIndex + 1
              End If

      Next i

      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11

          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

      Next i

      'Formatting
      Worksheets("All Stocks Analysis").Activate
      Range("A3:C3").Font.FontStyle = "Bold"
      Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
      Range("B4:B15").NumberFormat = "#,##0"
      Range("C4:C15").NumberFormat = "0.0%"
      Columns("B").AutoFit

      dataRowStart = 4
      dataRowEnd = 15

      For i = dataRowStart To dataRowEnd

          If Cells(i, 3) > 0 Then

              Cells(i, 3).Interior.Color = vbGreen

          Else

              Cells(i, 3).Interior.Color = vbRed

          End If

      Next i

      endTime = Timer
      MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

      End Sub


### Analysis
In 2017, all stocks but TERP saw positive growth, while in 2017 the only stocks with positive growth were RUN and ENPH. I would recomend RUN or ENPH for current investment, but only after a thorough review of their current and projected outlook based on more recent data (2018-2020).

#### Refactored run-times
Below are the screenshots of the time it took for each analysis to run. These run times (0.1-0.2 seconds) were at least 5 times faster the unrefactored code, which ran for approximately 1 second.
![Time to Run 2017 VBA](https://github.com/robbe-verhofste/stock-analysis/VBA_Challenge_2017.PNG)
![Time to Run 2017 VBA](https://github.com/robbe-verhofste/stock-analysis/VBA_Challenge_2018.PNG)
