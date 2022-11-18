
# Green Stock Analysis

## Overview of Project: Stock Analysis using VBA

###Purpose
The purpose of this analysis was to help my friend Steve and his parents make financially shrewd decisions investing in
stocks for environmentally sustainable companies. To achieve this, he provided me with a data set that included two years
worth of trading data for 12 companies. This data was then filtered using **Visual Basic Macros**, to create an a more
detailed understanding of the trading volume and whether the stock had a positive or negative return in a certain 
calendar year. **The goal was to provide details on the return on investment in a more clear and faster way.**

## Results
With there being two stated goals for this analysis, the **Results** section will be broken down in to two distinct
sections: **Macro Performance** and **Stock Performance**

### Macro Performance
The Macro provided to me by Steve was initially incomplete, needing to complete various **arrays** as well as complete
several **nested For loops** to get the outputs need to make the decisions that Steve and his parents have requested.
To achieve this the first thing that needed to be done was to create an array of all the tickers present within the
dataset, this was achieved through this section of code:

`Dim tickers(12) As String
    
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
    tickers(11) = "VSLR"`

This allowed us to more efficiently automate the macro by assigning each ticker a numerical value, which requires a
smaller piece of bite of information than a string of letters. After a few minor steps involving the layout of my
output, the next major piece of code was the **For loop** that initialized our primary variable:

`For i = 0 To 11
        tickerVolumes(i) = 0`

What this code does is to tell the Macro that we start our count at Ticker 0 which is AY, the macro then reads every
line of the data set for a certain year until it gets to the next ticker, Ticker 1. Now without the next section of 
code the Macro would have simply stopped and given us an output for AY and not proceeded any further. So, to prompt
the macro to continue running I need to make the first of what would ultimately be a three tiered **For loop**, in
order for the Macro to continue it was instructed:

` For i = 2 To RowCount`

What this code does is instruct the Macro to do is run from line 2 all the way to the bottom of the dataset. This 
gives us the fullest picture of the data possible. 

The next section of code gives the Macro instructions to glean the return data for every instance of a given ticker
starting at ticker 0 until it comes across the next variable, when the Macro arrives at the next ticker it takes 
the data from the previous ticker and puts it on our stated landing page, and then increases the count of the 
**tickerIndex** allowing the Macro to move forward with the next ticker and ending our first loop. This is not however 
the only time that the Macro would run through this code, those instructions are shown here:

`If Cells(i, 1).Value = tickers(tickerIndex) Then
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            End If
If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
             tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If`
This **End If** tells our Macro that once it has computed all the information for a ticker to then move on to the next
ticker, where it repeats the process with the next ticker.  The next section of code tells the Macro how to display the
output so that Steve and his parents can read it easily. The code that does is broken up in to three sections that follow:

Section 1
' For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i`     
   
The first section tells the Macro what each column on the output sheet will be labeled as and how to determine what data
goes in were. With the **tickers** going in the first, the **tickerVolumes** in the second and the **Return** going in the
third, until now there had been no direction for the Macro to determine the **Return** but that is achieved by a simple
equation, dividing the ending prices by the starting prices.

The next section of the code provides more stylistic instructions for the Macro, seen below:

    `Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15`

The first three instructions tell the Macro where to make the table and how to make column titles in bold with a solid line
running underneath across the whole of the table. The second section of instructions tells the Macro how to format the contents
and automatically fit the cells to the data.

The final section of the data tells the Macro how to make this table most useful to use visually. With positive returns shaded 
green and negative returns shaded red. 

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If`
The results of this **refactored** Macro are quite successful with the Macros performing at faster speed than desired. Those
measures can be found at the links below:

2017
(https://github.com/dh4rt/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

The **refactored** Macro performed 1/100 of a second faster than the original in its analysis of 2017

2018
(https://github.com/dh4rt/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

The **refactored** Macro performed 2/100 of a second faster than the original in its analysis of 2018

### Stock Performance

Upon completion of the analysis of both years there are two stocks that have shown to have outstanding returns on investment.
ENPH and RUN, both of these stocks have averaged jaw dropping numbers with ENPH achieving and average return of 105.7% and RUN
getting 44.75%. These two companies clearly have something truly remarkable going on and I would suggest that Steve and his
parents consult with a licensed financial advisor before fully committing to an investment.

## Summary
This challenge definitely lived up to its name with the refactoring of this code nearly taking 10 hours to complete. The biggest
stumbling blocks consistently proved to compiling and run-time errors. A large reason for the difficultly here was the lack of a
meaningful debug system with the VBA editor. But with all that being said, **Refactoring code is significantly more efficient that writing it from scratch**, in much the same way that editing is less usually less time consuming that writing original code. Having a general idea of what the basic structure of the code is to do allows for an easier time solving the problems at hand.

The advantages of refactoring VBA script follow much the same path, instead of having the outline, write, debug, and troubleshoot your own code from the ground up you can work off the skeleton, or even more, of code you are either given by the person with a problem to solve or the company who hired you in the first place. With the wealth of information available on StackOverflow the solutions to most problems is at the end of an effective web search through your engine of choice.

