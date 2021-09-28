# Stock Analysis

## Overview of Project

### Purpose
Refactor a stock analysis code to deliver the analysis faster. 

### Backgorund
Steve loved the workbook for stock analysis. He wants to expand the dataset to include the etire stock market over the last few years. Although the macro works for a few stocks, it may take a long time to execute for a thousand stocks. 

## Results

### Analysis
To compare the efficiency of the refactored code and the original code, we are going to use the total run time after the user input. For a fair compare, we are going to use the average running time of the same data.

The original code have the following running times:

* For 2017, it took 0.914s to analyse 12 stocks.

![image](https://user-images.githubusercontent.com/21972342/135027781-0e5d0967-7d31-4c65-98d5-732232364bd4.png)

* For 2018, it took 0.856s to analyse 12 stocks. 

![image](https://user-images.githubusercontent.com/21972342/135027846-722e6ac7-3cc3-4168-bc47-4c527dbc4825.png)

Meaning, the original code took an average of 0.885s to analyse all 12 stocks. 

The recatfored code have the following runnin times:

* For 2017, it took 0.104s to analyse 12 stocks.

![image](https://user-images.githubusercontent.com/21972342/135028730-e72f1686-be48-4dbb-907d-d834a5b41992.png)

* For 2018, it took 0.108s to analyse 12 stocks. 

![image](https://user-images.githubusercontent.com/21972342/135028764-81df1171-3e64-4b9f-89b7-ea744582a250.png)

Meaning, the refactored code took an average of 0.106s to analyse 12 stocks. 

In total, the refactored code is 734.9% more efficient (12 stocks/0.106s)/(12 stocks/0.885s).

The difference between original code and the refactor code is we loop over all rows. In the original code we have that the for loop do the compare over all the rows 12 times. Instead, the refactor code did the compare a single time.  

* The original code is as:
```
    For i = 0 To 11
   
        ticker = tickers(i)
        startingPrice = 0
        endingPrice = 0
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        'loop over all the rows
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then

                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value

            End If
        
            'Search if the previews cell is not current ticker and current cell is current ticker, store the cell as the starting price
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            'Search if the current cell is current ticker and the next cell is not current ticker, store the cell as the ending price
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If
            
        Next j
        
        Worksheets("All Stock Analysis").Activate
        Cells(i + 4, 1).Value = ticker
        Cells(i + 4, 2).Value = totalVolume
        Cells(i + 4, 3).Value = (endingPrice / startingPrice) - 1

    Next i
    
```

* The refactored code is:
```
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        
        'If the next rows ticker doesnt match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
```

## Summary

### Advantages and disadvantages of refactoring code
The advantage of refactoring code is create a more efficient and understandable code. But, for this is necessary a creative mind, to find alternative solutions, and coding experience, to implement the proposed solutions. 

The main disadvantage is that, it can become time-consuming (Stone, 2018)[<sup>1</sup>](#1).When you are on a deadline, get the things done is more important. Other situations that you may consider not refactoring is if you are sure you are not going to reuse the code.

### Advantages and disadvantages if the original and refactored VBA script
The main advnatage of the refactored code is its efficiency >700%. If we were to include a thousand stocks, the original code would take 71s to complete the task (>1 min). Instead, the refactored code would take only 9 seconds to complete the task. 

The advantage of the original code is, that doesn't matter how the tickers are sorted, as long that they are grouped by name, it will do the calculation. In the refactored code, if the sort of the tickers change, the macro will fail.

Still, the code can be improved in a next revision with at least following two things:
* First, to add a sorting by ticker name function. Because, both, the original and refactored code, fail if the tickers are not grouped by name. 
* Second, do the tickers array automatically. Curently, for both codes, in the case that we want to add tickers, we need to initialize the arrays for each ticker manually. In the case that there are a 1000 tickers or that in each year there are different tickers the macro will not be useful anymore.  

## Footnotes

<a class="anchor" id="1"></a>[1]: Stone, Sydney. "Code Refactoring Best Practices: When (and When Not) to Do It"._altexsoft_, 27 Sep 2018, https://www.altexsoft.com/blog/engineering/code-refactoring-best-practices-when-and-when-not-to-do-it/.




