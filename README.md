# Stock-Analysis

## Overview of Project:
In this project, we edited and refactored a Microsoft Excel VBA code to analyze the performances of 12 sets of stocks for both year 2017 and 2018. The goal is to loop through all the data at one time and achieve the results in a more efficient way. The refactored code uses less memory and make it easier for future users to read.

##Results
Below is the orginal VBA code. It contains 2 loops.

    Sub AllStocksAnalysis()

    Dim startTime As Single
    
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer


    Worksheets("All Stocks Analysis").Activate
    
        Range("A1").Value = "All Stocks(2018)"
    
        'Create a header row
    
        Cells(3, 1).Value = "Ticker"
    
        Cells(3, 2).Value = "Total Daily Volume"
    
        Cells(3, 3).Value = "Return"

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
     
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    Worksheets("2018").Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
    Worksheets("2018").Activate
        For j = 2 To RowCount
        
    If Cells(j, 1).Value = ticker Then
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
       End If
       
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        startingPrice = Cells(j, 6).Value
        
       End If
       
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
    
        endingPrice = Cells(j, 6).Value
        
       End If
       
    Next j
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1    
    Next i
  
  After we run the code, we can see that it takes almost 1 second for year 2017 and 2018's result
  



