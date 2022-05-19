# Stock-Analysis

## Overview of Project:
In this project, we edited and refactored a Microsoft Excel VBA code to analyze the performances of 12 sets of stocks for both year 2017 and 2018. The goal is to loop through all the data at one time and achieve the results in a more efficient way. The refactored code uses less memory and make it easier for future users to read.

## Results

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
  
  ![2017Originalcode](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/script%20run%20time%20for%20year%202017_original%20code.PNG)
  ![2018Originalcode](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/script%20run%20time%20for%20year%202018_orginal%20code.PNG)

  Below is the refactored VBA code 
  
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
    
     For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) Then
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
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

And below are the updated results after running the edited code

![2017newcode](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/script%20run%20time%20for%20year%202017.PNG)
![2018newcode](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/script%20run%20time%20for%20year%202018.PNG)

We can tell that now it only uses one fourth of the original length, which highly improves the efficiency.

As for the stock performance between 2017 and 2018, 

![2017stock](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/results%20year%202017.PNG)

![2018stock](https://github.com/hihilynette/Stock-Analysis/blob/main/Resources/results%20year%202018.PNG)

We can tell that 2017's stocks outperformed 2018, with only 1 stock "TERP" has a negative return.

## Summary

##### What are the advantages or disadvantages of refactoring code?
Pros: 
1. It saves more time and uses less memory.

2. It is easier for other users to read and further edit.
      
3. It can be applied to a larger database.
      
Cons: 

1. The original code must be fully understand before editing it. 
      
2. It is possible to create more bugs which cost time to fix.

##### How do these pros and cons apply to refactoring the original VBA script?

The pros of the refactored code definetly decreases the original macro running time which is more efficient. Also the edited code can be more functional by applying to a larger database. However, the new code must be readable and fully understanded before refactoring it. Otherwise it will have more errors coming up.

      


    

