# Analysis of All Stocks

## Overview of the Project

The purpose of this project was to analyze multiple stocks over 2017-2018. By doing this, I was able to help Steve present his parents with multiple options for investing.

## Results

### Code Explanation with Examples

In this analysis, I created a for loop to loop over all of the stocks in one of the worksheets. The worksheet that I looped over was dependent on the year I input when the code was ran. I refactored the original code to make it more efficient. I did this by creating an index for tickers. In this index, I created 3 arrays, including tickerVolume, tickerStartingPrice, and tickerEndingPrice. 

Included below is an example of the code I wrote to create the index and arrays

    Dim tickerIndex As Integer
      tickerIndex = 0

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single


I created a for loop which told the computer to run the code from row 2 to RowCount, which was previously defined in the code. Below is the code I used to start the for loop.

   `For i = 2 To RowCount`


I increased the volume for the current ticker using the code below. 
			
`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`
		
To tell the code where to start and where to stop, I added an "if, then" statement. I have included the code for where to start and stop the analysis below, respectively. 

`If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then            
     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If`


`If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
End If
`

I created another "if, then" conditional to increase the tickerIndex and closed the for loop. The code is included below.

`If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then            
                tickerIndex = tickerIndex + 1
End If
Next i`

Finally, I created another for loop to output the information onto the “All Stocks Analysis” Worksheet. This is the code I used to do so.

    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    
 ### Screenshots of Run Time
 
In the original code (Module 1 (Code)), I did not have a ticker index. The original code ran much slower than the later code. Included below are screenshots of the run time for each year using the original code.

![original_2017](path/to/original_2017.png)

![original_2018](path/to/original_2018.png)


In the refactored code (Module 3 (Code)), I included the ticker index with 3 output arrays. This improved the speed at which our code ran the analysis. Included below are screenshots of the run time for each year using the refactored code.

![refactor_2017](path/to/refactor_2017.png)

![refactor_2018](path/to/refactor_2018.png)


## Summary

### General Advantages and Disadvantages of Refactoring

In general, there are pros and cons to refactoring code. One of the advantages to refactoring code is that it can run analyses faster and with more efficiency. It can do this through editing out unnecessary steps and instead combining multiple steps into less lines of code. One of the disadvantages to refactoring code might be figuring out how to create a more efficient code while still creating the same outputs.  

### This Challenge's Advantages and Disadvantages of Refactoring

As previously shown through screenshots, the refactored code was able to increase the speed at which the analysis ran. This is a clear advantage of refactoring the VBA code in this assignment to make it more efficient. One of the disadvantages that I faced was that I ran into some issues figuring out how to create the index for the various arrays and then use them correctly throughout the code. It was difficult to find online resources that could help explain how to do this, so I opted for a tutoring session instead. The tutor helped me learn how to create this index and understand how it operates. Aside from my personal troubles with this particular assignment, I think that refactoring code has clear advantages over disadvantages. 




