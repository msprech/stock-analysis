# Analyzing Green Stocks Through VBA

## Purpose of Analysis 

In order to provide insight into possible prosperous green energy stocks for the client, I analyzed a variety of stocks using the metrics of total yearly volume and return on stocks. The client's prior interest was in the DAQO New Energy Corp, which was shown through initial analysis to have dropped by 63% in 2018. We then needed to widen the scope of our analysis to include both 2017 and 2018 as well as a much wider dataset of stocks. 

After automating initial analysis through VBA for the year of 2018, I then needed to refactor the code to work more efficiently for both years of stock data. Refactoring the code to collect all of the information on a single loop decreased the run time, improved readability, and allowed further analysis into client's desire to diversify his funds by investing in green energy stocks. 

## Results 

The refactored code produced a more flexible macro with a decreased run time and an expanded view of how well each stock performed in 2017 and 2018. All results were captured within one loop through the use of four different arrays. After declaring the four arrays, which were tickers(12), tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12), I activated the data worksheets, set the tickerIndex equal to 0, and ran the following snippet of code. Through loops and conditionals, I calculated the total daily volume of each ticker index, as well as the starting and ending prices used to calculate the return.  

```
For i = 0 to 11 
  tickerVolumes(i) = 0 
Next i 
 
For j = 2 to RowCount 
   tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value 
    
   If Cells(j, 1).Value = tickers(tickerIndex) And Cells(J - 1, 1).Value <> tickers(tickerIndex) Then 
      tickerStartingPrices(tickerIndex) = Cells(j, 6).Value 
      
   End If
   
   If Cells(j, 1).Value = tickers(tickerIndex) and Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
      tickerEndingPrice(tickerIndex) = Cells(j, 6).Value 
      tickerIndex = tickerIndex + 1
   
   End If
   
 Next j
 ```
 I then set up a for loop to direct output of these values to the worksheet "All Stocks Analysis," and ended by formatting the worksheet for better readability. I also included the following code to color code the cells based on the output. 
 
 ```
For i = dataRowStart to dataRowEnd 
 
  If Cells(i, 3) > 0 Then 
    Cells(i, 3).Interior.Color = vbGreen 
  Else 
    Cells(i, 3).Interior.Color = vbRed 
  End if 

Next i 
```
The resulting table for 2017 stock data showed that index TERP was the least successful with a -7.2% return, while all other stocks had a positive return. However, 2018 had only two stocks, ENPH and RUN, reporting positive returns, with ENPH also having the highest total daily volume. 

![2017 Green Stock Data](https://github.com/msprech/stock-analysis/blob/92915fd4e6a5566394bb276de70ae3fbbccbc8ec/Resources/Screen%20Shot%202021-09-19%20at%209.07.38%20AM.png)

![2018 Green Stock Data](https://github.com/msprech/stock-analysis/blob/92915fd4e6a5566394bb276de70ae3fbbccbc8ec/Resources/Screen%20Shot%202021-09-19%20at%209.07.54%20AM.png)

## Summary 

### Advantages and Disadvantages of Refactoring Code 

Refactoring can both simplify and clean up complicated and difficult to read code. It is a process that provides further opportunity to debug and catch errors in existing code. It can also improve run time and illuminate useful patterns in the macros. 

However, it can also be time-consuming for large data sets, and increases the risk of introducing more errors or breaking code if you aren't keeping close track of what each section of code does. Although refactored code can address many issues that come with coding and analyzing data, it is also very important to use comments and replace as many hardcoded values as possible in order to stay organized. 

### Refactoring the VBA Stock Analysis Script 

The run time for the green stock analysis script was vastly improved by refactoring the code. The initial code ran in close to .3 seconds, while the refactored code was closer to .08 seconds as shown below.  

![2017 run time](https://github.com/msprech/stock-analysis/blob/92915fd4e6a5566394bb276de70ae3fbbccbc8ec/Resources/VBA_Challenge_2017.png)
![2018 run time](https://github.com/msprech/stock-analysis/blob/92915fd4e6a5566394bb276de70ae3fbbccbc8ec/Resources/VBA_Challenge_2018.png)

By adding an InputBox() value that asked for a specific year of desired analysis, the refactored code also allowed further flexibility and applied to a much wider range of stocks, as opposed to only the year of 2018. There was also a reduced risk of accidents and errors in looping through the data a single time and keeping the calculations concise. 

However, with an increase amount of data, I did run into the overflow error on VBA, and it was also more difficult to keep track of all of the different loops and conditionals. Including comments was necessary to ensure in particular that all declared variables were where they would work best.  
