# VBA_Challenge
 ## Overview
The purpose of this analysis is to examine a vba script written to iterate over a spreadsheet containing stock prices over 2017 and 2018, and produce easily referencable metrics including volume traded and return from the start to the end of the year for each stock contained in the spreadsheet.
I regret nothing about my first submission of this project, and if I wasn't so worried about how my capstone is going to tank in a big way I wouldn't be resubmitting it sheerly out of spite for VBA.
 ## Results

 **2017 original code results**  
 ![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/OldCode2017.png)  
 **2018 original code results**  
 ![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/OldCode2018.png)   
  
**Refactored code results 2017**  
![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/NewCode2017.png)  
**Refactored code results 2018**  
![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/NewCode2018.png)  
I was able to get a significant improvement in speed out of the refactoried code- down from just under a second in the original code to less than a tenth of a second in the refactored code. 2018 runs significantly faster in both cases, but I don't know why, since it's running the same macro. The ticker volumes in 2018 seem slightly smaller overall, which could be the reason.
The results the refactored code outputs is correct with the module's code after a minor modification- the module (at least in 2.3.3) uses Cells(j,6).Value for both starting and ending prices, which gave me a headache for about half an hour before I noticed it.  The un-refactored macro I ran uses startingPrice = Cells(j, 3).Value for starting price and endingPrice = Cells(j, 6).Value for ending price instead of using startingPrice = Cells(j, 6).Value for both.  
The extra columns Open and Close in the refactored screenshots are there to check my work on starting and ending price values (refactored as 
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
           tickerStartingPrices(tickerIndex) = tickerStartingPrices(tickerIndex) + Cells(i, 3).Value
            
        End If
```  
and  
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = tickerEndingPrices(tickerIndex) + Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
```
respectively), and they do line up with what's in the spreadsheet, so I'm confident my results (while they don't match the screenshots in the module assignment) are correct.  
  
In terms of the actual stock performance, we can see that 2018 was overall a bad year for investments unless you were investing only in ENPH or RUN, in which case you were very happy.
DQ had an exceptionally bad year in 2018 compared to 2017, going from a 200% return down to a -60% return. Based on these numbers in a vaccuum, I would advise someone to invest in ENPH, which had good return in a mostly green year and proves resilient in a mostly red year as well. This seems to be a popular opinion- ENPH volumes went up by 400,000,000 in 2018.
## Summary
My conclusion on this project remains mostly the same: refactoring is excellent for clunky blocks of code with long runtimes. It can save time and computational resources if it's done well. Additionally, it can be a good way to learn more about a programming language (I certainly learned more about vba), if there's nothing resting on the refactoring. It's inadvisable to do for anything important to a project or system if you don't know exactly what's going on, since that's how spaghetti code breaks your whole program. And, again, on a personal note, I don't think it's worth refactoring something that takes a few seconds to run if the refactoring itself takes several hours, which this (again) did.  
  
The only advantage that I see in the original script over the refactored script is that it's a little easier to understand. I expect that, in a practical sense, the audience who wants to use an excel macro rather than perform this analysis in a useful language like python, is the excel power user, and as a former excel power user, we are not necessarily great at deciphering code.
Otherwise the refactored macro outputs the exact same results in a fraction of the time. The disadvantages of both macros, as they're written currently, is that neither can accomodate any additional stocks in the spreadsheet. The tickers array would have to be remade before the macro could run, or it would either not output the added stock or throw a 'subscript out of range' error. I'm not going to test to find out. The disadvantage of the original macro versus the refactored one is that the refactored macro is appreciably faster, even on a fairly fast computer. In a world where we weren't working with 12 stocks, but rather several hundred, I would definitely want to have put in the time to refactor this macro to save me time in the future.
