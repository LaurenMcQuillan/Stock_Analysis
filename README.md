# Green Energy Stock Analysis

## Overview of Project

### The purpose of this analysis is to help Steve determine which green energy stocks would be best for his parents to invest in. His parents want to invest in DAQO (ticker DQ), but Steve thinks they should diversify and invest in more than one green energy company. By obtaining an Excel spreadsheet with data for multiple stocks, and by using VBA, I will be able to quickly determine the Total Daily Volume and the Return for multiple green energy stocks in the years 2017 and 2018 to help Steve and his parents make a decision. Since I will be using VBA, the analysis will be automated, so it’s possible to reuse the code in the future if more stock data becomes available. The use of VBA also reduces the margin of error when performing this type of analysis.

## Results

### The green stock analysis results contain three columns of data: the Ticker name, the Total Daily Volume, and the Return %. In this analysis, I compared the stock performance in 2017 vs. 2018 for 12 different green energy stocks.

### Total Daily Volume

#### The volume of a stock is the total quantity of shares traded within the given period of time. Across all 12 stocks, the total daily volume traded was about 140,000 shares higher in 2018. When looking at each stock individually, there were 7 stocks that had an increase in total daily volume from 2017 to 2018, and 5 stocks that had a decrease in total daily volume from 2017 to 2018.

### Yearly Return

#### The return of a stock shows the increase or decrease in the price of a stock over a given period of time. In 2017, 11 of the 12 green energy stocks had a positive return. In 2018, only 2 of the 12 green energy stocks had a positive return.

### Overall Results

#### In 2017, DQ had a relatively low total daily volume but a 199.45% Return. In 2018, DQ had a higher total daily volume but a -62.6% return. If this trend continued, I would not recommend investing in DQ.  ENPH and RUN were the stocks that had positive returns on 2018. It would be helpful to obtain additional data on these stocks to try to determine a pattern and to see how these green energy companies have been performing in recent years.

### VBA Code

#### The VBA code from the original script and the refactored script both have the same output, but the refactored script runs faster than the original script. This was made possible in the refactored script by creating a tickerIndex variable, creating three output arrays, initializing the output arrays to 0, looping through all rows to analyze the stock data and find the results, and lastly outputting the results. The original script contains a nested loop that finds all results for one stock at a time and then outputs the results.

#### Original Script and Run Time
![Original Script](https://user-images.githubusercontent.com/115508658/199611271-cb68ba1d-c114-421a-b209-47e9f9995cf3.png)

![AllStocksAnalysis_2017](https://user-images.githubusercontent.com/115508658/199611330-477f7da5-43d6-4892-9fa8-b450bd76f4a1.png)

![AllStocksAnalysis_2018](https://user-images.githubusercontent.com/115508658/199611342-138cc4aa-e66f-4389-bd1f-d89552a6991a.png)

#### Refactored Script and Run Time
![Refactored](https://user-images.githubusercontent.com/115508658/199611412-d84ab09d-7eeb-4102-862e-7216a47d61ad.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/115508658/199611425-44dd1b8d-5864-44ce-9750-28666fd035be.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/115508658/199611430-fb370854-6df9-4d87-b820-14c30d63548b.png)

## Summary

### What are the advantages or disadvantages of refactoring code?

#### Advantages

* Refactoring code could make it more efficient, which allows it to take up less memory and run faster
* Refactoring code could make it easier to read and understand
* Refactoring code could help find bugs

#### Disadvantages

* Refactoring code could introduce new bugs 
* Refactoring code is time consuming, and depending on the code might not save much time to make it worth it

### How do these pros and cons apply to refactoring the original VBA script?

#### Pros
* After refactoring the original script, the code ran faster
* Without the nested loop, the code is easier to read

#### Cons
* At some points, I got stuck on certain refactoring elements which caused errors when running the script
* For a dataset on the smaller side like the green energy stocks, the amount of time it took to refactor was not outweighed by the time saved it took to run
