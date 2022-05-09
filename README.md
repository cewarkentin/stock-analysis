# Stock Analysis

## Overview of Project

### Purpose

We were tasked with providing Steve with data analysis on a selection of stocks so that he can appropriately advise his first clients (and also parents) on making informed investments in green energy.

### Methods

Steve provided information for green energy stocks for us to analyze; for each day of data collection, the excel sheet included information on the opening, high, low, and closing prices, as well as the adjusted closing price and volume for each organization.

Using this data, our goal was to find the Total Daily Volume and percent return on investment for each provided green energy stock.

To find the Total Daily Volume for each stock, the volumes for each stock for each day were added together.

To find the percent return, we first determined the closing price for the earliest data point date and the closing price for the most recent data point date. We used the following provided format to calculate the return:

return = (Final Closing Price / Starting Closing Price) - 1

## Results

Given that we are using Excel for these calculations, I would like to continue analysis using a Pivot Table to compare results from 2017 and 2018. More data would give a much more useful picture for larger trends in stock, especially considering the two years that we do have data for tell two different stories (i.e. the majority of the stocks had positive returns in 2017 but negative returns in 2018).

![2017_data](https://github.com/cewarkentin/stock-analysis/blob/main/2017%20data.png)
![2018_data](https://github.com/cewarkentin/stock-analysis/blob/main/2018%20data.png)

It is difficult to advise Steve's parents on investing in green energy based solely on the information provided. However, the only two stocks with positive returns on investment two years in a row belonged to ENPH and RUN. The remainder of the stocks had at least one year with a negative return. 

Between the two stocks with positive returns, RUN had a return that increased over the two-year period and ENPH had a return that decreased over the two-year period.

![VBA_Challenge_2017](https://github.com/cewarkentin/stock-analysis/blob/main/VBA%20Challenge%202017.png)
![VBA_Challenge_2018](https://github.com/cewarkentin/stock-analysis/blob/main/VBA%20Challenge%202018.png)

Run times for the program could have been improved if we had used functions that allowed us to skip over the remaining rows once we knew we had found the final data point for a given stock. For example, if instead of using a for loop to check every single row of data for the DQ ticker, we could go ahead to the next loop once we identified the final DQ data point. However, that does rely on the provided data to be organized in a specific way such that each ticker is grouped together. If the data was organized in such a way, it could improve run-time especially for larger data sets.

## Summary

-there is a detailed statement on the advantages and disadvantages of refactoring code in general

-there is a detailed statement on the advantages and disadvantages of the original and refactored VBA script
