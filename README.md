# vba-stock-data-analysis
VBA Script that Analyzes Stock Data and Outputs Tables Showing Various Stock Price Performance


This vba subroutine traverses stock data and outputs some key statistics, such as Annual Price Change, Percent Change, and Total Volume.  Then it creates a summary table that shows the best and worst performers by annual price percent change, as well as reports the ticker and volume value for the highest volume stock in each year.

Known Bugs
- On the Summary Statistics table, the script fails to report the best performer by percentage any sheet except the first sheet.  Cause is unknown at this time.

- If a stock starts trading after the first trading day of the year, the script sets the percentage change to zero, because the opening price shows zero.  Instead, the script should find the opening price of the stock when it does start trading and use that opening price as the real opening price.  This is a complication in the way that the script traverses the data.  
